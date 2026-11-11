# -*- coding: utf-8 -*-
"""Client API GoMobile + helpers pour l'automatisation des campagnes.

Ce module ne contient aucune interface Streamlit : il peut être testé seul.
La clé API n'est jamais écrite sur disque ni loguée.
"""

from __future__ import annotations

import re
import time
import unicodedata
from datetime import date, datetime, time as dt_time, timedelta
from io import BytesIO
from pathlib import Path
from typing import Any, Iterable, Optional
from zoneinfo import ZoneInfo

import pandas as pd
import requests

GOMOBILE_API_DEFAULT_BASE = "https://api.gomobile.ma/api"
GOMOBILE_TIMEZONE_NAME = "Africa/Casablanca"
GOMOBILE_TIMEZONE = ZoneInfo(GOMOBILE_TIMEZONE_NAME)

# Noms canoniques des 5 listes BOA demandées.
CANONICAL_LIST_NAMES = (
    "ACTIVATION",
    "ATTRITIONS",
    "CARTE NON ACTIVE",
    "CARTE NON REMISE",
    "WELCOME",
)

FLOW_CHANNEL_LABELS = {
    "call": "Voice",
    "agnostic": "Multicanale",
    "sms": "SMS",
    "whatsapp": "WhatsApp",
}


class GomobileAPIError(Exception):
    """Erreur lisible renvoyée par l'API GoMobile ou par le réseau."""

    def __init__(
        self,
        message: str,
        status_code: Optional[int] = None,
        code: Optional[str] = None,
        payload: Any = None,
    ):
        super().__init__(message)
        self.message = message
        self.status_code = status_code
        self.code = code
        self.payload = payload

    def __str__(self) -> str:
        bits = []
        if self.status_code:
            bits.append(f"HTTP {self.status_code}")
        if self.code:
            bits.append(str(self.code))
        prefix = f"[{' / '.join(bits)}] " if bits else ""
        return f"{prefix}{self.message}"


# ---------------------------------------------------------------------------
# Helpers génériques
# ---------------------------------------------------------------------------

def _first(mapping: dict, *keys: str, default: Any = None) -> Any:
    for key in keys:
        value = mapping.get(key)
        if value not in (None, ""):
            return value
    return default


def count_value(value: Any) -> Optional[int]:
    if value is None:
        return None
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        return int(value)
    if isinstance(value, (list, tuple, set)):
        return len(value)
    if isinstance(value, dict):
        for key in ("count", "total", "value"):
            if key in value:
                return count_value(value[key])
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def extract_item(response: Any) -> dict:
    """Déballe une réponse unitaire ({data:{...}} ou {...})."""
    if isinstance(response, dict):
        data = response.get("data")
        if isinstance(data, dict):
            # Certaines réponses imbriquent encore l'objet.
            for key in ("item", "result"):
                nested = data.get(key)
                if isinstance(nested, dict):
                    return nested
            return data
    return response if isinstance(response, dict) else {}


def extract_list(response: Any, *preferred_keys: str) -> list:
    """Déballe une liste paginée, y compris les réponses imbriquées."""
    if isinstance(response, list):
        return response
    if not isinstance(response, dict):
        return []

    candidates: list[Any] = []
    for key in preferred_keys:
        candidates.append(response.get(key))
    for key in ("data", "items", "results", "rows"):
        candidates.append(response.get(key))

    data = response.get("data")
    if isinstance(data, dict):
        for key in preferred_keys:
            candidates.append(data.get(key))
        for key in ("items", "results", "rows", "list"):
            candidates.append(data.get(key))

    for value in candidates:
        if isinstance(value, list):
            return value
    return []


def extract_detected_columns(response: Any) -> list[dict]:
    """Récupère les colonnes détectées par l'import GoMobile."""
    data = extract_item(response)
    for key in ("detectedColumns", "columns", "headers", "fields", "sourceColumns"):
        value = data.get(key)
        if not isinstance(value, list):
            continue
        out: list[dict] = []
        for index, item in enumerate(value):
            if isinstance(item, dict):
                name = _first(item, "name", "header", "label", "columnName", "column", default="")
                idx = _first(item, "index", "columnIndex", "position", default=index)
                try:
                    idx = int(idx)
                except (TypeError, ValueError):
                    idx = index
                out.append({"index": idx, "name": str(name)})
            elif isinstance(item, str):
                out.append({"index": index, "name": item})
        if out:
            return out
    return []


def extract_identity(me: Any) -> dict:
    """Extrait utilisateur / organisation d'une réponse /users/me."""
    data = extract_item(me)
    user = data.get("user") if isinstance(data.get("user"), dict) else data
    organization = data.get("organization") or user.get("organization") or {}
    if isinstance(organization, list):
        organization = organization[0] if organization else {}
    if not isinstance(organization, dict):
        organization = {}

    full_name = _first(user, "fullName", "name", "displayName", default="")
    if not full_name:
        parts = [
            str(_first(user, "firstName", "firstname", default="") or "").strip(),
            str(_first(user, "lastName", "lastname", default="") or "").strip(),
        ]
        full_name = " ".join(part for part in parts if part).strip()
    email = str(_first(user, "email", "primaryEmail", default="") or "")
    role = str(_first(user, "role", "roleName", default="") or "")
    org_name = str(_first(organization, "name", "displayName", "legalName", default="") or "")
    org_id = str(_first(organization, "id", "organizationId", default="") or "")
    return {
        "name": full_name or email or "Utilisateur",
        "email": email,
        "role": role,
        "organization": org_name,
        "organization_id": org_id,
    }


def normalize_key(value: Any) -> str:
    """Minuscule, sans accents, espaces/ponctuations compactés."""
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = text.upper()
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _strip_trailing_date(stem: str) -> str:
    """Retire un suffixe de date (DDMMYYYY, YYYYMMDD, DD-MM-YYYY...)."""
    text = stem.strip()

    match = re.search(r"(?:^|[\s_\-])(\d{8})$", text)
    if match:
        raw = match.group(1)
        for fmt in ("%d%m%Y", "%Y%m%d"):
            try:
                parsed = datetime.strptime(raw, fmt)
                if 2000 <= parsed.year <= 2100:
                    return text[: match.start(1)].strip(" _-")
            except ValueError:
                pass

    match = re.search(r"(?:^|[\s_\-])(\d{2})[./_\-](\d{2})[./_\-](\d{4})$", text)
    if match:
        day, month, year = match.groups()
        try:
            parsed = datetime(int(year), int(month), int(day))
            if 2000 <= parsed.year <= 2100:
                return text[: match.start()].strip(" _-")
        except ValueError:
            pass

    return text


def clean_list_base_name(filename: str) -> str:
    """Nettoie un nom de fichier pour en faire un nom d'audience."""
    stem = Path(str(filename)).stem
    stem = unicodedata.normalize("NFKD", stem)
    stem = "".join(char for char in stem if not unicodedata.combining(char))
    stem = re.sub(
        r"^(?:BOA_FIC_|BOA_|APT_|CLIENT_FIC_|CLIENT_|FIC_)",
        "",
        stem,
        flags=re.IGNORECASE,
    )
    stem = re.sub(r"[_\-]+", " ", stem)
    stem = re.sub(r"\s+", " ", stem).strip()
    cleaned = _strip_trailing_date(stem)
    return (cleaned or stem).upper().strip()


def canonical_list_name(base_name: str) -> str:
    """Ramène les variantes connues vers les 5 noms BOA officiels."""
    key = normalize_key(base_name)
    if "ACTIVATION" in key:
        return "ACTIVATION"
    if "ATTRITION" in key or "ATTRIBUTION" in key:
        return "ATTRITIONS"
    if "CARTE" in key and "NON" in key and "ACTIVE" in key:
        return "CARTE NON ACTIVE"
    if "CARTE" in key and "NON" in key and "REMISE" in key:
        return "CARTE NON REMISE"
    if "WELCOME" in key:
        return "WELCOME"
    return str(base_name or "").strip().upper()


def audience_name_for_file(filename: str, processing_date: date) -> str:
    """Ex : BOA_FIC_ACTIVATION 09112026.xlsx -> ACTIVATION 10112026."""
    base = canonical_list_name(clean_list_base_name(filename))
    return f"{base} {processing_date.strftime('%d%m%Y')}".strip()


def find_phone_column(headers: Iterable[Any], preferred: str = "TEL_GSM") -> Optional[str]:
    """Trouve la colonne téléphone, priorité à TEL_GSM."""
    header_list = [str(header) for header in headers]
    preferred_key = normalize_key(preferred).replace(" ", "")
    for header in header_list:
        if normalize_key(header).replace(" ", "") == preferred_key:
            return header

    for header in header_list:
        key = normalize_key(header)
        if "GSM" in key:
            return header

    for header in header_list:
        key = normalize_key(header)
        if (
            "TELEPHONE" in key
            or key.startswith("TEL")
            or "PHONE" in key
            or "MOBILE" in key
            or "PORTABLE" in key
        ):
            return header
    return None


def parse_headers_from_bytes(filename: str, data: bytes) -> list[str]:
    """Lit uniquement les en-têtes du fichier (première feuille)."""
    lower = str(filename).lower()
    try:
        if lower.endswith(".csv"):
            for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
                try:
                    frame = pd.read_csv(
                        BytesIO(data),
                        nrows=5,
                        dtype=str,
                        encoding=encoding,
                        sep=None,
                        engine="python",
                    )
                    return [str(column).strip() for column in frame.columns]
                except UnicodeDecodeError:
                    continue
        elif lower.endswith((".xlsx", ".xls")):
            frame = pd.read_excel(BytesIO(data), nrows=0, dtype=str)
            return [str(column).strip() for column in frame.columns]
        elif lower.endswith(".txt"):
            text = None
            for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
                try:
                    text = data[:65536].decode(encoding, errors="ignore")
                    break
                except UnicodeDecodeError:
                    continue
            if text:
                lines = text.splitlines()
                first_line = lines[0] if lines else ""
                if "|;" in first_line:
                    return [
                        part.strip()
                        for part in first_line.split("|;")
                        if part.strip()
                    ]
                for separator in ("\t", ";", "|"):
                    if separator in first_line:
                        return [
                            part.strip()
                            for part in first_line.split(separator)
                            if part.strip()
                        ]
                return [
                    part.strip()
                    for part in first_line.split(",")
                    if part.strip()
                ]
    except Exception:
        return []
    return []


def build_schedule(
    processing_date: date,
    first_time: dt_time,
    second_time: Optional[dt_time],
    third_time: Optional[dt_time],
    now: Optional[datetime] = None,
    skip_past: bool = True,
    shift_past_retries: bool = False,
    retry_spacing_minutes: int = 30,
) -> tuple[datetime, datetime, list[datetime]]:
    """Construit startAt, stopAt et les heures de relance (tentatives 2 et 3).

    GoMobile attend dans `scheduledTimes` uniquement les tentatives 2..N ;
    la tentative 1 est le lancement de la campagne (`startAt`).

    - `skip_past=True` : une relance déjà passée est simplement ignorée.
    - `shift_past_retries=True` : une relance déjà passée est décalée juste
      après l'heure courante afin de conserver le nombre de tentatives
      demandé sur la journée.
    """
    timezone = GOMOBILE_TIMEZONE
    if isinstance(processing_date, datetime):
        processing_date = processing_date.date()
    now = now or datetime.now(timezone)
    if now.tzinfo is None:
        now = now.replace(tzinfo=timezone)

    start_at = datetime.combine(processing_date, first_time, tzinfo=timezone)
    stop_at = datetime.combine(processing_date, dt_time(23, 59, 59), tzinfo=timezone)

    retries: list[datetime] = []
    spacing = timedelta(minutes=max(int(retry_spacing_minutes), 1))
    for candidate in (second_time, third_time):
        if candidate is None:
            continue
        retry_at = datetime.combine(processing_date, candidate, tzinfo=timezone)
        if retry_at <= start_at:
            continue
        if retry_at <= now:
            if shift_past_retries:
                # Relance ratée : on la décale juste après l'heure courante
                # pour conserver le nombre de tentatives demandé.
                retry_at = now + spacing * (len(retries) + 1)
            elif skip_past:
                continue
            # Sinon (skip_past=False) on conserve l'heure d'origine :
            # l'API peut la refuser, c'est l'utilisateur qui l'a demandé.
        retries.append(retry_at)

    # GoMobile exige des horodatages strictement croissants.
    ordered: list[datetime] = []
    for value in retries:
        if ordered and value <= ordered[-1]:
            value = ordered[-1] + spacing
        if value > stop_at:
            break
        ordered.append(value)

    return start_at, stop_at, ordered


def iso_datetime(value: datetime) -> str:
    return value.isoformat(timespec="seconds")


def flow_label(flow: dict) -> str:
    name = str(_first(flow, "name", "flowName", default="Sans nom"))
    channel = str(_first(flow, "channelType", "channel", default="") or "")
    channel_label = FLOW_CHANNEL_LABELS.get(channel, channel)
    status = str(_first(flow, "status", "state", default="") or "")
    parts = [name]
    if channel_label:
        parts.append(channel_label)
    if status:
        parts.append(status)
    return " — ".join(parts)


def did_label(did: dict) -> str:
    if not isinstance(did, dict):
        return str(did)
    number = _first(
        did,
        "number",
        "didNumber",
        "phoneNumber",
        "e164",
        "msisdn",
        "value",
        default="",
    )
    label = _first(
        did,
        "friendlyName",
        "label",
        "name",
        "description",
        default="",
    )
    label = str(label or "")
    number = str(number or "")
    if label and number and label != number:
        return f"{label} — {number}"
    return label or number or str(_first(did, "id", "didId", default="DID"))


def did_id(did: Any) -> str:
    if isinstance(did, dict):
        return str(_first(did, "id", "didId", "uuid", default="") or "")
    return str(did or "")


def build_campaign_payload(
    *,
    name: str,
    description: str,
    flow: dict,
    audience_id: str,
    did_ids: list[str],
    start_at: datetime,
    stop_at: datetime,
    retry_times: list[datetime],
    timeout_ms: int = 30000,
    enable_amd: bool = True,
    record_call: bool = False,
    retry_on_voicemail: bool = False,
) -> dict:
    """Construit le corps POST /campaigns attendu par GoMobile."""
    channel = str(_first(flow, "channelType", "channel", default="") or "")
    payload: dict[str, Any] = {
        "name": name,
        "description": description,
        "flowId": str(_first(flow, "id", "flowId", default="") or ""),
        "audienceId": str(audience_id),
        "startAt": iso_datetime(start_at),
        "stopAt": iso_datetime(stop_at),
    }

    if did_ids and channel in ("call", "agnostic", ""):
        payload["dialConfig"] = {
            "didIds": [str(did) for did in did_ids],
            "timeoutMs": int(timeout_ms),
            "enableAMD": bool(enable_amd),
            "recordCall": bool(record_call),
        }

    if retry_times and channel == "call":
        payload["retryConfig"] = {
            "policy": {
                "type": "scheduled",
                "scheduledTimes": [iso_datetime(retry) for retry in retry_times],
            },
            "retryOnVoicemail": bool(retry_on_voicemail),
        }

    return payload


# ---------------------------------------------------------------------------
# Client HTTP
# ---------------------------------------------------------------------------


def _parse_call_export_bytes(content: bytes) -> list[dict]:
    """Parse un export d'appels GoMobile (CSV ou Excel) en liste de dicts."""
    if not content:
        return []

    # Excel ?
    if content[:2] == b"PK" or content[:4] == b"\xd0\xcf\x11\xe0":
        try:
            df = pd.read_excel(BytesIO(content), dtype=str)
            return df.fillna("").to_dict("records")
        except Exception:
            pass

    # CSV / texte
    text = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            text = content.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = content.decode("utf-8", errors="replace")

    # D?tecter le s?parateur
    sample = "\n".join(text.splitlines()[:5])
    sep = ","
    try:
        import csv
        dialect = csv.Sniffer().sniff(sample, delimiters=[",", ";", "\t", "|"])
        sep = dialect.delimiter
    except Exception:
        if sample.count(";") > sample.count(","):
            sep = ";"

    try:
        df = pd.read_csv(BytesIO(text.encode("utf-8")), sep=sep, dtype=str)
    except Exception:
        try:
            df = pd.read_csv(BytesIO(text.encode("utf-8")), dtype=str)
        except Exception:
            return []
    return df.fillna("").to_dict("records")



def _pick(record: dict, *keys):
    if not isinstance(record, dict):
        return None
    for k in keys:
        v = record.get(k)
        if v not in (None, ""):
            return v
    lower = {str(k).lower().strip(): v for k, v in record.items()}
    for k in keys:
        v = lower.get(str(k).lower().strip())
        if v not in (None, ""):
            return v
    return None


def _find_dtmf_in_record(record: dict) -> str:
    if not isinstance(record, dict):
        return ""
    dtmf_keywords = ['input', 'dtmf', 'digit', 'key', 'touche', 'saisie', 'chiffre', 'reponse']
    for key, value in record.items():
        k = str(key).lower().replace(' ', '').replace('_', '').replace(':', '')
        if any(word in k for word in dtmf_keywords) and 'branch' not in k and value is not None and str(value) != '':
            val = str(value).strip()
            if any(ch.isdigit() for ch in val):
                return val
    return ""


def _normalize_call_record(record: dict) -> dict:
    """Aplati un enregistrement d'appel API pour matcher les champs BOA."""
    if not isinstance(record, dict):
        return {}
    flat = dict(record)

    for key in ("contact", "leg", "call", "details", "report", "attributes", "meta"):
        nested = record.get(key)
        if isinstance(nested, dict):
            for nk, nv in nested.items():
                if nk not in flat or flat[nk] in (None, "", []):
                    flat[nk] = nv

    if not _pick(flat, "phone", "telephone", "tel", "to", "phoneNumber", "phone_number", "destination", "callee"):
        for cand in ("primaryPhone", "primary_phone", "contactPhone", "contact_phone", "number"):
            if flat.get(cand):
                flat["phone"] = flat[cand]
                break
        contact = record.get("contact") if isinstance(record.get("contact"), dict) else {}
        if not flat.get("phone") and contact:
            for cand in ("phone", "primaryPhone", "telephone", "tel"):
                if contact.get(cand):
                    flat["phone"] = contact[cand]
                    break

    outcome = _pick(flat, "call outcome", "call_outcome", "callOutcome", "outcome", "status", "result", "disposition")
    if outcome and "outcome" not in flat:
        flat["outcome"] = outcome
    if outcome and "call outcome" not in flat:
        flat["call outcome"] = outcome

    answered = _pick(flat, "answered by", "answered_by", "answeredBy", "answeredby", "amd", "amdResult", "amd_result")
    if answered and "answered by" not in flat:
        flat["answered by"] = answered
    if not _pick(flat, "answered by", "answered_by", "answeredBy", "answeredby"):
        oc = str(outcome or "").lower()
        reason = str(_pick(flat, "outcomeReason", "outcome_reason", "reason", "hangupCause") or "").lower()
        if oc in ("completed", "answered", "human") or "human" in reason or "contact_hung" in reason:
            flat["answered by"] = "human"
        elif "voicemail" in oc or "machine" in oc or "amd" in oc or "voicemail" in reason or "machine" in reason:
            flat["answered by"] = "machine"

    dur_s = _pick(flat, "duration (s)", "duration", "duration_s", "durationSec", "duration_sec", "billSec", "callDuration")
    dur_ms = _pick(flat, "durationMs", "duration_ms", "durationMillis")
    if dur_ms and not dur_s:
        try:
            flat["duration (s)"] = float(str(dur_ms).replace(",", ".")) / 1000.0
        except Exception:
            pass
    elif dur_s and "duration (s)" not in flat:
        flat["duration (s)"] = dur_s

    dtmf = _pick(flat, "input", "dtmf", "digit", "digits", "key", "touche", "saisie", "chiffre", "reponse", "dtmfInput", "dtmf_input", "userInput")
    if dtmf and "dtmf" not in flat:
        flat["dtmf"] = dtmf
    if dtmf and "input" not in flat:
        flat["input"] = dtmf

    cle = _pick(flat, "cle_contact", "clecontact", "cleContact", "contactId", "contact_id", "externalId")
    if cle and "cle_contact" not in flat:
        flat["cle_contact"] = cle

    return flat


def _merge_call_detail(base: dict, detail: dict) -> dict:
    if not isinstance(detail, dict):
        return base
    merged = dict(base)
    dflat = _normalize_call_record(detail)

    if not _find_dtmf_in_record(merged):
        for container_key in ("timeline", "nodeRuns", "node_runs", "events", "steps", "interactions", "messages", "legs"):
            container = detail.get(container_key)
            if isinstance(container, list):
                for item in container:
                    if isinstance(item, dict):
                        val = _find_dtmf_in_record(_normalize_call_record(item))
                        if val:
                            merged["input"] = val
                            merged["dtmf"] = val
                            break
            if _find_dtmf_in_record(merged):
                break
        if not _find_dtmf_in_record(merged):
            for k, v in list(detail.items()):
                if isinstance(v, dict):
                    val = _find_dtmf_in_record(_normalize_call_record(v))
                    if val:
                        merged["input"] = val
                        merged["dtmf"] = val
                        break

    for src in (dflat, detail):
        if isinstance(src, dict):
            for k, v in src.items():
                if k not in merged or merged[k] in (None, ""):
                    merged[k] = v
    return _normalize_call_record(merged)


class GomobileAPI:
    """Petit client REST pour https://api.gomobile.ma/api."""

    def __init__(
        self,
        api_key: str,
        base_url: str = GOMOBILE_API_DEFAULT_BASE,
        timeout: int = 60,
    ):
        key = (api_key or "").strip()
        if not key:
            raise GomobileAPIError("Clé API vide.")
        base = (base_url or GOMOBILE_API_DEFAULT_BASE).strip().rstrip("/")
        if "/api" not in base:
            base = f"{base}/api"
        self.api_key = key
        self.base_url = base
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update(
            {
                "X-API-Key": key,
                "Accept": "application/json",
                "User-Agent": "BOA-Marketing-GoMobile/1.0",
            }
        )

    # -- HTTP -------------------------------------------------------------
    def request(
        self,
        method: str,
        path: str,
        *,
        params: Optional[dict] = None,
        json_body: Any = None,
        files: Optional[dict] = None,
        data: Optional[dict] = None,
        timeout: Optional[int] = None,
    ) -> Any:
        method = method.upper()
        url = f"{self.base_url}/{str(path).lstrip('/')}"
        request_timeout = timeout or self.timeout
        attempt = 0

        while True:
            try:
                response = self.session.request(
                    method,
                    url,
                    params=params,
                    json=json_body,
                    files=files,
                    data=data,
                    timeout=request_timeout,
                )
            except requests.Timeout as exc:
                if method == "GET" and attempt < 1:
                    attempt += 1
                    continue
                raise GomobileAPIError(
                    "Délai dépassé. Pour une création, vérifiez dans GoMobile "
                    "avant de relancer pour éviter un doublon."
                ) from exc
            except requests.RequestException as exc:
                raise GomobileAPIError(f"Erreur réseau : {exc}") from exc

            if response.status_code == 429 and attempt < 3:
                retry_after = response.headers.get("Retry-After")
                try:
                    wait_seconds = float(retry_after) if retry_after else min(2 ** attempt, 10)
                except (TypeError, ValueError):
                    wait_seconds = min(2 ** attempt, 10)
                time.sleep(max(0.5, min(wait_seconds, 30)))
                attempt += 1
                continue

            if response.status_code >= 500 and method == "GET" and attempt < 2:
                time.sleep(1)
                attempt += 1
                continue

            break

        if 200 <= response.status_code < 300:
            if response.status_code == 204 or not response.content:
                return {}
            try:
                return response.json()
            except ValueError:
                return {"raw": response.text}

        raise self._error_from_response(response)

    @staticmethod
    def _error_from_response(response: requests.Response) -> GomobileAPIError:
        try:
            payload = response.json()
        except ValueError:
            payload = {"message": response.text[:1000]}

        if not isinstance(payload, dict):
            payload = {"message": str(payload)}

        code = payload.get("code")
        message = (
            payload.get("message")
            or payload.get("error")
            or f"Erreur HTTP {response.status_code}"
        )

        field_errors = payload.get("fieldErrors")
        if isinstance(field_errors, dict) and field_errors:
            details = []
            for field, errors in field_errors.items():
                if isinstance(errors, (list, tuple)):
                    details.append(f"{field}: {', '.join(map(str, errors))}")
                else:
                    details.append(f"{field}: {errors}")
            message = f"{message} — {'; '.join(details)}"

        return GomobileAPIError(
            str(message),
            status_code=response.status_code,
            code=str(code) if code else None,
            payload=payload,
        )

    def get(self, path: str, **kwargs) -> Any:
        return self.request("GET", path, **kwargs)

    def post(self, path: str, **kwargs) -> Any:
        return self.request("POST", path, **kwargs)

    def _paginate(
        self,
        path: str,
        *,
        params: Optional[dict] = None,
        limit: int = 100,
        max_pages: int = 50,
    ) -> list:
        items: list = []
        page = 1
        previous_page: Optional[list] = None
        while page <= max_pages:
            query = dict(params or {})
            query.update({"page": page, "limit": limit})
            response = self.get(path, params=query)
            page_items = extract_list(response)
            if page_items == previous_page:
                # Certains endpoints ignorent `page` : on évite de boucler.
                break
            items.extend(page_items)

            meta = response.get("meta", {}) if isinstance(response, dict) else {}
            if not isinstance(meta, dict):
                meta = {}
            has_next = meta.get("hasNextPage")
            total_pages = meta.get("totalPages")
            if has_next is False:
                break
            if not page_items:
                break
            if isinstance(total_pages, int) and total_pages <= 0:
                break
            if isinstance(total_pages, int) and page >= total_pages:
                break
            previous_page = page_items
            page += 1

        # Déduplique les éléments paginés qui exposent un identifiant.
        deduped: list = []
        seen_ids: set[str] = set()
        for item in items:
            identifier = str(item.get("id", "")) if isinstance(item, dict) else ""
            if identifier:
                if identifier in seen_ids:
                    continue
                seen_ids.add(identifier)
            deduped.append(item)
        return deduped

    # -- Connexion / organisation ----------------------------------------
    def me(self) -> Any:
        return self.get("/users/me")

    def validate_connection(self) -> Any:
        """Valide la clé API et renvoie les infos de compte si disponibles.

        Certains environnements GoMobile n'autorisent pas `GET /users/me` avec
        une clé API (403 FORBIDDEN). Dans ce cas on valide l'accès via un
        endpoint de lecture autorisé par la clé (flows, audiences, campagnes).
        """
        try:
            return self.get("/users/me")
        except GomobileAPIError as exc:
            # 401 = clé invalide/expirée/révoquée -> erreur réelle.
            if exc.status_code == 401:
                raise
            # 403/404 sur /users/me : endpoint non accessible avec une clé API.
            # On tente des endpoints de lecture pour confirmer l'authentification.
            if exc.status_code not in (403, 404):
                raise

            probes = (
                "/flows",
                "/audience",
                "/campaigns",
                "/dids/available",
                "/dids",
            )
            last_error: Optional[GomobileAPIError] = exc
            for path in probes:
                try:
                    self.get(path, params={"page": 1, "limit": 1})
                    return {
                        "user": {"fullName": "Clé API"},
                        "_validated_via": path,
                    }
                except GomobileAPIError as probe_exc:
                    last_error = probe_exc
                    # 401 = clé invalide -> remonter immédiatement.
                    if probe_exc.status_code == 401:
                        raise
                    # 403/404 sur ce probe : essayer le suivant (permission/service).
                    if probe_exc.status_code in (403, 404):
                        continue
                    # Autre erreur (5xx, réseau mappé, etc.) : on la conserve
                    # mais on tente quand même les autres probes.
                    continue

            # Aucun probe n'a confirmé l'accès : on remonte l'erreur la plus
            # parlante (la dernière rencontrée).
            raise last_error or exc

    # -- Audiences --------------------------------------------------------
    def list_audiences(self, search: Optional[str] = None) -> list:
        params = {"search": search} if search else None
        return self._paginate("/audience", params=params)

    def find_audiences_by_name(self, name: str) -> list[dict]:
        """Toutes les audiences dont le nom correspond exactement (casse ignorée)."""
        wanted = str(name or "").strip().casefold()
        if not wanted:
            return []
        search = str(name).strip()[:100]
        return [
            audience
            for audience in self.list_audiences(search=search)
            if str(audience.get("name", "")).strip().casefold() == wanted
        ]

    def find_audience_by_name(self, name: str) -> Optional[dict]:
        matches = self.find_audiences_by_name(name)
        return matches[0] if matches else None

    def create_audience(self, name: str, description: str = "") -> dict:
        body = {"name": name, "type": "static"}
        if description:
            body["description"] = description
        return extract_item(self.post("/audience", json_body=body))

    def get_audience(self, audience_id: str) -> dict:
        return extract_item(self.get(f"/audience/{audience_id}"))

    # -- Contacts / ingestion --------------------------------------------
    def upload_ingestion(
        self,
        filename: str,
        file_bytes: bytes,
        audience_id: Optional[str] = None,
    ) -> dict:
        lower = filename.lower()
        if lower.endswith(".csv"):
            mime = "text/csv"
        elif lower.endswith(".xls"):
            mime = "application/vnd.ms-excel"
        else:
            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        form = {"hasHeaderRow": "true", "previewRows": "20"}
        if audience_id:
            form["audienceId"] = str(audience_id)

        response = self.post(
            "/contact/ingestion/upload",
            files={"file": (filename, file_bytes, mime)},
            data=form,
            timeout=180,
        )
        return extract_item(response)

    def validate_ingestion(self, session_id: str, column_mappings: list[dict]) -> dict:
        response = self.post(
            "/contact/ingestion/validate",
            json_body={
                "sessionId": session_id,
                "columnMappings": column_mappings,
            },
            timeout=180,
        )
        return extract_item(response)

    def execute_ingestion(self, session_id: str) -> dict:
        response = self.post(
            "/contact/ingestion/execute",
            json_body={"sessionId": session_id},
            timeout=300,
        )
        return extract_item(response)

    # -- Flows ------------------------------------------------------------
    def list_flows(self, channel_type: Optional[str] = None) -> list:
        params = {"channelType": channel_type} if channel_type else None
        return self._paginate("/flows", params=params)

    # -- DIDs -------------------------------------------------------------
    def list_available_dids(self) -> list:
        """Liste les numéros disponibles. L'endpoint n'est pas public dans
        le catalogue, mais il est utilisé par l'interface GoMobile."""
        for path in ("/dids/available", "/dids"):
            try:
                return self._paginate(path, limit=100)
            except GomobileAPIError as exc:
                if exc.status_code in (403, 404):
                    continue
                raise
        return []

    # -- Campagnes --------------------------------------------------------
    def list_campaigns(self, search: Optional[str] = None) -> list:
        params = {"search": search} if search else None
        return self._paginate("/campaigns", params=params)

    def find_campaigns_by_name(self, name: str) -> list[dict]:
        """Toutes les campagnes dont le nom correspond exactement (casse ignorée)."""
        wanted = str(name or "").strip().casefold()
        if not wanted:
            return []
        search = str(name).strip()[:100]
        return [
            campaign
            for campaign in self.list_campaigns(search=search)
            if str(campaign.get("name", "")).strip().casefold() == wanted
        ]

    def find_campaign_by_name(self, name: str) -> Optional[dict]:
        matches = self.find_campaigns_by_name(name)
        return matches[0] if matches else None

    def create_campaign(self, payload: dict) -> dict:
        return extract_item(self.post("/campaigns", json_body=payload))

    def activate_campaign(self, campaign_id: str) -> dict:
        return extract_item(self.post(f"/campaigns/{campaign_id}/activate"))

    def get_campaign(self, campaign_id: str) -> dict:
        return extract_item(self.get(f"/campaigns/{campaign_id}"))


    # -- Appels / exports -------------------------------------------------
    def list_calls(self, **params) -> list:
        """Liste les appels (legs voix). Param?tres : campaignId, contactId,
        from/to (ISO), outcome, source, page/limit, etc."""
        clean = {k: v for k, v in params.items() if v not in (None, "", [])}
        return self._paginate("/calls", params=clean or None)

    def create_calls_export(self, filters: Optional[dict] = None) -> dict:
        """Cr?e un export CSV asynchrone des appels (m?mes filtres que GET /calls)."""
        body = {k: v for k, v in (filters or {}).items() if v not in (None, "", [])}
        return extract_item(self.post("/calls/export", json_body=body, timeout=120))

    def get_calls_export(self, export_id: str) -> dict:
        """Statut d'un export d'appels (completed ? URL de t?l?chargement)."""
        return extract_item(self.get(f"/calls/exports/{export_id}"))

    def download_calls_export(
        self,
        export_id: str,
        poll_interval: float = 2.0,
        timeout: float = 300.0,
    ) -> bytes:
        """Attend la fin d'un export d'appels et t?l?charge le CSV."""
        deadline = time.time() + max(timeout, 10.0)
        last_status = ""
        while time.time() < deadline:
            info = self.get_calls_export(export_id)
            last_status = str(_first(info, "status", "state", default="") or "").lower()
            if last_status == "completed":
                url = _first(
                    info,
                    "url",
                    "downloadUrl",
                    "downloadURL",
                    "fileUrl",
                    "fileURL",
                    "signedUrl",
                    "signedURL",
                    default="",
                )
                if not url:
                    for key in ("data", "file", "result", "export"):
                        nested = info.get(key)
                        if isinstance(nested, dict):
                            url = _first(
                                nested,
                                "url",
                                "downloadUrl",
                                "downloadURL",
                                "fileUrl",
                                "signedUrl",
                                default="",
                            )
                            if url:
                                break
                if not url:
                    raise GomobileAPIError(
                        "L'export est termin? mais aucune URL de t?l?chargement n'a ?t? renvoy?e.",
                        payload=info,
                    )
                try:
                    resp = self.session.get(str(url), timeout=self.timeout)
                    if resp.status_code in (401, 403):
                        resp = self.session.get(
                            str(url),
                            headers=self._auth_headers(),
                            timeout=self.timeout,
                        )
                    resp.raise_for_status()
                    return resp.content
                except requests.RequestException as exc:
                    raise GomobileAPIError(
                        f"Impossible de t?l?charger l'export d'appels : {exc}"
                    ) from exc
            if last_status in {"failed", "error", "cancelled", "canceled"}:
                raise GomobileAPIError(
                    f"L'export d'appels a ?chou? (statut : {last_status or '?'})",
                    payload=info,
                )
            time.sleep(max(poll_interval, 0.5))

        raise GomobileAPIError(
            f"Export d'appels non termin? apr?s {int(timeout)}s (dernier statut : {last_status or '?'})"
        )

    def fetch_calls_export_records(
        self,
        filters: Optional[dict] = None,
        poll_interval: float = 2.0,
        timeout: float = 300.0,
        ) -> list[dict]:
        """Exporte les appels filtrés et renvoie les lignes sous forme de dicts.

        Stratégie : export CSV asynchrone (équivalent export manuel), puis
        fallback sur la liste d'appels si l'export n'est pas disponible.
        """
        filters = {k: v for k, v in (filters or {}).items() if v not in (None, "", [])}
        errors: list[str] = []

        # 1) Export CSV
        try:
            export = self.create_calls_export(filters)
            direct_id = str(
                _first(export, "id", "exportId", "export_id", "jobId", "job_id", default="")
            )
            direct_url = _first(export, "url", "downloadUrl", "downloadURL", "fileUrl", "signedUrl", default="")
            if direct_id:
                content = self.download_calls_export(direct_id, poll_interval, timeout)
                rows = _parse_call_export_bytes(content)
                if rows:
                    return rows
            elif direct_url:
                try:
                    resp = self.session.get(str(direct_url), timeout=self.timeout)
                    if resp.status_code in (401, 403):
                        resp = self.session.get(str(direct_url), headers=self._auth_headers(), timeout=self.timeout)
                    resp.raise_for_status()
                    rows = _parse_call_export_bytes(resp.content)
                    if rows:
                        return rows
                except Exception as exc:
                    errors.append(f"export direct: {exc}")
            else:
                rows = export if isinstance(export, list) else extract_list(export)
                if rows and isinstance(rows[0], dict):
                    return [_normalize_call_record(r) for r in rows]
                errors.append("export: identifiant/URL manquant")
        except Exception as exc:
            errors.append(f"export: {exc}")

        # 2) Fallback : liste des appels
        try:
            calls = self.list_calls(**filters)
            if calls:
                normalized = [_normalize_call_record(c) for c in calls if isinstance(c, dict)]
                needs_detail = any(
                    not _find_dtmf_in_record(r) and not _pick(r, "answered by", "answered_by", "answeredBy", "answeredby")
                    for r in normalized[:10]
                )
                if needs_detail:
                    enriched = []
                    for r in normalized:
                        cid = str(r.get("id") or r.get("callId") or r.get("call_id") or "")
                        if cid:
                            try:
                                detail = extract_item(self.get(f"/calls/{cid}"))
                                r = _merge_call_detail(r, detail)
                            except Exception:
                                pass
                        enriched.append(r)
                    normalized = enriched
                return normalized
        except Exception as exc:
            errors.append(f"liste appels: {exc}")

        if errors:
            raise GomobileAPIError("Impossible de récupérer les appels — " + " | ".join(errors))
        return []

