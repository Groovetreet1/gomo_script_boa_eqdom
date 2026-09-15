# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import zipfile
import re
import io
from io import BytesIO, StringIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side, Font, Alignment, PatternFill
import os
import json
import math

# ==============================
# CONFIG GLOBALE STREAMLIT
# ==============================
st.set_page_config(
    page_title="Suite Outils Téléphonie / Marketing",
    page_icon="https://media.licdn.com/dms/image/v2/D4E0BAQHd1vQ5srIY4w/company-logo_200_200/B4EZZdUUMCHMAI-/0/1745322330921/gomobile_africa_logo?e=2147483647&v=beta&t=h3LJTeBhImOortH_t5PmBxDMwzEi3vyIelylBx9lKuU",
    layout="wide"
)


# ==============================
# FRONTEND — Style professionnel (Light bleu/blanc)
# ==============================

def inject_css():
    """Injette un CSS global : thème clair professionnel (bleu navy + blanc)."""
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

        :root {
            --navy: #1f4e79;
            --navy-dark: #163a5c;
            --navy-soft: #3a7abe;
            --bg: #f5f7fa;
            --card: #ffffff;
            --line: #e3e8f0;
            --ink: #1a2833;
            --muted: #5b6b7b;
            --accent: #f0a500;
        }

        html, body, [class*="css"] {
            font-family: 'Inter', -apple-system, Segoe UI, Roboto, sans-serif;
            color: var(--ink);
        }
        .stApp {
            background-color: var(--bg);
        }

        /* ---- Topbar / titre ---- */
        .app-hero {
            background: linear-gradient(135deg, var(--navy-dark), var(--navy));
            color: #fff;
            padding: 22px 26px;
            border-radius: 14px;
            margin-bottom: 20px;
            box-shadow: 0 8px 24px rgba(31,78,121,.18);
        }
        .app-hero h1 { margin: 0; font-size: 1.6rem; font-weight: 800; letter-spacing: -.2px; }
        .app-hero p  { margin: 6px 0 0; opacity: .9; font-size: .95rem; }

        /* ---- Sidebar ---- */
        [data-testid="stSidebar"] {
            background: var(--navy-dark);
            border-right: 1px solid rgba(255,255,255,.08);
        }
        [data-testid="stSidebar"] * {
            color: #eaf1f8;
        }
        [data-testid="stSidebar"] [data-testid="stCaptionContainer"] {
            color: #a9bfd4;
        }
        .side-brand {
            text-align: center;
            padding: 14px 8px 10px;
            border-bottom: 1px solid rgba(255,255,255,.12);
            margin-bottom: 12px;
        }
        .side-brand img { width: 52px; border-radius: 10px; }
        .side-brand .name { font-weight: 800; font-size: 1.05rem; color: #fff; letter-spacing: .3px; }
        .side-brand .sub  { font-size: .72rem; color: #a9bfd4; }

        /* Radio menu items (sidebar) */
        [data-testid="stSidebar"] .stRadio label {
            border-radius: 8px;
            padding: 6px 10px;
            transition: background .15s;
        }
        [data-testid="stSidebar"] .stRadio label:hover {
            background: rgba(255,255,255,.08);
        }
        [data-testid="stSidebar"] .stRadio [role="radiogroup"] {
            gap: 4px;
        }

        /* ---- Cards / containers ---- */
        [data-testid="stMetric"] {
            background: var(--card);
            border: 1px solid var(--line);
            border-radius: 12px;
            padding: 14px 16px;
            box-shadow: 0 2px 8px rgba(26,40,51,.05);
        }
        [data-testid="stMetricValue"] { color: var(--navy); font-weight: 800; }
        [data-testid="stMetricLabel"] { color: var(--muted); }

        div[data-testid="stExpander"] {
            background: var(--card);
            border: 1px solid var(--line);
            border-radius: 10px;
        }

        /* ---- Boutons ---- */
        .stButton > button {
            border-radius: 8px;
            font-weight: 600;
            border: 1px solid var(--navy);
        }
        .stButton > button[kind="primary"] {
            background: var(--navy);
            color: #fff;
        }
        .stButton > button[kind="primary"]:hover {
            background: var(--navy-soft);
            border-color: var(--navy-soft);
        }

        /* ---- Inputs / selectbox ---- */
        [data-testid="stFileUploader"] section {
            background: var(--card);
            border: 1px dashed var(--navy-soft);
            border-radius: 10px;
        }
        .stSelectbox [data-baseweb="select"] > div {
            border-radius: 8px;
            border-color: var(--line);
        }

        /* ---- Alertes ---- */
        .stAlert {
            border-radius: 10px;
            border-left-width: 5px;
        }

        /* ---- Titres ---- */
        h1, h2, h3 {
            color: var(--ink);
            font-weight: 700;
            letter-spacing: -.3px;
        }

        /* ---- Expander/selects dans la sidebar ---- */
        [data-testid="stSidebar"] h1,
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3 {
            color: #fff;
        }
        [data-testid="stSidebar"] .stNumberInput input,
        [data-testid="stSidebar"] .stTextInput input {
            background: #ffffff;
            border: 1px solid #cdd6e0;
            color: #3a4a5a;
            border-radius: 7px;
        }
        [data-testid="stSidebar"] .stNumberInput input::placeholder,
        [data-testid="stSidebar"] .stTextInput input::placeholder {
            color: #9aa7b3;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def app_hero(title: str, subtitle: str = ""):
    """Bannière d'en-tête stylisée pour chaque outil."""
    st.markdown(
        f"""
        <div class="app-hero">
            <h1>{title}</h1>
            {f"<p>{subtitle}</p>" if subtitle else ""}
        </div>
        """,
        unsafe_allow_html=True,
    )

# ==============================
# 1) APP EQDOM MARKETING
#    (fichier eqdom_mrkg_app.py)
# ==============================

DATE_EXPORT = datetime.now().strftime("%d%m%Y")  # ex : 21/11/2025 -> "21112025"


def traiter_classeur(uploaded_file):
    # Lecture de TOUTES les feuilles en texte
    xls = pd.read_excel(uploaded_file, sheet_name=None, dtype=str)

    processed_sheets = {}
    stats_rows = []

    for idx, (sheet_name, df) in enumerate(xls.items(), start=1):
        df_original = df.copy()
        total_initial = len(df_original)

        # Nettoyage des noms de colonnes
        df.columns = [c.strip() for c in df.columns]

        # Renommage colonnes
        rename_map = {}
        for col in df.columns:
            cu = col.strip().upper()
            if cu == "ID_CLIENT_INTERNE":
                rename_map[col] = "user_identify"
            elif cu == "CLIENT_TEL_VALIDE":
                rename_map[col] = "telephone"
        df = df.rename(columns=rename_map)

        # Si pas de colonne telephone -> on laisse la feuille comme elle est
        if "telephone" not in df.columns:
            processed_sheets[sheet_name] = df_original
            stats_rows.append({
                "Feuille": sheet_name,
                "Lignes originales": total_initial,
                "Lignes valides": total_initial,
                "Lignes non valides": 0,
                "Sans téléphone": 0,
                "Mauvais préfixe (≠ 05/06/07)": 0,
                "Doublons": 0,
                "Remarque": "Colonne CLIENT_TEL_VALIDE / telephone absente"
            })
            continue

        # Tout en texte
        df = df.astype(str)
        df["telephone"] = df["telephone"].fillna("").astype(str).str.strip()

        # Garde seulement les chiffres
        df["telephone_digits"] = df["telephone"].str.replace(r"\D", "", regex=True)

        # Option : transformer 2126xxxxxxx -> 06xxxxxxx
        def normaliser_212(x):
            if x.startswith("212") and len(x) >= 11:
                return "0" + x[3:]
            return x

        df["telephone_norm"] = df["telephone_digits"].apply(normaliser_212)

        # 3) Supprimer lignes avec téléphone vide
        mask_non_vide = df["telephone_norm"] != ""
        df_non_vide = df[mask_non_vide].copy()
        suppr_tel_vide = total_initial - len(df_non_vide)

        # 4) Garder préfixes 05 / 06 / 07
        mask_prefix = df_non_vide["telephone_norm"].str.startswith(("05", "06", "07"))
        df_prefix = df_non_vide[mask_prefix].copy()
        suppr_mauvais_prefixe = len(df_non_vide) - len(df_prefix)

        # 5) Supprimer doublons
        avant_doublons = len(df_prefix)
        df_final = df_prefix.drop_duplicates(subset="telephone_norm", keep="first").copy()
        suppr_doublons = avant_doublons - len(df_final)

        lignes_valides = len(df_final)
        lignes_invalides = total_initial - lignes_valides

        # On enlève les colonnes techniques
        df_final = df_final.drop(columns=["telephone_digits", "telephone_norm"])

        processed_sheets[sheet_name] = df_final

        stats_rows.append({
            "Feuille": sheet_name,
            "Lignes originales": total_initial,
            "Lignes valides": lignes_valides,
            "Lignes non valides": lignes_invalides,
            "Sans téléphone": suppr_tel_vide,
            "Mauvais préfixe (≠ 05/06/07)": suppr_mauvais_prefixe,
            "Doublons": suppr_doublons,
            "Remarque": ""
        })

    # Stats dans un DataFrame
    stats_df = pd.DataFrame(stats_rows)

    # Ligne TOTAL
    if not stats_df.empty:
        total_row = {
            "Feuille": "TOTAL",
            "Lignes originales": stats_df["Lignes originales"].sum(),
            "Lignes valides": stats_df["Lignes valides"].sum(),
            "Lignes non valides": stats_df["Lignes non valides"].sum(),
            "Sans téléphone": stats_df["Sans téléphone"].sum(),
            "Mauvais préfixe (≠ 05/06/07)": stats_df["Mauvais préfixe (≠ 05/06/07)"].sum(),
            "Doublons": stats_df["Doublons"].sum(),
            "Remarque": ""
        }
        stats_df = pd.concat([stats_df, pd.DataFrame([total_row])], ignore_index=True)

    return processed_sheets, stats_df


def build_combined_excel(processed_sheets: dict) -> bytes:
    output = BytesIO()
    import xlsxwriter  # s'assure que le moteur est dispo
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, df in processed_sheets.items():
            safe_name = str(sheet_name)[:31]
            df.to_excel(writer, index=False, sheet_name=safe_name)
    output.seek(0)
    return output.getvalue()


def build_zip(processed_sheets: dict) -> bytes:
    zip_buffer = BytesIO()
    import xlsxwriter

    with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for sheet_name, (sheet_name, df) in enumerate(processed_sheets.items(), start=1):
            # Création d'un fichier Excel en mémoire pour cette feuille
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
                safe_name = str(sheet_name)[:31]
                df.to_excel(writer, index=False, sheet_name=safe_name)
            excel_buffer.seek(0)

            # Nouveau format de nom : EQDOM_RECOUV_FIC_{INDEX}_{DATE}.xlsx
            file_name = f"EQDOM_RECOUV_FIC_{sheet_name}_{DATE_EXPORT}.xlsx"
            zf.writestr(file_name, excel_buffer.getvalue())

    zip_buffer.seek(0)
    return zip_buffer.getvalue()


def app_eqdom_marketing():
    st.title("✅ Traitement EQDOM - MARKETING")

    st.write("""
    ✅ Pour chaque feuille du fichier :
    - Colonnes en **texte (plain text)**  
    - `ID_CLIENT_INTERNE` → **user_identify**  
    - `CLIENT_TEL_VALIDE` → **telephone**  
    - Suppression des lignes sans téléphone  
    - Filtre numéros marocains **05 / 06 / 07**  
    - Suppression des doublons de téléphone  
    - Statistiques par feuille + TOTAL  

    📤 Export :
    - 1 fichier global : **EQDOM_RECOUV_FIC_SEGEMENT.xlsx**  
    - 1 fichier ZIP avec :  
      **EQDOM_RECOUV_FIC_SEGMENT1_21112025.xlsx**,  
      **EQDOM_RECOUV_FIC_SEGMENT2_21112025.xlsx**, etc.
    """)

    uploaded_file = st.file_uploader(
        "Choisir le fichier Excel (.xlsx ou .xls)",
        type=["xlsx", "xls"],
        key="eqdom_file"
    )

    if uploaded_file is not None:
        st.success(f"Fichier chargé : {uploaded_file.name}")

        if st.button("🚀 Lancer le traitement sur toutes les feuilles", key="eqdom_run"):
            # Traitement
            processed_sheets, stats_df = traiter_classeur(uploaded_file)

            # Stats
            st.subheader("📊 Statistiques par feuille")
            st.dataframe(stats_df, width='stretch')

            # Aperçu d'une feuille traitée
            st.subheader("👀 Aperçu d’une feuille traitée")
            feuille_choice = st.selectbox(
                "Choisir une feuille :",
                list(processed_sheets.keys()),
                key="eqdom_sheet_choice"
            )
            st.dataframe(processed_sheets[feuille_choice].head(), width='stretch')

            # Génération des fichiers
            combined_excel_bytes = build_combined_excel(processed_sheets)
            zip_bytes = build_zip(processed_sheets)

            st.subheader("📥 Téléchargements")

            # 1) Fichier global
            st.download_button(
                label="Télécharger le fichier global",
                data=combined_excel_bytes,
                file_name="EQDOM_RECOUV_FIC_SEGEMENT.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="eqdom_download_global"
            )

            # 2) ZIP avec 1 fichier par segment / feuille
            st.download_button(
                label="Télécharger le ZIP (1 fichier par segment)",
                data=zip_bytes,
                file_name="EQDOM_RECOUV_FIC_SEGMENTS.zip",
                mime="application/zip",
                key="eqdom_download_zip"
            )


# ==============================
# 2) APP AVT → APT (GoMobile)
#    (fichier avt2apt_app.py)
# ==============================

ALLOWED_PREFIXES = ("05", "06", "07", "08")


def normalize_phone(val: str) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    if s.upper() in ("", "NAN", "NA", "NONE"):
        return ""
    s = re.sub(r"[^\d\+]", "", s)
    if s.startswith("+212"):
        s = "0" + s[4:]
    elif s.startswith("212"):
        s = "0" + s[3:]
    if s.startswith("+"):
        s = s[1:]
    if s.startswith("00212"):
        s = "0" + s[5:]
    s = re.sub(r"\D", "", s)
    if s.startswith("212"):
        s = "0" + s[3:]
    return s


def is_valid_ma_number(s: str) -> bool:
    return bool(s) and len(s) == 10 and s.startswith(ALLOWED_PREFIXES)


def read_avt(file) -> pd.DataFrame:
    # Format attendu: CLE_CONTACT|;TEL_DOM|;TEL_PRO|;TEL_GSM|;CAMPAGNE
    df = pd.read_csv(file, sep=r"\|\;", engine="python", dtype=str)
    df.columns = [c.strip().replace("\ufeff", "") for c in df.columns]
    for c in df.columns:
        df[c] = df[c].astype(str).str.strip()
    return df


def apply_business_rules(df_in: pd.DataFrame):
    """Applique normalisation + remplissage GSM + NA DOM/PRO. Renvoie df rempli + masque validité GSM Maroc."""
    cols = ["CLE_CONTACT", "TEL_DOM", "TEL_PRO", "TEL_GSM", "CAMPAGNE"]
    df = df_in.reindex(columns=cols).copy()

    for c in ["TEL_DOM", "TEL_PRO", "TEL_GSM"]:
        df[c] = df[c].apply(normalize_phone)

    def fill_gsm(row):
        if not row["TEL_GSM"]:
            if row["TEL_DOM"]:
                return row["TEL_DOM"]
            if row["TEL_PRO"]:
                return row["TEL_PRO"]
        return row["TEL_GSM"]

    df["TEL_GSM"] = df.apply(fill_gsm, axis=1)

    for c in ["TEL_DOM", "TEL_PRO"]:
        df[c] = df[c].apply(lambda x: x if x else "NA")

    for c in cols:
        df[c] = df[c].astype(str)

    valid_mask = df["TEL_GSM"].apply(is_valid_ma_number)
    return df, valid_mask


def df_to_styled_excel_bytes(df: pd.DataFrame) -> io.BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "DATA"
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(row=r, column=c).border = border
    # ajustement des largeurs
    for col in ws.columns:
        max_len = 10
        col_letter = col[0].column_letter
        for cell in col:
            max_len = max(max_len, len(str(cell.value)) if cell.value is not None else 0)
        ws.column_dimensions[col_letter].width = min(max_len + 2, 40)
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


def app_avt_to_apt():
    st.title("📞 AVT → APT – Nettoyage & Normalisation (GoMobile)")
    st.caption(
        "Dépose tes fichiers AVT (.txt) et récupère des .xlsx prêts : "
        "normalisation, remplissage GSM, filtres Maroc, déduplication, "
        "listing des invalides, export cadrillé, et nommage CLIENT_FIC_*."
    )

    uploaded_files = st.file_uploader(
        "Glisse-dépose un ou plusieurs fichiers AVT (.txt) au format note",
        type=["txt"],
        accept_multiple_files=True,
        key="avt_files"
    )

    if uploaded_files:
        zip_ok = io.BytesIO()
        zip_bad = io.BytesIO()

        with zipfile.ZipFile(zip_ok, "w", zipfile.ZIP_DEFLATED) as z_ok, \
             zipfile.ZipFile(zip_bad, "w", zipfile.ZIP_DEFLATED) as z_bad:

            for f in uploaded_files:
                try:
                    df_in = read_avt(f)
                    rows_in = len(df_in)
                    gsm_empty_raw = int((df_in["TEL_GSM"].astype(str).str.strip() == "").sum())

                    df_filled, valid_mask = apply_business_rules(df_in)
                    df_valid = df_filled[valid_mask].copy()
                    df_invalid = df_filled[~valid_mask].copy()

                    # Déduplication TEL_GSM (garder la 1ère occurrence)
                    before = len(df_valid)
                    df_valid = df_valid.drop_duplicates(subset=["TEL_GSM"], keep="first").copy()
                    dup_removed = int(before - len(df_valid))

                    rows_out = len(df_valid)
                    dropped_invalid = len(df_invalid)

                    # Nom de sortie : CLIENT_FIC_<base>.xlsx
                    base_name = f.name.replace("( AVT )", "").replace(".txt", "").strip()
                    out_name_ok = f"BOA_FIC_{base_name}.xlsx"
                    out_name_bad = f"BOA_FIC_{base_name}_supprimes.xlsx"

                    # ZIP des valides
                    z_ok.writestr(out_name_ok, df_to_styled_excel_bytes(df_valid).read())
                    # ZIP des invalides (si présents)
                    if not df_invalid.empty:
                        z_bad.writestr(out_name_bad, df_to_styled_excel_bytes(df_invalid).read())

                    # UI par fichier
                    st.subheader(f"🗂️ {f.name}")
                    st.markdown(f"- Lignes en entrée : **{rows_in}**")
                    st.markdown(f"- `TEL_GSM` vides avant remplissage : **{gsm_empty_raw}**")
                    st.markdown(f"- Doublons GSM supprimés : **{dup_removed}**")
                    st.markdown(f"- Lignes valides conservées : **{rows_out}**")
                    st.markdown(f"- Lignes invalides supprimées : **{dropped_invalid}**")

                    st.markdown("**Aperçu (valides)**")
                    st.dataframe(df_valid.head(10))
                    st.download_button(
                        "⬇️ Télécharger cet Excel (valides)",
                        data=df_to_styled_excel_bytes(df_valid),
                        file_name=out_name_ok,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"ok_{f.name}"
                    )

                    if not df_invalid.empty:
                        st.markdown("🚫 **Numéros supprimés (invalides)** – aperçu")
                        st.dataframe(
                            df_invalid[["CLE_CONTACT", "TEL_DOM", "TEL_PRO", "TEL_GSM"]].head(20)
                        )
                        st.download_button(
                            "⬇️ Télécharger la liste complète des supprimés",
                            data=df_to_styled_excel_bytes(df_invalid),
                            file_name=out_name_bad,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"bad_{f.name}"
                        )
                    else:
                        st.success("✅ Aucun numéro supprimé pour ce fichier.")

                except Exception as e:
                    st.error(f"Erreur sur {f.name} : {e}")

        # Téléchargements groupés
        zip_ok.seek(0)
        zip_bad.seek(0)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.markdown("---")
        st.subheader("📦 Téléchargements groupés")
        st.download_button(
            "⬇️ ZIP – tous les Excel valides",
            data=zip_ok,
            file_name=f"APT_OK_{ts}.zip",
            mime="application/zip",
            key="avt_zip_ok"
        )
        st.download_button(
            "⬇️ ZIP – toutes les listes supprimées",
            data=zip_bad,
            file_name=f"APT_SUPPRIMES_{ts}.zip",
            mime="application/zip",
            key="avt_zip_bad"
        )

    st.markdown("---")
    st.caption(
        "Règles : remplir TEL_GSM depuis DOM/PRO si vide, DOM/PRO vides → 'NA', "
        "normaliser (+212/212→0), garder 10 chiffres démarrant par 05/06/07/08, "
        "supprimer les doublons GSM, lister les invalides. Noms: CLIENT_FIC_<base>.xlsx"
    )



# ==============================
# 3) APP TRAITEMENT AGENCE (Assurcall)
#    (fichier app_traitement_agence_v10.py - intégré)
# ==============================

MAPPING_AGENCES_DEFAULT = {
    "AGENT_A6666": "SRAGHNA",
    "AGENT_A6481": "LE COMPTOIR D ASSURANCES",
    "AGENT_A5836": "AVENIR DU SAHARA",
    "AGENT_A6662": "SOLUDO",
    "AGENT_A6541": "ATTAOUKIL",
    "AGENT_A6581": "TALIOUINE",
    "AGENT_A6548": "LAKHMIRI ASSURANCES SARL",
    "AGENT_A6594": "BOUTADGHART",
    "AGENT_A6597": "MOKRIM",
    "AGENT_A5886": "PACKASSUR",
    "AGENT_A6657": "BELBEKKAR ASSURANCES",
    "AGENT_A6586": "GHANANE",
    "AGENT_A6630": "ASSURANCE JANTI",
    "AGENT_A6634": "ASSURANCES RAMI",
    "AGENT_A4409": "ASSURANCES ELBAHIA",
    "AGENT_A5882": "ASSURANCE CHARAF",
    "AGENT_A5970": "IZIKI ASSUR",
    "AGENT_A6444": "ERRAFAY ASSURANCE CONSEIL SARL",
    "AGENT_A6452": "AGENT ASSIF ASSURANCES CONSEIL",
    "AGENT_A6472": "CHAKROUNI RACHAD ASSURANCES",
    "AGENT_A6496": "BOUAGLOU ASSURANCES SARL",
    "AGENT_A6513": "ASSURANCES OUED SOUSS",
    "AGENT_A6525": "ASSURANCE LM MOURAD LOTFI & FILS",
    "AGENT_A6575": "ASSURANCES ELKIRAM",
    "AGENT_A6576": "ASSURANCE BOUZAGHRANE",
    "AGENT_A6587": "EL BOUHTOURI ASSURANCE",
    "AGENT_A6648": "HAFIDA ASSURANCES",
    "AGENT_A6533": "ASSURANCES LES PORTES DE MARRAKECH SARL",
    "AGENT_A6599": "ASSURANCES ANOUAL MARRAKECH",
    "AGENT_A5998": "TARIKALKHEIR",
    "AGENT_A6536": "DYNAMIC ASSURANCESS SARL",
    "AGENT_A6638": "ASSURANCE BELAKRY",
    "AGENT_A6538": "MAYAR ASSURANCES SARL",
    "AGENT_A6633": "YARASSURANCE",
    "AGENT_A6639": "SAWAB ASSURANCE",
    "AGENT_A5960": "ASSURANCE CONSEIL BOUHNIK",
    "AGENT_A6655": "AIT MELLOUL",
    "AGENT_A6563": "AATIKFILS ASSURANCES SARL",
    "AGENT_A6647": "FAMILY ASSURANCE SARL",
    "AGENT_A6653": "EL BOUHTOURI BAB DOUKKALA",
    "BGD_B8057": "BGD AKIOD",
    "AGENT_A6656": "ASSUNOR",
    "AGENT_A6661": "DIFAF ASSURANCE",
    "AGENT_A6665": "ASSURANCES GHOUDAIGUI",
    "AGENT_A6609": "EL AZZOUZIA ASSURANCE",
    "AGENT_A6623": "ASSURANCE BOUHMAD",
    "BGD_B7778": "BGD FES SALAMA",
    "BGD_B7790": "BGD MISSOUR",
    "BGD_B7791": "ZAIO",
    "BGD_B7816": "AIT KAMRA",
    "BGD_B7774": "BGD SEFROU",
    "BGD_B7780": "TAROUDANT",
    "BGD_B7779": "TIZNIT",
    "BGD_B7792": "BGD CHAOUEN",
    "BGD_B7775": "BGD KHEMISSET",
    "BGD_B7803": "BGD Ras ElMa",
    "BGD_B7866": "BGD AUTOHALL OUJDA TRAFISS",
    "BGD_B7781": "BGD OUARZAZATE",
    "BGD_B7952": "BGD DEROUA",
    "BGD_B7821": "BGD TINGHIR",
    "BGD_B7947": "BGD SALAMA KSAR EL KEBIR",
    "BGD_B7969": "BGD KIA MARRAKECH",
    "BGD_B7948": "BGD KIA CASA",
    "BGD_B7807": "SALE KARIAT",
    "BGD_B8025": "BGD BUGSHAN",
    "AGENT_A6524": "ASSURANCES EL GUARMAI SARL",
    "AGENT_A6582": "REDASSUR",
    "AGENT_A6441": "Vecteur",
    "AGENT_A6445": "IFRY",
    "AGENT_A5830": "ASSURANCES GIRS",
    "AGENT_A6455": "ASSURANCE SOUFIANE SARL",
    "AGENT_A6489": "ASSURANCES ASSIHAM",
    "AGENT_A5887": "ASSURANCES LA FRATERNITE",
    "AGENT_A6456": "ASSURANCES CASA EST SARL",
    "AGENT_A6462": "ASSURANCE RIAH",
    "AGENT_A6593": "OLKOM ASSURANCES",
    "AGENT_A6649": "ASSURANCE PERVALIS",
    "AGENT_A6566": "NOURASSUR SARL",
    "AGENT_A6492": "SIDI BENNOUR",
    "AGENT_A6577": "STE NAHJASSUR SARL",
    "AGENT_A6551": "ALHIKMA ASSURANCE SARL",
    "AGENT_A6651": "ASSURANCES OULAD FREJ",
    "AGENT_A6637": "UNION ASSUR",
    "AGENT_A5878": "Chariba",
    "AGENT_A6522": "STE SIDRA ASSURANCE",
    "AGENT_A5864": "CONFORT ASSURANCES MIDELT",
    "AGENT_A5837": "ASSURANCES GUERRAB",
    "AGENT_A5841": "ASSURANCES HMAMI",
    "AGENT_A5846": "ASSURANCES TALLA",
    "AGENT_A5900": "ASSURANCES FILALI",
    "AGENT_A5980": "ASSURANCES CONSEIL TAOUFIK",
    "AGENT_A5996": "ASSURANCES FOUGHAL SARL",
    "AGENT_A6190": "ASSURANCES BENCHEMSI SAIAR",
    "AGENT_A6442": "REFLEX ASSURANCE SARL",
    "AGENT_A6465": "JAADA ET CHBILI SARL",
    "AGENT_A6561": "ASSURANCES AL HOCEIMA",
    "AGENT_A6588": "ASSURANCE OUSSOU",
    "AGENT_A6589": "PROXY Assurance",
    "AGENT_A6592": "STE ASSURANCES ZAAF",
    "AGENT_A6622": "ASSURANCES ZEGHARI",
    "AGENT_A6583": "ELGHAZAOUATE ASSURANCES",
    "AGENT_A6451": "OASIS ASSURANCES",
    "AGENT_A5863": "OUADIE ASSURANCES",
    "AGENT_A6484": "ASSURANCES OUEDZA TAOURIRT",
    "AGENT_A6619": "ASSURANCES AL YOUSSRE",
    "AGENT_A5770": "ASSURANCES MTALSI",
    "AGENT_A6614": "RANAT ASSURANCE-CONSEIL",
    "AGENT_A6646": "ZAERS",
    "AGENT_A6530": "SAAD ASSUR SARL",
    "AGENT_A6448": "ASSURANCES AL CHAMAL",
    "AGENT_A6571": "ESSBAIY MHAMMED",
    "AGENT_A5260": "ASSURANCES MAHRAZ SARL",
    "AGENT_A5850": "ASSURANCES WATANI",
    "AGENT_A5875": "ASSURANCES LA CAPITALE SARL",
    "AGENT_A5876": "ASSURANCE ZAID",
    "AGENT_A6170": "ASSURANCES BENABDELLAH",
    "AGENT_A6340": "ASSURANCES AFIAC REGRAGUI",
    "AGENT_A6380": "TANGER ASSURANCES CONSEIL",
    "AGENT_A6454": "BALAFREJ ASSURANCE",
    "AGENT_A6483": "Moderne ASSURANCE",
    "AGENT_A6519": "ASSURANCES ZOUHAIR SARL",
    "AGENT_A6537": "ASSURANCES ALBOUGHAZ SARL",
    "AGENT_A6590": "ASSURANCE AJZENNAI OURIAGHLI MBM",
    "AGENT_A6627": "CONECTE ASSURANCE",
    "AGENT_A6620": "GROUPE ASSURANCES MOSTAINE",
    "AGENT_A6626": "ASSURANCES ZOUMI",
    "AGENT_A5842": "ASSURANCES BENNOUNA",
    "AGENT_A6570": "TAAMINAT EDDAMINE",
    "AGENT_A6250": "ASSURANCES ELMARKAZ",
    "AGENT_A6628": "ASSURANCE DAR DMANA",
    "BGD_B8026": "Aioun Sidi Mellouk",
    "BGD_B8019": "BGD DAR OULAD ZIDOUH",
    "BGD_B8023": "BGD ZEMAMRA",
    "BGD_B8049": "BGD Ain Sebaa",
    "BGD_B7987": "TAHANAOUT",
    "BGD_B7080": "BGD RABAT",
    "BGD_B8050": "BGD IMINTANOUTE",
    "BGD_B7800": "ER RICH",
    "BGD_B8051": "BGD TANTAN",
    "BGD_B7990": "BIR JDID",
    "BGD_B8039": "BGD OURIKA",
    "BGD_B7977": "AMIZMIZ",
    "BGD_B8016": "BOUGUEDRA",
    "BGD_B8005": "CHICHAOUA",
    "BGD_B7872": "BGD Ferkhana",
    "BGD_B7899": "BGD INEZGANE",
    "BGD_B7086": "BGD AGADIR 1",
    "BGD_B7862": "BGD SIDI KACEM",
    "BGD_B8047": "BGD TASSOULTANTE",
    "BGD_B7982": "BGD HAD SOUALEM",
    "BGD_B8030": "BGD BOUJDOUR",
    "BGD_B7956": "BGD ELHAJEB",
    "BGD_B8048": "BGD AOULOUZ",
    "BGD_B8031": "BGD TINEJDAD",
    "BGD_B8018": "BGD TADLA",
    "BGD_B8010": "BGD DRARGA",
    "BGD_B7989": "BGD TAMANSOURTE",
    "BGD_B8028": "BGD LOUDAYA",
    "BGD_B8006": "BGD LAAOUNATE",
    "BGD_B8015": "BGD AOURIR",
    "BGD_B7925": "SIDIBIBI",
    "BGD_B7787": "BGD MARRAKECH HIVERNAGE",
    "BGD_B7092": "BGD CASA",
    "BGD_B7074": "BGD OUJDA BELHAJ",
    "BGD_B8027": "BGD SIDI BOUKNADEL",
    "BGD_B7084": "BGD FES",
    "BGD_B8024": "BGD OULED BERHIL",
    "BGD_B8040": "BGD TIKIOUINE",
    "BGD_B7890": "BGD Echemmaia",
    "BGD_B7986": "BGD LQLIAA",
    "BGD_B8021": "BGD AIN TAOUJDATE",
    "BGD_B8008": "Zaida",
    "BGD_B8022": "BGD BEN AHMED",
    "BGD_B8029": "BGD EL BOROUJ",
    "BGD_B7918": "BGD BIOUGRA",
    "AGENT_A6683": "Agence Luxora assurance Ain HARROUDA",
    "BGD_B7820": "BGD FQUIH BEN SALEH",
    "BGD_B7991": "BGD DEROUA CENTRE",
    "BGD_B7937": "BGD El GARA",
    "BGD_B7978": "AIT AMIRA",
    "BGD_B7985": "ARFOUD",
    "BGD_B7992": "BGD AGADIR HAY SALAM",
    "BGD_B7988": "BGD BERKANE",
    "BGD_B8007": "Mediouna",
    "BGD_B8009": "Oulad Ayad",
    "BGD_B7993": "TAZA",
    "BGD_B8012": "DEMNATE",
    "BGD_B8014": "BGD MOHAMMEDIA",
    "BGD_B7979": "DAKHLA",
    "BGD_B8011": "BRADIA",
    "AGENT_A6607": "ASSURANCES AL MADAR",
    "AGENT_A6608": "SEBKI ASSURANCES",
    "AGENT_A6615": "COIN DE L ASSURANCE",
    "AGENT_A6625": "ASSURANCES TALMESTE",
    "AGENT_A5180": "ASSUR-ALWATA-ESS",
    "AGENT_A6446": "JNK",
    "BGD_B7955": "BGD Selouane",
    "AGENT_A6668": "WEST MED ASSURANCE",
    "AGENT_A6684": "Assurances Unies",
    "AGENT_A5920": "RIAD ASSURANCES",
    "BGD_B8013": "BGD BELFAA",
    "AGENT_A6549": "ASSURANCES BOUZID"
}

# Fichier persistant pour le mapping (JSON a cote du script)
MAPPING_AGENCES_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd(), "mapping_agences.json")

def load_mapping_agences_from_file():
    """Charge le mapping : si fichier JSON existe → retourne le fichier tel quel (respecte Remplacer tout). Sinon défaut."""
    try:
        if os.path.exists(MAPPING_AGENCES_FILE):
            with open(MAPPING_AGENCES_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict) and data:
                    return data
    except Exception as e:
        print(f"load mapping error: {e}")
    return MAPPING_AGENCES_DEFAULT.copy()

def save_mapping_agences_to_file(mapping):
    """Sauvegarde le mapping dans le fichier JSON persistant."""
    try:
        with open(MAPPING_AGENCES_FILE, "w", encoding="utf-8") as f:
            json.dump(mapping, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"save mapping error: {e}")
        return False

def get_mapping_agences():
    """Retourne le mapping courant (session_state > fichier > defaut)."""
    try:
        if "mapping_agences" in st.session_state:
            return st.session_state["mapping_agences"]
    except:
        pass
    mapping = load_mapping_agences_from_file()
    try:
        st.session_state["mapping_agences"] = mapping
    except:
        pass
    return mapping

def get_mapping_numero_vers_code():
    mapping = get_mapping_agences()
    d = {}
    for code in mapping.keys():
        mm = re.search(r"(\d+)$", code)
        if mm:
            d[mm.group(1)] = code
    return d

def get_liste_agences_dropdown():
    mapping = get_mapping_agences()
    return [""] + sorted([f"{nom} ({code})" for code, nom in mapping.items()], key=lambda x: x.lower())

# Compatibilite : garde les anciens noms mais dynamiques au chargement initial
try:
    _initial = load_mapping_agences_from_file()
    MAPPING_AGENCES = _initial
    MAPPING_NUMERO_VERS_CODE = {}
    for _code in MAPPING_AGENCES.keys():
        _match = re.search(r"(\d+)$", _code)
        if _match:
            MAPPING_NUMERO_VERS_CODE[_match.group(1)] = _code
    LISTE_AGENCES_DROPDOWN = [""] + sorted([f"{nom} ({code})" for code, nom in MAPPING_AGENCES.items()], key=lambda x: x.lower())
except:
    MAPPING_AGENCES = MAPPING_AGENCES_DEFAULT.copy()
    MAPPING_NUMERO_VERS_CODE = {}
    for _code in MAPPING_AGENCES.keys():
        _match = re.search(r"(\d+)$", _code)
        if _match:
            MAPPING_NUMERO_VERS_CODE[_match.group(1)] = _code
    LISTE_AGENCES_DROPDOWN = [""] + sorted([f"{nom} ({code})" for code, nom in MAPPING_AGENCES.items()], key=lambda x: x.lower())

def _sync_mapping_globals(mapping):
    """Met a jour les globals pour compatibilite apres modification."""
    global MAPPING_AGENCES, MAPPING_NUMERO_VERS_CODE, LISTE_AGENCES_DROPDOWN
    MAPPING_AGENCES = mapping
    MAPPING_NUMERO_VERS_CODE = {}
    for _code in mapping.keys():
        _match = re.search(r"(\d+)$", _code)
        if _match:
            MAPPING_NUMERO_VERS_CODE[_match.group(1)] = _code
    LISTE_AGENCES_DROPDOWN = [""] + sorted([f"{nom} ({code})" for code, nom in mapping.items()], key=lambda x: x.lower())
    try:
        st.session_state["mapping_agences"] = mapping
    except:
        pass
    save_mapping_agences_to_file(mapping)

MODELE_COLONNES = [
    "N° Police", "Intermediaire", "CIN", "Nom/Raison sociale", "Usage",
    "Immatriculation", "Duree", "Date Effet", "Date Echeance", "Heure Echeance",
    "Segment", "Risque départ", "Appétence MRH", "Etat", "Telephone"
]

COLONNES_VIDES = ["Segment", "Risque départ", "Appétence MRH"]

ALIASES_COLONNES = {
    "numero_police": "N° Police", "numéro_police": "N° Police", "num police": "N° Police",
    "police": "N° Police", "n police": "N° Police", "npolice": "N° Police",
    "n° police": "N° Police", "n°police": "N° Police",
    "interm": "Intermediaire", "intermediaire": "Intermediaire", "agence": "Intermediaire",
    "cin": "CIN", "nom": "Nom/Raison sociale", "nom_client": "Nom/Raison sociale",
    "nom client": "Nom/Raison sociale", "raison sociale": "Nom/Raison sociale",
    "nomraison sociale": "Nom/Raison sociale",
    "client": "Nom/Raison sociale", "usage": "Usage", "designation": "Usage",
    "immatriculation": "Immatriculation", "immat": "Immatriculation", "matricule": "Immatriculation", "matricule vehicule": "Immatriculation", "matriculevehicule": "Immatriculation", "plaque": "Immatriculation", "plaque immatriculation": "Immatriculation",
    "duree": "Duree", "durée": "Duree", "date_effet": "Date Effet",
    "date effet": "Date Effet", "date_echeance": "Date Echeance",
    "date echeance": "Date Echeance", "echeance": "Date Echeance",
    "heure echeance": "Heure Echeance", "heure_echeance": "Heure Echeance",
    "segment": "Segment",
    "risque depart": "Risque départ", "risque dpart": "Risque départ",
    "appetence mrh": "Appétence MRH", "apptence mrh": "Appétence MRH",
    "etat": "Etat", "statut": "Etat",
    "telephone": "Telephone", "telephone_1": "Telephone", "tel": "Telephone",
    "gsm": "Telephone", "mobile": "Telephone"
}

def lire_fichier_agence(fichier):
    nom_fichier = fichier.name.lower()
    contenu = fichier.read()
    fichier.seek(0)
    if contenu.startswith(b'PK'):
        return pd.read_excel(BytesIO(contenu), dtype=str, engine='openpyxl')
    elif contenu.startswith(b'\xd0\xcf\x11\xe0'):
        # Vrai fichier .xls OLE - essais multiples (xlrd → calamine → openpyxl → HTML)
        _errs = []
        try:
            return pd.read_excel(BytesIO(contenu), dtype=str, engine='xlrd')
        except ImportError as e:
            _errs.append(f"xlrd manquant: {e}")
        except Exception as e:
            # Corruption type seen[2]==4 ou autre -> on continue avec fallbacks, pas raise direct
            _errs.append(f"xlrd: {e}")
        # 2. calamine (gère mieux les xls corrompus/spéciaux)
        try:
            return pd.read_excel(BytesIO(contenu), dtype=str, engine='calamine')
        except ImportError as e:
            _errs.append(f"calamine manquant: {e}")
        except Exception as e:
            _errs.append(f"calamine: {e}")
        # 3. openpyxl (au cas où mal détecté)
        try:
            return pd.read_excel(BytesIO(contenu), dtype=str, engine='openpyxl')
        except Exception as e:
            _errs.append(f"openpyxl: {str(e)[:120]}")
        # 4. HTML déguisé (beaucoup d'exports .xls sont des tableaux HTML)
        try:
            try:
                _txt = contenu.decode('utf-8', errors='ignore')
            except:
                _txt = contenu.decode('latin-1', errors='ignore')
            if '<table' in _txt.lower() or '<html' in _txt.lower():
                try:
                    _tables = pd.read_html(StringIO(_txt), flavor='lxml')
                except:
                    _tables = pd.read_html(StringIO(_txt))
                if _tables:
                    _df = _tables[0].astype(str)
                    # Nettoyer les 'nan' issus du cast
                    _df = _df.replace({'nan': None, 'None': None, 'NaT': None})
                    return _df
        except Exception as e:
            _errs.append(f"html: {str(e)[:120]}")
        # 5. Texte brut (TSV/CSV avec header OLE bizarre)
        try:
            try:
                _txt2 = contenu.decode('utf-8', errors='ignore')
            except:
                _txt2 = contenu.decode('latin-1', errors='ignore')
            if '\t' in _txt2[:2000] or ';' in _txt2[:2000]:
                _first = _txt2.split('\n')[0]
                if '\t' in _first:
                    return pd.read_csv(StringIO(_txt2), dtype=str, sep='\t')
                elif ';' in _first:
                    return pd.read_csv(StringIO(_txt2), dtype=str, sep=';')
        except Exception as e:
            _errs.append(f"texte: {str(e)[:120]}")
        raise Exception(f"Erreur lecture xls OLE (fichier corrompu ou format special, ex: seen[2]==4). Essais: {' | '.join(_errs[:4])}. Astuce: ouvrez-le dans Excel -> Enregistrer sous .xlsx puis reuploadez.")
    else:
        try:
            try:
                texte = contenu.decode('utf-8', errors='ignore')
            except:
                texte = contenu.decode('latin-1', errors='ignore')
            # HTML deguise (export .xls en tableau HTML sans header OLE)
            if '<table' in texte.lower()[:5000] or '<html' in texte.lower()[:5000]:
                try:
                    try:
                        _tables = pd.read_html(StringIO(texte), flavor='lxml')
                    except:
                        _tables = pd.read_html(StringIO(texte))
                    if _tables:
                        _dfh = _tables[0].astype(str)
                        _dfh = _dfh.replace({'nan': None, 'None': None, 'NaT': None})
                        return _dfh
                except:
                    pass
            premiere_ligne = texte.split('\n')[0]
            if '\t' in premiere_ligne:
                return pd.read_csv(StringIO(texte), dtype=str, sep='\t')
            elif ';' in premiere_ligne:
                return pd.read_csv(StringIO(texte), dtype=str, sep=';')
            elif ',' in premiere_ligne:
                return pd.read_csv(StringIO(texte), dtype=str, sep=',')
            else:
                return pd.read_csv(StringIO(texte), dtype=str, sep=None, engine='python')
        except Exception as e:
            raise Exception(f"Format de fichier non reconnu: {str(e)}")

def nettoyer_colonne_agence(col):
    col = str(col).strip().lower()
    col = re.sub(r"[^\w\s]", "", col)
    col = col.replace('é', 'e').replace('è', 'e').replace('ê', 'e')
    col = col.replace('à', 'a').replace('â', 'a')
    col = col.replace('ô', 'o').replace('î', 'i').replace('û', 'u')
    return col.strip()

def nettoyer_nom_fichier_agence(nom):
    return re.sub(r"\s+", "-", re.sub(r"[^\w\s-]", "", nom).strip()).upper()

def trouver_colonne_police_agence(df):
    for col in df.columns:
        col_clean = nettoyer_colonne_agence(col)
        if "police" in col_clean or "numero" in col_clean:
            return col
    return None

def detecter_agence_depuis_police_agence(df):
    _mapping = get_mapping_agences()
    _num2code = get_mapping_numero_vers_code()
    col_police = trouver_colonne_police_agence(df)
    numeros_testes = []
    if col_police is not None:
        for val in df[col_police].dropna().head(10):
            val_str = str(val).strip()
            # Si format scientifique (5,86E+14) -> invalide, on ignore pour fallback Intermediaire
            if 'E+' in val_str.upper():
                # essayer de voir si c'est un float scientifique : on le considère invalide
                numeros_testes.append('SCI:' + val_str[:10])
                continue
            chiffres = re.findall(r'\d', val_str)
            if len(chiffres) < 4:
                continue
            quatre = ''.join(chiffres[:4])
            numeros_testes.append(quatre)
            if quatre in _num2code:
                code = _num2code[quatre]
                return code, _mapping[code], quatre, None
        # Si on a trouvé un code mais non référencé, on ne retourne pas tout de suite, on tente fallback Intermediaire
        # On garde numeros_testes pour message si fallback echoue aussi
    # Fallback : essayer via Intermediaire (plus fiable quand N° Police est en 5,86E+14)
    # On utilise nettoyer_colonne_agence pour gérer accents (Intermédiaire) + alias Agence/Interm
    col_inter = None
    for col in df.columns:
        try:
            cleaned = nettoyer_colonne_agence(col)
        except:
            cleaned = str(col).lower()
        low = str(col).lower()
        if 'intermediaire' in cleaned or 'intermediaire' in low or 'intermediair' in low or cleaned in ['agence', 'interm'] or low.strip() in ['agence', 'interm']:
            col_inter = col
            break
    if col_inter is not None:
        for val in df[col_inter].dropna().head(10):
            s = str(val).strip()
            if not s or s.lower() == 'nan':
                continue
            m = re.search(r'(\d{4,5})', s)
            if m:
                quatre = m.group(1)[:4]
                if quatre in _num2code:
                    code = _num2code[quatre]
                    return code, _mapping[code], quatre, None
                else:
                    numeros_testes.append(quatre)
    if col_police is None:
        colonnes = ", ".join([str(c)[:20] for c in df.columns[:5]])
        return None, None, None, f"Colonne 'N° Police' non trouvée. Colonnes: {colonnes}..."
    if numeros_testes:
        # Filtrer les SCI pour message plus clair
        vrais_codes = [c for c in numeros_testes if not c.startswith('SCI:')]
        if vrais_codes:
            codes_non_trouves = list(set(vrais_codes))[:3]
            # Hint : si le code existe dans le défaut mais pas dans le mapping courant (après Remplacer), proposer de l'ajouter via upsert
            hint = ""
            try:
                _missing_in_curr = [c for c in codes_non_trouves if c not in _num2code and any(c == re.sub(r'\D', '', k)[-4:] if re.sub(r'\D', '', k) else False for k in MAPPING_AGENCES_DEFAULT.keys())]
                if _missing_in_curr:
                    hint = f" - '{', '.join(_missing_in_curr)}' existe dans le défaut mais pas dans votre mapping actuel (Remplacer). Ajoutez-le via '➕ Ajouter' ou '📤 Import upsert' si besoin."
            except:
                pass
            return None, None, None, f"Code(s) '{', '.join(codes_non_trouves)}' non référencé(s){hint} - Vérifiez via '⚙️ Gestion MAPPING' ci-dessous puis complétez manuellement."
        else:
            return None, None, None, f"N° Police en format scientifique (ex: 5,86E+14) - détection impossible, et Intermédiaire non trouvé"
    return None, None, None, f"Colonne '{col_police}' vide ou format invalide"

def get_nom_agence_from_code(code_input):
    _mapping = get_mapping_agences()
    _num2code = get_mapping_numero_vers_code()
    code_input = str(code_input).strip().upper()
    if code_input in _mapping:
        return code_input, _mapping[code_input]
    chiffres = re.sub(r'\D', '', code_input)
    if chiffres in _num2code:
        code = _num2code[chiffres]
        return code, _mapping[code]
    return None, None

def extraire_code_from_dropdown(selection):
    if not selection:
        return None, None
    _mapping = get_mapping_agences()
    match = re.search(r"\(([^)]+)\)$", selection)
    if match:
        code = match.group(1)
        return code, _mapping.get(code)
    return None, None

def detecter_format_date_agence(df, col_name):
    if col_name is None:
        return 'FR'
    for val in df[col_name].dropna().head(20):
        if isinstance(val, (pd.Timestamp, datetime)):
            return 'ISO'
        val_str = str(val).strip()
        match = re.match(r'^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})$', val_str)
        if match:
            p1, p2 = int(match.group(1)), int(match.group(2))
            if p1 > 12:
                return 'FR'
            elif p2 > 12:
                return 'US'
        if re.match(r'^\d{4}-\d{2}-\d{2}', val_str):
            return 'ISO'
    return 'FR'

def convertir_date_agence(valeur, format_detecte='FR'):
    if pd.isna(valeur) or str(valeur).strip() == "":
        return None
    if isinstance(valeur, (pd.Timestamp, datetime)):
        return valeur
    val_str = str(valeur).strip()
    if re.match(r'^\d{4}-\d{2}-\d{2}', val_str):
        try:
            return pd.to_datetime(val_str)
        except:
            pass
    match = re.match(r'^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})$', val_str)
    if match:
        p1, p2, p3 = match.group(1), match.group(2), match.group(3)
        if len(p3) == 2:
            p3 = str(2000 + int(p3)) if int(p3) < 50 else str(1900 + int(p3))
        try:
            if format_detecte == 'US':
                return datetime(int(p3), int(p1), int(p2))
            else:
                return datetime(int(p3), int(p2), int(p1))
        except:
            try:
                if format_detecte == 'US':
                    return datetime(int(p3), int(p2), int(p1))
                else:
                    return datetime(int(p3), int(p1), int(p2))
            except:
                pass
    try:
        return pd.to_datetime(val_str, dayfirst=(format_detecte == 'FR'))
    except:
        return None

def detecter_mois_echeance_agence(df, format_date='FR'):
    col_ech = None
    for col in df.columns:
        col_lower = nettoyer_colonne_agence(col)
        if "echeance" in col_lower:
            col_ech = col
            break
    if col_ech is None:
        return datetime.now().strftime("%m%Y")
    mois = {}
    for val in df[col_ech].dropna():
        d = convertir_date_agence(val, format_date)
        if d:
            k = d.strftime("%m%Y")
            mois[k] = mois.get(k, 0) + 1
    return max(mois, key=mois.get) if mois else datetime.now().strftime("%m%Y")

def formater_telephone_agence(valeur):
    if pd.isna(valeur):
        return "NA"
    tel = re.sub(r"\D", "", str(valeur))
    if tel.startswith("212") and len(tel) > 9:
        tel = "0" + tel[3:]
    if len(tel) == 9 and tel[0] in ['6', '7', '5']:
        tel = "0" + tel
    return tel if tel else "NA"

def traiter_fichier_agence(df, nom_agence, format_date='FR', code_base_intermediaire=None, mois_courant=None, annee_courante=None, force_regenerer_tous=False):
    # Conserver df original pour extraction si besoin (deja fait en amont)
    df.columns = [nettoyer_colonne_agence(c) for c in df.columns]
    df = df.rename(columns={k: v for k, v in ALIASES_COLONNES.items() if k in df.columns})
    # Cas double/force : priorité 1 = code digits depuis Intermediaire (si dans mapping jdid) -> base unique pour tout le fichier.
    # Priorité 2 = nouveau par ligne depuis colonnes dédiées (ancien/nouveau/code), jamais via noms (nominations proches).
    _nouveau_par_ligne = {}
    _base_intermediaire = None
    if force_regenerer_tous:
        try:
            _num2c_loc = get_mapping_numero_vers_code()
            # 1) Intermediaire générique : digits + vérif mapping (pas de matching par nom)
            _inter_cols = [c for c in df.columns if c in ['intermediaire', 'agence', 'interm'] or 'intermediaire' in c]
            _inter_cols = [c for c in _inter_cols if 'police' not in c and 'numero' not in c]
            # Exclure les dédiées qui contiennent aussi 'agence' mais avec code/ancien/nouveau (ex: 'code agence')
            _inter_gen = [c for c in _inter_cols if not any(k in c for k in ['code', 'ancien', 'nouveau', 'nouv', 'old', 'new'])]
            if _inter_gen:
                from collections import Counter as _Counter
                _codes_inter = []
                for _cc in _inter_gen:
                    try:
                        for _vv in df[_cc].dropna().head(20):
                            _mI = re.search(r'(\d{4})', str(_vv))
                            if _mI and _mI.group(1)[:4] in _num2c_loc:
                                _codes_inter.append(_mI.group(1)[:4])
                    except:
                        pass
                if _codes_inter:
                    _base_intermediaire = _Counter(_codes_inter).most_common(1)[0][0]
            # 2) Par ligne depuis dédiées (si pas de base Intermediaire)
            if not _base_intermediaire:
                _ded_loc = [c for c in df.columns if any(k in c for k in ['code', 'ancien', 'nouveau', 'nouv', 'old', 'new']) and 'police' not in c and 'numero' not in c]
                _ded_loc = [c for c in _ded_loc if not (c in ['intermediaire', 'agence', 'interm'])]
                if _ded_loc:
                    for _idx, _row in df.iterrows():
                        _cl = []
                        for _cc in _ded_loc:
                            try:
                                _m4 = re.search(r'(\d{4})', str(_row[_cc]))
                                if _m4 and _m4.group(1)[:4] in _num2c_loc:
                                    _cl.append(_m4.group(1)[:4])
                            except:
                                pass
                        if _cl:
                            _nouveau_par_ligne[_idx] = _cl[-1]
        except:
            pass
    if len(df.columns) > 15:
        df = df.iloc[:, :15]
    for col in MODELE_COLONNES:
        if col not in df.columns:
            if col == "Intermediaire":
                df[col] = nom_agence
            elif col == "Duree":
                df[col] = "F"
            elif col == "Heure Echeance":
                df[col] = "12"
            elif col == "Etat":
                df[col] = "En cours"
            elif col in COLONNES_VIDES:
                df[col] = None
            else:
                df[col] = "NA"
    # Ne pas écraser Intermediaire : garder la valeur d\'origine (ex: "A5863 - OUADIE ASSURANCES")
    # Remplir seulement si vide/NA/manquant
    if "Intermediaire" in df.columns:
        df["Intermediaire"] = df["Intermediaire"].apply(lambda x: nom_agence if pd.isna(x) or str(x).strip() == "" or str(x).strip().upper() == "NA" else x)
    else:
        df["Intermediaire"] = nom_agence
    df["Telephone"] = df["Telephone"].apply(formater_telephone_agence)
    df["Date Effet"] = df["Date Effet"].apply(lambda x: convertir_date_agence(x, format_date))
    df["Date Echeance"] = df["Date Echeance"].apply(lambda x: convertir_date_agence(x, format_date))
    # === Si Date Effet vide → Jour/Mois de l\'échéance + année précédente (ex: 10/10/2026 → 10/10/2025) ===
    try:
        for idx in df.index:
            eff = df.at[idx, "Date Effet"]
            ech = df.at[idx, "Date Echeance"]
            if (pd.isna(eff) or eff is None) and pd.notna(ech) and ech is not None:
                try:
                    # ech est déjà datetime/Timestamp
                    # On garde jour/mois de ech, année -1
                    if isinstance(ech, pd.Timestamp):
                        ech_dt = ech.to_pydatetime()
                    else:
                        ech_dt = ech
                    # Gestion 29 fevrier -> 28 fevrier si année précédente non bissextile
                    try:
                        new_eff = ech_dt.replace(year=ech_dt.year - 1)
                    except ValueError:
                        # 29/02 -> 28/02
                        import calendar
                        # Reculer d\'un an en gardant jour valide
                        new_eff = ech_dt - pd.DateOffset(years=1)
                        # pd.DateOffset renvoie Timestamp, convertir
                        if isinstance(new_eff, pd.Timestamp):
                            new_eff = new_eff.to_pydatetime()
                    df.at[idx, "Date Effet"] = pd.to_datetime(new_eff)
                except Exception as _e:
                    pass
    except Exception as _e:
        pass
    for col in MODELE_COLONNES:
        if col in ["Date Effet", "Date Echeance"]:
            continue
        if col in COLONNES_VIDES:
            # Segment / Risque départ / Appétence MRH : laisser vide (pas de NA)
            df[col] = df[col].apply(lambda x: None if pd.isna(x) or str(x).strip() == "" or str(x).strip().upper() == "NA" else x)
        elif col == "Etat":
            # Etat : toujours En cours par défaut si vide
            df[col] = df[col].apply(lambda x: "En cours" if pd.isna(x) or str(x).strip() == "" or str(x).strip().upper() == "NA" else x)
        elif col == "Immatriculation":
            # Immatriculation : vide, NA, ou que des caractères non alphanumériques (---, --, ///, ...) -> NA obligatoire
            def _clean_immat(x):
                if pd.isna(x):
                    return "NA"
                s = str(x).strip()
                if s == "" or s.upper() == "NA":
                    return "NA"
                if not re.search(r'[A-Za-z0-9]', s):
                    return "NA"
                return x
            df[col] = df[col].apply(_clean_immat)
        else:
            # CIN / Nom/Raison sociale / Usage / autres : vide -> NA
            df[col] = df[col].apply(lambda x: "NA" if pd.isna(x) or str(x).strip() == "" else x)
    df = df.reindex(columns=MODELE_COLONNES)
    # === N° Police : garder si déjà m9ada, régénérer seulement si scientifique/vide/NA (sans doublons) ===
    # Si force_regenerer_tous (cas double code ancien+nouveau) : tout régénérer à partir de 0 (00001)
    if code_base_intermediaire is not None:
        try:
            if mois_courant is None:
                mois_courant = datetime.now().strftime('%m')
            if annee_courante is None:
                annee_courante = datetime.now().strftime('%Y')
            if force_regenerer_tous:
                mask_a_regenerer = pd.Series([True] * len(df), index=df.index)
            else:
                mask_a_regenerer = df["N° Police"].apply(est_police_scientifique_ou_vide)
            nb_a_regenerer = int(mask_a_regenerer.sum())
            if nb_a_regenerer > 0:
                if force_regenerer_tous and _base_intermediaire:
                    # Priorité Intermediaire (digits vérifiés mapping) : base unique pour tout le fichier dès 00001
                    _bc = re.sub(r'\D', '', str(_base_intermediaire)) or re.sub(r'\D', '', str(code_base_intermediaire)) or '0000'
                    _nouv_all = generer_numeros_police(_bc, nb_a_regenerer, mois_courant, annee_courante)
                    df.loc[mask_a_regenerer, "N° Police"] = _nouv_all[:nb_a_regenerer]
                elif force_regenerer_tous and _nouveau_par_ligne:
                    # Par ligne : base = nouveau de SA ligne (dédiée), compteur par base dès 00001
                    _code_fallback = re.sub(r'\D', '', str(code_base_intermediaire)) or '0000'
                    _compteurs = {}
                    _gardes2 = set()
                    _vals_par_idx = {}
                    for _idx in df.index[mask_a_regenerer]:
                        _b = _nouveau_par_ligne.get(_idx, _code_fallback)
                        _b = re.sub(r'\D', '', str(_b)) or _code_fallback
                        _cpt = _compteurs.get(_b, 0) + 1
                        # Sauter doublons improbables (sécurité)
                        _guard2 = 0
                        while f"{_b}{mois_courant}{annee_courante}{_cpt:05d}" in _gardes2 and _guard2 < 10000:
                            _cpt += 1
                            _guard2 += 1
                        _compteurs[_b] = _cpt
                        _cand2 = f"{_b}{mois_courant}{annee_courante}{_cpt:05d}"
                        _gardes2.add(_cand2)
                        _vals_par_idx[_idx] = _cand2
                    for _idx, _vv in _vals_par_idx.items():
                        df.at[_idx, "N° Police"] = _vv
                else:
                    # Éviter doublons avec les N° déjà gardés
                    try:
                        gardes = set(df.loc[~mask_a_regenerer, "N° Police"].astype(str).tolist())
                    except:
                        gardes = set()
                    nouveaux = []
                    _i = 1
                    _code_clean = re.sub(r'\D', '', str(code_base_intermediaire)) or '0000'
                    _base = f"{_code_clean}{mois_courant}{annee_courante}"
                    # Générer en sautant les existants (sécurité anti-boucle infinie)
                    _guard = 0
                    while len(nouveaux) < nb_a_regenerer and _guard < nb_a_regenerer + len(gardes) + 1000:
                        _guard += 1
                        cand = f"{_base}{_i:05d}"
                        _i += 1
                        if cand in gardes:
                            continue
                        nouveaux.append(cand)
                        gardes.add(cand)
                    df.loc[mask_a_regenerer, "N° Police"] = nouveaux[:nb_a_regenerer]
        except Exception as e:
            print(f"Erreur generation N° Police: {e}")
    return df

def _clean_excel_str(val):
    """Nettoie les caractères illégaux pour Excel (openpyxl) : \\x00-\\x08 etc. ne peuvent pas être utilisés dans les worksheets."""
    if pd.isna(val) or val is None:
        return val
    s = str(val)
    # Supprime les caractères de contrôle illégaux pour openpyxl
    s = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', s)
    return s

def to_excel_bytes_agence(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Donnees"
    ws.sheet_view.showGridLines = False
    bordure = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    header_font = Font(bold=True)
    header_fill = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell_alignment = Alignment(vertical='center')
    nb_colonnes = len(df.columns)
    nb_lignes = len(df) + 1
    for col_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = bordure
        cell.alignment = header_alignment
    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        for col_idx, (col_name, value) in enumerate(zip(df.columns, row), 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if col_name == "N° Police":
                v = _clean_excel_str(str(value) if pd.notna(value) else "NA")
                cell.value = v
                cell.number_format = '@'
            elif col_name in ["Date Effet", "Date Echeance"]:
                if pd.notna(value):
                    cell.value = value
                    cell.number_format = 'DD/MM/YYYY'
                else:
                    cell.value = None
            else:
                if pd.isna(value):
                    cell.value = None
                else:
                    cell.value = _clean_excel_str(value)
            cell.border = bordure
            cell.alignment = cell_alignment
    largeurs = {
        "N° Police": 18, "Intermediaire": 22, "CIN": 12,
        "Nom/Raison sociale": 25, "Usage": 15, "Immatriculation": 15,
        "Duree": 8, "Date Effet": 14, "Date Echeance": 14,
        "Heure Echeance": 14, "Segment": 10, "Risque départ": 13,
        "Appétence MRH": 14, "Etat": 12, "Telephone": 14
    }
    for col_idx, col_name in enumerate(df.columns, 1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = largeurs.get(col_name, 15)
    ws.freeze_panes = 'A2'
    derniere_colonne = get_column_letter(nb_colonnes)
    ws.print_area = f'A1:{derniere_colonne}{nb_lignes}'
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

def creer_zip_agence(fichiers):
    buf = BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for nom, data in fichiers:
            zf.writestr(nom, data)
    buf.seek(0)
    return buf.getvalue()


def analyser_double_code_agence(df):
    """Détecte le cas 'deux colonnes code agence (ancien + nouveau)'.
    Retourne (double_detecte: bool, nouveau_code_fichier: str|None, details: dict).
    Règle : parmi les colonnes type agence/code (hors N° Police), si >=2 colonnes
    contiennent des codes, on prend par ligne le DERNIER code présent dans le mapping
    (nouveau), et au niveau fichier le plus fréquent (mode).
    """
    try:
        _mapping = get_mapping_agences()
        _num2code = get_mapping_numero_vers_code()
    except:
        return False, None, {}
    # Colonne police à exclure
    try:
        _col_police = trouver_colonne_police_agence(df)
    except:
        _col_police = None
    # Colonnes candidates : header évoque agence/code/intermediaire/ancien/nouveau,
    # OU valeurs = purs codes agence (ex: '6656', 'AGENT_A6656') même si header atypique
    def _is_pur_code(val):
        if pd.isna(val):
            return False
        s = str(val).strip()
        if not s or s.lower() in ('nan', 'na', 'none'):
            return False
        return bool(re.match(r'^\s*(AGENT_A|BGD_B|A|B)?\s*\d{4}\s*$', s, re.IGNORECASE))
    _cands = []
    for col in df.columns:
        if col == _col_police:
            continue
        low = str(col).lower()
        try:
            cleaned = nettoyer_colonne_agence(col)
        except:
            cleaned = low
        if 'police' in cleaned or 'numero' in cleaned:
            continue
        _header_ok = (
            'agence' in cleaned or 'intermediaire' in cleaned or 'intermediaire' in low
            or 'intermediair' in low or 'code' in cleaned
            or 'ancien' in cleaned or 'nouveau' in cleaned or 'nouv' in cleaned
            or 'old' in cleaned or 'new' in cleaned
        )
        if _header_ok:
            # Exclure dates / téléphone / cin / immat / usage / nom / etat / duree / segment...
            # sauf si c'est clairement une colonne code/agence/ancien/nouveau
            if any(k in cleaned for k in ['date', 'telephone', 'tel ', 'tel', 'cin', 'immatriculation', 'immat', 'matricule', 'plaque', 'usage', 'nom', 'client', 'raison', 'etat', 'statut', 'duree', 'dure', 'heure', 'segment', 'risque', 'appetence', 'echeance', 'effet']):
                if not ('agence' in cleaned or 'intermediaire' in cleaned or 'interm' in cleaned or 'code' in cleaned or 'ancien' in cleaned or 'nouveau' in cleaned):
                    continue
            _cands.append(col)
            continue
        # Header atypique : accepter si valeurs = purs codes agence (>=50% des 10 premières)
        try:
            _vals = df[col].dropna().head(10)
            _vals = [v for v in _vals if str(v).strip() not in ('', 'nan', 'NA', 'None')]
            if len(_vals) >= 3 and sum(1 for v in _vals if _is_pur_code(v)) >= max(2, len(_vals) * 0.5):
                _cands.append(col)
        except:
            pass
    if len(_cands) < 2:
        return False, None, {'colonnes_candidates': [str(c) for c in _cands]}
    # Extraire codes par colonne (4 chiffres) et vérifier présence mapping
    def _code4_of(val):
        if pd.isna(val):
            return None
        s = str(val).strip()
        if not s or s.lower() in ('nan', 'na', 'none'):
            return None
        m = re.search(r'(\d{4})', s)
        if m:
            return m.group(1)[:4]
        return None
    # Compter colonnes qui contiennent vraiment des codes du mapping
    _cols_avec_codes = []
    for col in _cands:
        try:
            _vals = df[col].dropna().head(10)
            _hits = sum(1 for v in _vals if (_code4_of(v) in _num2code) if _code4_of(v))
            if _hits >= 1:
                _cols_avec_codes.append(col)
        except:
            pass
    if len(_cols_avec_codes) < 2:
        # Pas deux colonnes avec codes du mapping : vérifier quand même si 2 colonnes avec codes quelconques
        # (ancien code peut ne plus être dans le mapping) -> on prend quand même le dernier existant
        _cols_avec_quelconque = []
        for col in _cands:
            try:
                _vals = df[col].dropna().head(10)
                _hq = sum(1 for v in _vals if _code4_of(v))
                if _hq >= 1:
                    _cols_avec_quelconque.append(col)
            except:
                pass
        if len(_cols_avec_quelconque) < 2:
            return False, None, {'colonnes_candidates': [str(c) for c in _cands]}
        _cols_avec_codes = _cols_avec_quelconque
    # Nouveau = DERNIERE colonne dédiée UNIQUEMENT (jamais Intermediaire générique -> nominations proches = erreur).
    # Dédiée = header contient 'code'/'ancien'/'nouveau'/'new'/'old'. Intermediaire générique exclu d'office.
    def _is_dediee(col):
        try:
            _cl = nettoyer_colonne_agence(col)
        except:
            _cl = str(col).lower()
        return any(k in _cl for k in ['code', 'ancien', 'nouveau', 'nouv', 'old', 'new'])
    def _is_intermediaire_generique(col):
        try:
            _cl = nettoyer_colonne_agence(col)
        except:
            _cl = str(col).lower()
        _low = str(col).lower()
        return ('intermediaire' in _cl or 'intermediaire' in _low or 'intermediair' in _low or _cl in ['agence', 'interm'])
    _dediees = [c for c in _cols_avec_codes if _is_dediee(c)]
    if _dediees:
        _cols_nouveau = _dediees
    else:
        # Aucune dédiée : exclure Intermediaire générique (sinon nomination proche = faux nouveau)
        _non_generiques = [c for c in _cols_avec_codes if not _is_intermediaire_generique(c)]
        if not _non_generiques:
            return False, None, {'colonnes_candidates': [str(c) for c in _cols_avec_codes], 'raison': 'que Intermediaire, pas de colonne nouveau dediee'}
        _cols_nouveau = _non_generiques
    # Cas 1 dédiée + générique avec codes différents (ancien en Intermediaire, nouveau en Code) -> double aussi
    if len(_dediees) == 1 and len(_cols_avec_codes) >= 2:
        pass  # on garde _cols_nouveau = dédiée (nouveau), Intermediaire ignoré pour le choix
    # Par ligne : prendre le DERNIER code dédié présent dans le mapping (nouveau)
    _choisis = []
    for _, row in df.head(1000).iterrows():
        _codes_ligne = []
        for col in _cols_nouveau:
            try:
                _c4 = _code4_of(row[col])
                if _c4 and _c4 in _num2code:
                    _codes_ligne.append(_c4)
            except:
                pass
        if _codes_ligne:
            _choisis.append(_codes_ligne[-1])
    if not _choisis:
        return False, None, {'colonnes_candidates': [str(c) for c in _cols_avec_codes]}
    # Mode (plus fréquent) au niveau fichier
    try:
        _nouveau_dedie = max(set(_choisis), key=_choisis.count)
    except:
        _nouveau_dedie = _choisis[-1]
    # Priorité Intermediaire (digits vérifiés mapping) : si le code Intermediaire est dans le mapping jdid, on le prend (demande utilisateur)
    try:
        from collections import Counter as _Counter2
        _inter_gen2 = [c for c in df.columns if _is_intermediaire_generique(c)]
        _codes_inter2 = []
        for _cc2 in _inter_gen2:
            try:
                for _vv2 in df[_cc2].dropna().head(20):
                    _mI2 = re.search(r'(\d{4})', str(_vv2))
                    if _mI2 and _mI2.group(1)[:4] in _num2code:
                        _codes_inter2.append(_mI2.group(1)[:4])
            except:
                pass
        if _codes_inter2:
            _nouveau_inter = _Counter2(_codes_inter2).most_common(1)[0][0]
            return True, _nouveau_inter, {'colonnes_candidates': [str(c) for c in _cols_avec_codes], 'colonnes_nouveau': [str(c) for c in _cols_nouveau], 'exemples': _choisis[:5], 'source': 'intermediaire', 'code_intermediaire': _nouveau_inter, 'code_dedie': _nouveau_dedie}
    except:
        pass
    return True, _nouveau_dedie, {'colonnes_candidates': [str(c) for c in _cols_avec_codes], 'colonnes_nouveau': [str(c) for c in _cols_nouveau], 'exemples': _choisis[:5], 'source': 'dediee'}

def extraire_code_intermediaire_pour_generation(df, nom_agence=None, code_agence=None):
    """Extrait le code numerique de base depuis la colonne Intermediaire ou fallback agence.
    Retourne ex: '5863' ou None si non trouve.
    """
    # 1) Essayer colonne Intermediaire (avec nettoyage accents + alias Agence/Interm)
    col_inter = None
    for col in df.columns:
        low = str(col).lower()
        try:
            cleaned = nettoyer_colonne_agence(col)
        except:
            cleaned = low
        if 'intermediaire' in cleaned or 'intermediaire' in low or 'intermediair' in low or cleaned in ['agence', 'interm'] or low.strip() in ['agence', 'interm']:
            col_inter = col
            break
    if col_inter is not None:
        for val in df[col_inter].dropna().head(5):
            s = str(val).strip()
            if not s or s.lower() == 'nan':
                continue
            # chercher 4 chiffres consecutifs (ex: 5863, 8057)
            m = re.search(r'(\d{4,5})', s)
            if m:
                return m.group(1)
            # fallback: tout chiffres
            digs = re.sub(r'\D', '', s)
            if len(digs) >= 4:
                return digs[:4]
    # 2) Fallback via nom_agence / code_agence
    if code_agence:
        digs = re.sub(r'\D', '', str(code_agence))
        if digs:
            # prendre les 4 derniers chiffres (ex: A5863 -> 5863)
            return digs[-4:] if len(digs) >=4 else digs
    if nom_agence:
        # essayer de retrouver code via mapping
        try:
            _map = get_mapping_agences()
            for c, n in _map.items():
                if n == nom_agence:
                    digs = re.sub(r'\D', '', c)
                    if digs:
                        return digs[-4:]
        except:
            pass
    return None

def generer_numeros_police(code_base, n, mois=None, annee=None):
    """Genere n numeros sequentiels: code_base + MM + YYYY + 00001..."""
    if mois is None:
        mois = datetime.now().strftime('%m')
    if annee is None:
        annee = datetime.now().strftime('%Y')
    # s'assurer que code_base est bien 4 chiffres (ou plus)
    code_base = str(code_base).strip()
    # enlever tout non-digit
    code_base = re.sub(r'\D', '', code_base)
    if not code_base:
        code_base = '0000'
    base = f"{code_base}{mois}{annee}"
    return [f"{base}{i:05d}" for i in range(1, n+1)]

def est_police_scientifique_ou_vide(val):
    """Retourne True si N° Police doit être régénéré : vide/NA ou format scientifique (5,86E+14). Sinon False (garder l'original m9ada)."""
    if pd.isna(val) or val is None:
        return True
    s = str(val).strip()
    if s == "" or s.upper() in ("NA", "NAN", "NONE"):
        return True
    # Format scientifique Excel : contient E+ (ex: 5,86E+14, 5.86E+14)
    if "E+" in s.upper():
        return True
    # Variante sans + mais avec E et séparateur décimal (ex: 5,86E14)
    if re.search(r'\d[,\.]\d+\s*[Ee]\s*\d+', s):
        return True
    return False

def app_traitement_agence():
    st.markdown("Téléversez vos fichiers d'agence (Excel, CSV, TSV, .xls déguisé) — détection automatique de l'agence via les 4 premiers chiffres du N° Police.")
    fichiers_upload = st.file_uploader(
        "📁 Glissez vos fichiers ici",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        key="agence_uploader"
    )
    if fichiers_upload:
        fichiers_info = []
        for fichier in fichiers_upload:
            try:
                df = lire_fichier_agence(fichier)
                code, nom, numero, raison_echec = detecter_agence_depuis_police_agence(df)
                col_date = None
                for col in df.columns:
                    if 'date' in str(col).lower():
                        col_date = col
                        break
                format_date = detecter_format_date_agence(df, col_date)
                # Extraction code intermediaire pour generation N° Police (depuis Intermediaire ou fallback agence)
                try:
                    code_base = extraire_code_intermediaire_pour_generation(df, nom, code)
                except:
                    code_base = None
                # Cas double code (ancien + nouveau) : toujours dernier mapping, nouveau code, regen totale des 00001
                _double = False
                _nouveau_code = None
                _double_details = {}
                try:
                    _double, _nouveau_code, _double_details = analyser_double_code_agence(df)
                except:
                    pass
                if _double and _nouveau_code:
                    # Forcer le nouveau code pour generation + detection agence si besoin
                    code_base = _nouveau_code
                    try:
                        _c2, _n2 = get_nom_agence_from_code(_nouveau_code)
                        if _n2:
                            code, nom, numero = _c2, _n2, _nouveau_code
                            raison_echec = None
                    except:
                        pass
                fichiers_info.append({
                    'fichier': fichier,
                    'nom_fichier': fichier.name,
                    'df': df,
                    'code': code,
                    'nom_agence': nom,
                    'numero': numero,
                    'format_date': format_date,
                    'detecte': nom is not None,
                    'raison_echec': raison_echec,
                    'code_base': code_base,
                    'code_base_effectif': code_base,
                    'double_code': _double,
                    'nouveau_code': _nouveau_code,
                    'double_details': _double_details
                })
            except Exception as e:
                fichiers_info.append({
                    'fichier': fichier,
                    'nom_fichier': fichier.name,
                    'df': None,
                    'code': None,
                    'nom_agence': None,
                    'numero': None,
                    'format_date': 'FR',
                    'detecte': False,
                    'raison_echec': f"Erreur lecture: {str(e)}",
                    'code_base': None,
                    'code_base_effectif': None,
                    'double_code': False,
                    'nouveau_code': None,
                    'double_details': {}
                })
        detectes = [f for f in fichiers_info if f['detecte']]
        non_detectes = [f for f in fichiers_info if not f['detecte']]
        if detectes:
            st.success(f"✅ **{len(detectes)}** fichier(s) détecté(s) automatiquement")
            with st.expander("📋 Fichiers détectés", expanded=False):
                for f in detectes:
                    st.write(f"• **{f['nom_fichier']}** → {f['nom_agence']} ({f['code']})")
        if non_detectes:
            st.warning(f"⚠️ **{len(non_detectes)}** fichier(s) nécessite(nt) une saisie manuelle")
            st.markdown("### 🔧 Compléter les agences manquantes")
            for idx, f in enumerate(non_detectes):
                with st.container():
                    st.markdown(f"**📄 {f['nom_fichier']}**")
                    if f['raison_echec']:
                        st.error(f"❌ **Raison :** {f['raison_echec']}")
                    if f['df'] is not None:
                        col1, col2 = st.columns(2)
                        with col1:
                            code_manuel = st.text_input(
                                "Code agence (ex: 5180)",
                                key=f"agence_manuel_{idx}",
                                placeholder="Tapez le code..."
                            )
                        with col2:
                            selection = st.selectbox(
                                "Ou choisir dans la liste",
                                get_liste_agences_dropdown(),
                                key=f"agence_liste_{idx}"
                            )
                        code_final, nom_final = None, None
                        if code_manuel:
                            code_final, nom_final = get_nom_agence_from_code(code_manuel)
                            if nom_final:
                                st.success(f"✅ Agence trouvée : **{nom_final}**")
                            else:
                                st.error("❌ Code non reconnu")
                        elif selection:
                            code_final, nom_final = extraire_code_from_dropdown(selection)
                            if nom_final:
                                st.success(f"✅ Agence sélectionnée : **{nom_final}**")
                        f['code'] = code_final
                        f['nom_agence'] = nom_final
                        f['detecte'] = nom_final is not None
                        # Re-extraire code_base apres correction agence
                        if f['detecte']:
                            try:
                                f['code_base'] = extraire_code_intermediaire_pour_generation(f['df'], f['nom_agence'], f['code'])
                                f['code_base_effectif'] = f['code_base']
                            except:
                                pass
                    st.markdown("---")
        # === N° Police : garder si m9ada, régénérer si scientifique/vide, ou TOUT si double code ancien+nouveau ===
        st.markdown("### 🔢 N° Police — conservation si m9ada, génération si scientifique / double code")
        st.caption(f"Si N° Police déjà correct → conservé. Si scientifique (5,86E+14) / vide → régénéré. Si 2 codes agence (ancien+nouveau) → nouveau code + régénération totale dès 00001.")
        mois_courant = datetime.now().strftime('%m')
        annee_courante = datetime.now().strftime('%Y')
        st.info(f"📅 Mois/Année en cours pour génération : **{mois_courant}/{annee_courante}**")
        # Pour chaque fichier valide : compter combien à régénérer vs déjà corrects (+ cas double code)
        for idx2, f in enumerate([x for x in fichiers_info if x['df'] is not None]):
            _is_double_auto = bool(f.get('double_code') and f.get('nouveau_code'))
            # Compter N° scientifiques/vides (base)
            try:
                _col_pol = trouver_colonne_police_agence(f['df'])
                if _col_pol is not None:
                    _nb_invalid_auto = int(f['df'][_col_pol].apply(est_police_scientifique_ou_vide).sum())
                else:
                    _nb_invalid_auto = len(f['df'])
            except:
                _col_pol = None
                _nb_invalid_auto = 0
            _nb_ok_auto = len(f['df']) - _nb_invalid_auto
            # Re-essayer extraction code base si besoin
            if not f.get('code_base_effectif'):
                try:
                    f['code_base'] = extraire_code_intermediaire_pour_generation(f['df'], f.get('nom_agence'), f.get('code'))
                    f['code_base_effectif'] = f['code_base']
                except:
                    pass
            code_base_auto = f.get('code_base')
            with st.container():
                if _is_double_auto:
                    _det = f.get('double_details', {}) or {}
                    _cols_txt = ', '.join(_det.get('colonnes_candidates', [])[:3])
                    st.markdown(f"**📄 {f['nom_fichier']}** — *{f.get('nom_agence') or 'Agence non définie'}* — 🔀 Double code détecté (ancien+nouveau) [{_cols_txt}] → nouveau **{f.get('nouveau_code')}**, régénération totale dès 00001")
                    f['code_base'] = f.get('nouveau_code')
                    f['code_base_effectif'] = f.get('nouveau_code')
                    f['nb_a_regenerer'] = len(f['df'])
                    f['nb_deja_ok'] = 0
                    f['force_regenerer'] = True
                    _eff2 = f.get('code_base_effectif')
                    if _eff2:
                        _prev2 = generer_numeros_police(_eff2, min(2, len(f['df'])), mois_courant, annee_courante)
                        st.code(" → ".join(_prev2) + (f" … +{len(f['df'])-2} autres (total {len(f['df'])})" if len(f['df'])>2 else ""), language=None)
                        st.caption(f"Tout sera régénéré depuis {_eff2} (ancien effacé). Intermediaire et reste inchangés.")
                else:
                    # Automatique : si mapping jdid fih nouveau o mafihch 9dim → forcer total sans demander
                    _auto_force = False
                    _auto_nouveau = None
                    try:
                        _num2c_now = get_mapping_numero_vers_code()
                        _col_pol2 = _col_pol
                        if _col_pol2 is not None:
                            _pol_codes = set()
                            for _v in f['df'][_col_pol2].dropna().head(20):
                                _m4 = re.search(r'(\d{4})', str(_v))
                                if _m4:
                                    _pol_codes.add(_m4.group(1)[:4])
                            _pol_hors = [c for c in _pol_codes if c not in _num2c_now]
                            if _pol_hors:
                                # Ancien code (dans police) absent du mapping → nouveau depuis colonnes DEDIEES uniquement (jamais Intermediaire générique)
                                _cands_now = []
                                for _cc in f['df'].columns:
                                    if _cc == _col_pol2:
                                        continue
                                    try:
                                        _ccl = nettoyer_colonne_agence(_cc)
                                    except:
                                        _ccl = str(_cc).lower()
                                    _cclow = str(_cc).lower()
                                    _is_ded = any(k in _ccl for k in ['code', 'ancien', 'nouveau', 'nouv', 'old', 'new'])
                                    _is_gen = ('intermediaire' in _ccl or 'intermediaire' in _cclow or 'intermediair' in _cclow or _ccl in ['agence', 'interm'])
                                    if _is_gen and not _is_ded:
                                        continue  # Intermediaire générique exclu (nominations proches = erreur)
                                    if not _is_ded:
                                        continue  # Seulement colonnes dédiées ancien/nouveau/code
                                    try:
                                        for _vv in f['df'][_cc].dropna().head(20):
                                            _mm = re.search(r'(\d{4})', str(_vv))
                                            if _mm and _mm.group(1)[:4] in _num2c_now:
                                                _cands_now.append(_mm.group(1)[:4])
                                                break
                                    except:
                                        pass
                                if _cands_now:
                                    _auto_nouveau = max(set(_cands_now), key=_cands_now.count)
                                    _auto_force = True
                    except:
                        pass
                    if _auto_force and _auto_nouveau:
                        st.markdown(f"**📄 {f['nom_fichier']}** — *{f.get('nom_agence') or 'Agence non définie'}* — 🔀 Ancien code absent du mapping jdid → nouveau **{_auto_nouveau}**, régénération totale automatique dès 00001")
                        f['code_base'] = _auto_nouveau
                        f['code_base_effectif'] = _auto_nouveau
                        f['nb_a_regenerer'] = len(f['df'])
                        f['nb_deja_ok'] = 0
                        f['force_regenerer'] = True
                        _prevA = generer_numeros_police(_auto_nouveau, min(2, len(f['df'])), mois_courant, annee_courante)
                        st.code(" → ".join(_prevA) + (f" … +{len(f['df'])-2} autres (total {len(f['df'])})" if len(f['df'])>2 else ""), language=None)
                        st.caption(f"Tout sera régénéré depuis {_auto_nouveau} (ancien effacé). Intermediaire et reste inchangés.")
                    else:
                        st.markdown(f"**📄 {f['nom_fichier']}** — *{f.get('nom_agence') or 'Agence non définie'}* — ✅ {_nb_ok_auto} déjà m9ada / 🔄 {_nb_invalid_auto} à générer")
                        f['nb_a_regenerer'] = _nb_invalid_auto
                        f['nb_deja_ok'] = _nb_ok_auto
                        f['force_regenerer'] = False
                        if _nb_invalid_auto == 0:
                            st.success("Tous les N° Police déjà corrects — conservés.")
                            try:
                                _ex = f['df'][_col_pol].dropna().head(2).tolist() if _col_pol is not None else []
                                if _ex:
                                    st.code("Ex conservés : " + " → ".join([str(x)[:20] for x in _ex]), language=None)
                            except:
                                pass
                        else:
                            col_g1, col_g2, col_g3 = st.columns([2, 2, 3])
                            with col_g1:
                                if code_base_auto:
                                    st.success(f"Code extrait : **{code_base_auto}** → ex: {code_base_auto}{mois_courant}{annee_courante}00001")
                                else:
                                    st.warning("⚠️ Code Intermédiaire non trouvé (colonne Intermediaire vide ou illisible)")
                            with col_g2:
                                key_manual = f"agence_codebase_{idx2}_{f['nom_fichier']}"
                                val_manual = st.text_input(
                                    "Code base manuel (4 chiffres, ex: 5863) — laisser vide pour garder auto",
                                    value="",
                                    placeholder=code_base_auto or "Ex: 5863",
                                    key=key_manual,
                                    help="Seulement pour les N° scientifiques/vides. Tapez ici les 4 chiffres si besoin."
                                )
                                if val_manual.strip():
                                    digits = re.sub(r'\D','', val_manual.strip())
                                    if digits:
                                        f['code_base_effectif'] = digits[:4] if len(digits)>=4 else digits
                                        st.caption(f"→ Effectif manuel : **{f['code_base_effectif']}**")
                                    else:
                                        st.error("Code invalide (chiffres requis)")
                                        f['code_base_effectif'] = None
                                else:
                                    f['code_base_effectif'] = code_base_auto
                            with col_g3:
                                eff = f.get('code_base_effectif')
                                if eff:
                                    preview = generer_numeros_police(eff, min(2, _nb_invalid_auto), mois_courant, annee_courante)
                                    st.code(" → ".join(preview) + (f" … +{_nb_invalid_auto-2} autres à générer" if _nb_invalid_auto>2 else ""), language=None)
                                    st.caption(f"{_nb_invalid_auto} à générer (sur {len(f['df'])}), {_nb_ok_auto} gardés tels quels.")
                                else:
                                    st.error("Aucun code → génération impossible — saisissez manuel")
                st.markdown("---")
        fichiers_valides = [f for f in fichiers_info if f['df'] is not None]
        # Tous prêts si agence ok ET (pas de génération nécessaire OU code_base ok)
        tous_prets = all((f['detecte'] or f['nom_agence']) and (f.get('nb_a_regenerer', 0) == 0 or f.get('code_base_effectif')) for f in fichiers_valides)
        st.markdown("### 🚀 Traitement")
        if not fichiers_valides:
            st.error("❌ Aucun fichier valide")
        elif not tous_prets:
            # Message plus precis
            manques = [f['nom_fichier'] for f in fichiers_valides if not (f['detecte'] or f['nom_agence']) or (f.get('nb_a_regenerer', 0) > 0 and not f.get('code_base_effectif'))]
            st.warning(f"ℹ️ Complétez les agences / codes manquants : {', '.join(manques[:3])}")
            # Detail
            for f in fichiers_valides:
                if f.get('nb_a_regenerer', 0) > 0 and not f.get('code_base_effectif'):
                    st.info(f"📄 {f['nom_fichier']} : {f.get('nb_a_regenerer', 0)} N° scientifiques → saisissez le Code base Intermédiaire (ex: 5863) ci-dessus")
        if st.button("Lancer le traitement", type="primary", disabled=not tous_prets or not fichiers_valides, use_container_width=True, key="agence_lancer"):
            fichiers_traites = []
            resultats = []
            progress = st.progress(0)
            # mois/annee en cours pour generation (deja calcule plus haut, on recalcule pour securite)
            _mois_c = datetime.now().strftime('%m')
            _annee_c = datetime.now().strftime('%Y')
            for idx, f in enumerate(fichiers_info):
                if f['df'] is not None and f['nom_agence'] and (f.get('nb_a_regenerer', 0) == 0 or f.get('code_base_effectif')):
                    try:
                        df_traite = traiter_fichier_agence(f['df'].copy(), f['nom_agence'], f['format_date'], f['code_base_effectif'], _mois_c, _annee_c, bool(f.get('force_regenerer', False)))
                        mois = detecter_mois_echeance_agence(f['df'], f['format_date'])
                        # Le nom de sortie garde le mois d'echeance detecte, mais le N° Police utilise mois/annee en cours
                        nom_sortie = f"ASSURCALL_{nettoyer_nom_fichier_agence(f['nom_agence'])}_{mois}.xlsx"
                        excel_bytes = to_excel_bytes_agence(df_traite)
                        fichiers_traites.append((nom_sortie, excel_bytes))
                        resultats.append({
                            "Source": f['nom_fichier'],
                            "Agence": f['nom_agence'],
                            "Lignes": len(df_traite),
                            "Sortie": nom_sortie,
                            "Statut": "✅ OK"
                        })
                    except Exception as e:
                        resultats.append({
                            "Source": f['nom_fichier'],
                            "Agence": f['nom_agence'] or "-",
                            "Lignes": 0,
                            "Sortie": "-",
                            "Statut": f"❌ {e}"
                        })
                elif f['df'] is None:
                    resultats.append({
                        "Source": f['nom_fichier'],
                        "Agence": "-",
                        "Lignes": 0,
                        "Sortie": "-",
                        "Statut": f"❌ {f['raison_echec']}"
                    })
                progress.progress((idx + 1) / len(fichiers_info))
            st.markdown("---")
            st.subheader("📊 Résumé")
            st.dataframe(pd.DataFrame(resultats), use_container_width=True, hide_index=True)
            if fichiers_traites:
                st.markdown("---")
                st.subheader("💾 Téléchargement")
                if len(fichiers_traites) == 1:
                    nom, data = fichiers_traites[0]
                    st.download_button(
                        f"📥 Télécharger {nom}",
                        data=data,
                        file_name=nom,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                        key="agence_dl_single"
                    )
                else:
                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button(
                            f"📦 Télécharger ZIP ({len(fichiers_traites)} fichiers)",
                            data=creer_zip_agence(fichiers_traites),
                            file_name=f"ASSURCALL_BATCH_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                            mime="application/zip",
                            type="primary",
                            use_container_width=True,
                            key="agence_dl_zip"
                        )
                    with col2:
                        st.info("Téléchargements individuels ↓")
                    for i, (nom, data) in enumerate(fichiers_traites):
                        st.download_button(
                            f"📥 {nom}",
                            data=data,
                            file_name=nom,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"agence_dl_{i}"
                        )
    # =========================
    # GESTION MAPPING AGENCES - Recherche / Ajout / Import / Suppression
    # =========================
    st.markdown("---")
    with st.expander("⚙️ Gestion MAPPING Agences — Recherche / Ajout / Import / Export / Suppression", expanded=False):
        _mapping_curr = get_mapping_agences()
        col_m1, col_m2, col_m3 = st.columns([2, 1, 1])
        with col_m1:
            st.metric("Agences référencées", len(_mapping_curr))
        with col_m2:
            _export_df = pd.DataFrame([{"CODE": k, "NOM": v} for k, v in sorted(_mapping_curr.items())])
            _export_buf = BytesIO()
            with pd.ExcelWriter(_export_buf, engine="xlsxwriter") as writer:
                _export_df.to_excel(writer, index=False, sheet_name="mapping")
            _export_buf.seek(0)
            st.download_button(
                "📥 Exporter mapping (Excel)",
                data=_export_buf.getvalue(),
                file_name=f"mapping_agences_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="agence_export_mapping"
            )
        with col_m3:
            if st.button("🔄 Recharger défaut", key="agence_reset_default", help="Restaure le mapping par défaut (écrase les modifs)"):
                _sync_mapping_globals(MAPPING_AGENCES_DEFAULT.copy())
                st.success("Mapping restauré par défaut")
                st.rerun()

        st.markdown("#### 🔍 Recherche agence (par code ou nom)")
        rech = st.text_input("🔍 Recherche nom ou code", placeholder="Ex: ALWATA, 5180, B8057, SRAGHNA…", key="agence_recherche_v2")
        if rech:
            _rech_lower = rech.lower().strip()
            res = [(c, n) for c, n in _mapping_curr.items() if _rech_lower in n.lower() or _rech_lower in c.lower()]
            st.caption(f"{len(res)} résultat(s) pour '{rech}'")
            if res:
                _res_display = res[:50]
                for _code, _nom in _res_display:
                    cc1, cc2, cc3 = st.columns([3, 4, 1])
                    with cc1:
                        st.code(_code, language=None)
                    with cc2:
                        st.write(_nom)
                    with cc3:
                        if st.button("🗑️", key=f"agence_del_{_code}", help=f"Supprimer {_code}"):
                            new_map = _mapping_curr.copy()
                            new_map.pop(_code, None)
                            _sync_mapping_globals(new_map)
                            st.success(f"Agence {_code} supprimée")
                            st.rerun()
                if len(res) > 50:
                    st.info(f"… et {len(res)-50} autres (affinez la recherche)")
                # CSV export filtré simple et fiable
                _filt_df = pd.DataFrame([{"CODE": c, "NOM": n} for c, n in res])
                _csv_buf = _filt_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    "📥 Télécharger résultats filtrés (CSV)",
                    data=_csv_buf,
                    file_name="recherche_agences.csv",
                    mime="text/csv",
                    key="agence_dl_filt"
                )
            else:
                st.warning("Aucun résultat — vérifiez l'orthographe ou le code")
        else:
            with st.expander("👀 Aperçu des 10 premières agences", expanded=False):
                for c, n in list(sorted(_mapping_curr.items()))[:10]:
                    st.code(f"{c} → {n}")

        st.markdown("---")
        st.markdown("#### ➕ Ajouter / Mettre à jour une agence")
        with st.form("agence_add_form", clear_on_submit=True):
            col_a1, col_a2 = st.columns(2)
            with col_a1:
                new_code = st.text_input("Code agence", placeholder="Ex: AGENT_A9999 ou BGD_B9999", key="agence_add_code")
            with col_a2:
                new_nom = st.text_input("Nom agence", placeholder="Ex: NOUVELLE AGENCE CASA", key="agence_add_nom")
            submitted = st.form_submit_button("✅ Ajouter / Mettre à jour", type="primary", use_container_width=True)
            if submitted:
                if not new_code or not new_nom:
                    st.error("Code et Nom sont obligatoires")
                else:
                    _c = new_code.strip().upper()
                    _n = new_nom.strip()
                    if _c.isdigit():
                        _c = f"AGENT_A{_c}"
                    new_map = _mapping_curr.copy()
                    is_update = _c in new_map
                    new_map[_c] = _n
                    _sync_mapping_globals(new_map)
                    st.success(f"{'Mise à jour' if is_update else 'Ajout'} : {_c} → {_n}")
                    st.rerun()
        st.markdown("---")
        st.markdown("#### 📤 Mettre à jour le mapping par fichier (Excel/CSV)")
        st.caption("À chaque fois que tu veux mettre à jour le MAPPING, upload un nouveau fichier. Colonnes attendues : **CODE** et **NOM** (ou 2 colonnes sans en-tête).")
        col_u1, col_u2 = st.columns(2)
        with col_u1:
            up_mapping = st.file_uploader(
                "Choisir fichier mapping",
                type=["xlsx", "xls", "csv"],
                key="agence_upload_mapping",
                help="Excel avec colonnes CODE | NOM ou CSV"
            )
        with col_u2:
            mode_import = st.radio(
                "Mode d'import",
                ("Ajouter / Mettre à jour (upsert)", "Remplacer tout le mapping"),
                key="agence_mode_import",
                horizontal=False
            )
        if up_mapping is not None:
            try:
                _fname = up_mapping.name.lower()
                if _fname.endswith(".csv"):
                    try:
                        _txt = up_mapping.getvalue().decode("utf-8")
                    except:
                        _txt = up_mapping.getvalue().decode("latin-1")
                    _first = _txt.split("\n")[0]
                    _sep = ";" if ";" in _first else ("," if "," in _first else "\t")
                    if _sep == "\t":
                        _sep = "\t"
                    _df_map = pd.read_csv(StringIO(_txt), dtype=str, sep=_sep)
                else:
                    _df_map = pd.read_excel(BytesIO(up_mapping.getvalue()), dtype=str)

                st.write("Aperçu du fichier uploadé :")
                st.dataframe(_df_map.head(), use_container_width=True)

                # Détection robuste CODE vs NOM (gère 'Agences' pluriel, 'CODE AGENCE', ordre inversé)
                def _norm_col(c):
                    s = str(c).strip().lower().replace('_', ' ')
                    s = re.sub(r'\s+', ' ', s)
                    return s
                _cols_norm = {_norm_col(c): c for c in _df_map.columns}
                _col_code = None
                _col_nom = None
                _code_names = {"code", "code agence", "agence code", "id", "code intermediaire", "code intermediate"}
                _nom_names = {"nom", "nom agence", "agences", "agence", "libelle", "libelle ", "raison sociale", "nom raison sociale", "agences "}
                # 'agences' (pluriel) -> NOM ; 'code agence' -> CODE (priorité au code si ambigu)
                for k_norm, v_orig in _cols_norm.items():
                    if k_norm in _code_names:
                        _col_code = v_orig
                    elif k_norm in _nom_names:
                        # 'agence' seul peut être ambigu, mais si on a déjà un code on le prend comme nom
                        if _col_nom is None:
                            _col_nom = v_orig
                # Si les deux sont détectés, c'est bon (même si ordre inversé : Agences | CODE AGENCE)
                if _col_code is None or _col_nom is None:
                    # Fallback par contenu : la colonne qui ressemble le plus à des codes (AGENT_/BGD_/4 chiffres)
                    def _score_code(col):
                        try:
                            _s = _df_map[col].dropna().astype(str).head(20)
                            _n = 0
                            for _v in _s:
                                _vu = _v.strip().upper()
                                if re.search(r'(AGENT|BGD)', _vu) or re.search(r'\d{4}', _vu):
                                    # Nom type 'BGD BOUARFA' contient aussi BGD mais sans chiffres -> exiger chiffres ou underscore
                                    if '_' in _vu or re.search(r'\d{4}', _vu):
                                        _n += 1
                            return _n
                        except:
                            return 0
                    if len(_df_map.columns) >= 2:
                        _c0, _c1 = _df_map.columns[0], _df_map.columns[1]
                        if _col_code is None and _col_nom is None:
                            # Les deux manquants : choisir par score
                            if _score_code(_c1) >= _score_code(_c0):
                                _col_code, _col_nom = _c1, _c0
                            else:
                                _col_code, _col_nom = _c0, _c1
                        elif _col_code is None:
                            # Code manquant : l'autre colonne (non-nom) est le code
                            _cands = [c for c in _df_map.columns if c != _col_nom]
                            _col_code = max(_cands, key=_score_code) if _cands else None
                        elif _col_nom is None:
                            _cands = [c for c in _df_map.columns if c != _col_code]
                            # Le nom = colonne restante (même si elle s'appelle 'Agences')
                            _col_nom = _cands[0] if _cands else None
                    else:
                        raise Exception("Fichier doit avoir au moins 2 colonnes (CODE, NOM)")
                st.caption(f"Colonnes détectées → CODE: '{_col_code}' | NOM: '{_col_nom}'")

                _new_entries = {}
                _skipped = 0
                for _, row in _df_map.iterrows():
                    _c = str(row[_col_code]).strip().upper() if pd.notna(row[_col_code]) else ""
                    _n = str(row[_col_nom]).strip() if pd.notna(row[_col_nom]) else ""
                    if not _c or _c.lower() == "nan" or not _n or _n.lower() == "nan":
                        _skipped += 1
                        continue
                    if _c.isdigit():
                        _c = f"AGENT_A{_c}"
                    _new_entries[_c] = _n

                # Détection inversion : si peu de clés contiennent des chiffres mais beaucoup de valeurs en contiennent → colonnes inversées
                if _new_entries:
                    _keys_with_digits = sum(1 for k in _new_entries.keys() if re.search(r'\d{4}', k))
                    _vals_with_digits = sum(1 for v in _new_entries.values() if re.search(r'\d{4}', str(v)))
                    if _keys_with_digits < len(_new_entries) * 0.5 and _vals_with_digits > len(_new_entries) * 0.5:
                        st.warning("⚠️ Colonnes CODE/NOM probablement inversées dans le fichier — inversion automatique appliquée. Vérifiez l'aperçu ci-dessous.")
                        _swapped = {}
                        for k, v in _new_entries.items():
                            # swap : ancienne valeur (code réel) devient clé
                            _nc = str(v).strip().upper()
                            _nn = str(k).strip()
                            if _nc.isdigit():
                                _nc = f"AGENT_A{_nc}"
                            _swapped[_nc] = _nn
                        _new_entries = _swapped
                st.info(f"{len(_new_entries)} agence(s) détectée(s) dans le fichier ({_skipped} ligne(s) ignorée(s))")
                # Vérif ciblée 6656 pour debug
                _check6656 = [f"{k} → {v}" for k, v in _new_entries.items() if '6656' in k or '6656' in str(v)]
                if _check6656:
                    st.success(f"✅ 6656 trouvé dans l'import : {_check6656[0]}")
                else:
                    st.warning("⚠️ 6656 NON trouvé dans le fichier importé — vérifiez que la ligne ASSUNOR / AGENT_A6656 existe bien.")
                if st.button(f"✅ Confirmer import ({len(_new_entries)} agences) - {mode_import}", key="agence_confirm_import", type="primary"):
                    if mode_import == "Remplacer tout le mapping":
                        final_map = _new_entries
                    else:
                        final_map = _mapping_curr.copy()
                        final_map.update(_new_entries)
                    _sync_mapping_globals(final_map)
                    st.success(f"Mapping mis à jour ! Total maintenant : {len(final_map)} agences")
                    st.rerun()
            except Exception as e:
                st.error(f"Erreur lecture fichier mapping : {e}")

        st.markdown("---")
        st.markdown("#### 🗑️ Suppression en masse / gestion")
        with st.expander("Supprimer plusieurs agences d'un coup", expanded=False):
            st.caption("Collez une liste de codes séparés par virgule, espace ou saut de ligne")
            txt_codes = st.text_area("Codes à supprimer", placeholder="Ex: AGENT_A5180, BGD_B8057\nA6666", key="agence_bulk_del")
            if st.button("🗑️ Supprimer la liste", key="agence_bulk_del_btn", type="secondary"):
                if txt_codes.strip():
                    import re as _re
                    _codes_raw = _re.split(r"[\s,;]+", txt_codes.strip())
                    _codes_norm = []
                    for _c in _codes_raw:
                        _c = _c.strip().upper()
                        if not _c:
                            continue
                        if _c.isdigit():
                            _num2c = get_mapping_numero_vers_code()
                            if _c in _num2c:
                                _c = _num2c[_c]
                            else:
                                _c = f"AGENT_A{_c}"
                        _codes_norm.append(_c)
                    new_map = _mapping_curr.copy()
                    removed = 0
                    not_found = []
                    for _c in _codes_norm:
                        if _c in new_map:
                            new_map.pop(_c)
                            removed += 1
                        else:
                            not_found.append(_c)
                    _sync_mapping_globals(new_map)
                    st.success(f"{removed} agence(s) supprimée(s)")
                    if not_found:
                        st.warning(f"Non trouvés : {', '.join(not_found[:10])}")
                    st.rerun()
                else:
                    st.error("Liste vide")

        st.markdown("**📂 Formats supportés traitement :** Excel (.xlsx), Excel ancien (.xls), CSV / TSV, Fichiers texte déguisés")
        st.caption("Version 11.0 | Traitement Agence — Mapping dynamique persistant (mapping_agences.json)")


# ==============================
# 4) APP BOA ADAPTER
#    (portage de boa-adapter.ts en Python/Streamlit)
# ==============================

# ─── Types & Constants ───────────────────────────────────────────────────────

DELIM = '|;'

# ─── Mapping préfixe fichier d'appels → préfixe OUT (pour le mode BATCH) ─────
CALL_PREFIX_TO_OUT = {
    'ACT': 'ACTIVATION',
    'ATT': 'ATTRITIONS',
    'CNA': 'CARTENACTIVE',
    'CNR': 'CARTENREMISE',
    'WLCM': 'WELCOME',
}

# ─── Parsers ─────────────────────────────────────────────────────────────────

def parse_out_svi(content: str) -> list[dict]:
    """Parse OUT_SVI file content. Returns list of records.
    
    Si le fichier ne contient que l'en-tête (vide), retourne une liste vide
    sans lever d'erreur pour permettre la génération d'un IN_SVI vide."""
    lines = [l.strip() for l in content.splitlines() if l.strip()]

    if len(lines) < 2:
        # Fichier vide (juste header) → retourne liste vide
        return []

    records = []
    for i in range(1, len(lines)):
        fields = lines[i].split(DELIM)
        if len(fields) < 5:
            continue
        records.append({
            'cleContact': fields[0].strip(),
            'telDom': fields[1].strip(),
            'telPro': fields[2].strip(),
            'telGsm': fields[3].strip(),
            'campagne': fields[4].strip(),
        })

    return records


def parse_out_report(content: str) -> list[dict]:
    """Parse OUT_REPORT file content. Returns list of rows.
    
    Si le fichier ne contient que l'en-tête (vide), retourne une liste vide
    sans lever d'erreur pour permettre la génération d'un IN_REPORT vide."""
    lines = [l.strip() for l in content.splitlines() if l.strip()]

    if len(lines) < 2:
        # Fichier vide (juste header) → retourne liste vide
        return []

    rows = []
    for i in range(1, len(lines)):
        fields = lines[i].split(DELIM)
        if len(fields) < 4:
            continue
        rows.append({
            'iteration': fields[0].strip(),
            'campagne': fields[1].strip(),
            'nombre': int(fields[2].strip()) if fields[2].strip().isdigit() else 0,
            'dateGeneration': fields[3].strip(),
        })

    return rows


# ─── IN_REPORT Generation ───────────────────────────────────────────────────

def generate_in_report(
    out_svi_records: list[dict],
    out_report_rows: list[dict],
    call_report: list[dict],
    audio_duration_seconds: float,
    treatment_date: str | None,
    filename_prefix: str,
    file_date_compact: str | None = None,
) -> dict:
    """Generate IN_REPORT content and stats."""
    audio_dur = audio_duration_seconds
    treatment_date = treatment_date or format_today_dotted()

    # ── Cas OUT_REPORT vide ──
    if not out_report_rows:
        # Fichier vide : générer uniquement l'en-tête, sans ligne de données
        header_fields = [
            'ITERATION', 'CAMPAGNE', 'NOMBRE', 'DATE_GENERATION',
            'DATE_TRAITEMENT_GoMobile',
            'NOMBRE_APPELS_TOTAL', 'NOMBRE_CONTACTS_TOTAL',
            'NOMBRE_CONTACTS_CONFORMES', 'NOMBRE_CONTACTS_NON_CONFORMES',
            'TAUX_CONFORMITE',
            'NOMBRE_DECROCHES', 'TAUX_DECROCHES',
            'NOMBRE_MESSAGE_ECOUTES', 'TAUX_ECOUTES',
            'NB_REECOUTE', 'TAUX_REECOUTE',
            'TOTAL_INTERACTIONS', 'TAUX_INTERACTION',
            'MINUTES_CONSOMMEES',
            'NB_BUTTON_PRESSE_1', 'NB_BUTTON_PRESSE_2', 'NB_BUTTON_PRESSE_3',
            'NB_BUTTON_PRESSE_4', 'NB_BUTTON_PRESSE_5', 'NB_BUTTON_PRESSE_6',
            'NB_BUTTON_PRESSE_7', 'NB_BUTTON_PRESSE_8', 'NB_BUTTON_PRESSE_9',
        ]

        content = DELIM.join(header_fields)

        # Filename avec la date du fichier source si disponible
        if file_date_compact:
            date_compact = file_date_compact
        else:
            date_match = re.search(r'(\d{8})', filename_prefix)
            date_compact = date_match.group(1) if date_match else datetime.now().strftime('%Y%m%d')
        filename = f"{filename_prefix}_IN_REPORT_{date_compact}.txt"

        return {
            'content': content,
            'filename': filename,
            'stats': {
                'totalCalls': 0,
                'contactsTotal': 0,
                'contactsConformes': 0,
                'contactsNonConformes': 0,
                'decroches': 0,
                'ecoutes': 0,
                'reecoutes': 0,
                'interactions': 0,
                'minutesConsommees': 0,
                'buttonCounts': [0] * 9,
            },
        }

    # Reference row from OUT_REPORT (first iteration)
    ref_row = out_report_rows[0].copy()
    # Si date du fichier source fournie, remplacer la date dans ref_row
    if file_date_compact:
        ref_row['dateGeneration'] = f"{file_date_compact[:4]}.{file_date_compact[4:6]}.{file_date_compact[6:8]}"
    # NOMBRE = sum across all iterations
    total_nombre = sum(r['nombre'] for r in out_report_rows)

    # ── Compute all KPIs from GoMobile call report ──
    total_calls = 0
    decroches = 0
    ecoutes = 0
    reecoutes = 0
    total_duration_seconds = 0
    button_counts = [0] * 9  # buttons 1-9
    unique_phones = set()

    for row in call_report:
        total_calls += 1

        phone = find_field(row, 'phone', 'telephone', 'tel', 'to')
        norm = normalize_to9(phone)
        if norm:
            unique_phones.add(norm)

        call_outcome = find_field(row, 'call outcome', 'call_outcome', 'outcome')
        answered_by = find_field(row, 'answered by', 'answered_by', 'answeredby')
        duration_str = find_field(row, 'duration (s)', 'duration', 'duration_s', 'durée')
        try:
            duration = float(str(duration_str).replace(',', '.')) if duration_str and str(duration_str).lower() not in ('nan', 'nat', 'none', '') else 0.0
            if math.isnan(duration):
                duration = 0.0
        except (ValueError, TypeError):
            duration = 0.0
        dtmf = find_dtmf_input(row)

        total_duration_seconds += duration

        is_answered = (
            str(call_outcome).lower() == 'completed' and
            str(answered_by).lower() == 'human'
        )

        if is_answered:
            decroches += 1

            # Écoute: audio played completely (duration >= audio duration)
            if duration >= audio_dur:
                ecoutes += 1

            # Réécoute: audio replayed (duration >= 2x audio duration)
            if duration >= audio_dur * 2:
                reecoutes += 1

            # Button presses
            if dtmf:
                for ch in str(dtmf):
                    if ch.isdigit():
                        digit = int(ch)
                        if 1 <= digit <= 9:
                            button_counts[digit - 1] += 1

    # CONFORMES = unique valid phones from GoMobile call report
    contacts_conformes = len(unique_phones)
    contacts_total = total_nombre
    contacts_non_conformes = contacts_total - contacts_conformes

    # TOTAL_INTERACTIONS = sum of all button presses
    total_interactions = sum(button_counts)

    minutes_consommees = round(total_duration_seconds / 60) if not math.isnan(total_duration_seconds) else 0

    # ── Format percentages: 2 decimals + '%' ──
    def pct(n: float, d: float) -> str:
        return f"{(n / d * 100):.2f}%" if d > 0 else "0.00%"

    # ── Build IN_REPORT content ──
    header_fields = [
        'ITERATION', 'CAMPAGNE', 'NOMBRE', 'DATE_GENERATION',
        'DATE_TRAITEMENT_GoMobile',
        'NOMBRE_APPELS_TOTAL', 'NOMBRE_CONTACTS_TOTAL',
        'NOMBRE_CONTACTS_CONFORMES', 'NOMBRE_CONTACTS_NON_CONFORMES',
        'TAUX_CONFORMITE',
        'NOMBRE_DECROCHES', 'TAUX_DECROCHES',
        'NOMBRE_MESSAGE_ECOUTES', 'TAUX_ECOUTES',
        'NB_REECOUTE', 'TAUX_REECOUTE',
        'TOTAL_INTERACTIONS', 'TAUX_INTERACTION',
        'MINUTES_CONSOMMEES',
        'NB_BUTTON_PRESSE_1', 'NB_BUTTON_PRESSE_2', 'NB_BUTTON_PRESSE_3',
        'NB_BUTTON_PRESSE_4', 'NB_BUTTON_PRESSE_5', 'NB_BUTTON_PRESSE_6',
        'NB_BUTTON_PRESSE_7', 'NB_BUTTON_PRESSE_8', 'NB_BUTTON_PRESSE_9',
    ]

    data_fields = [
        ref_row['iteration'], ref_row['campagne'], total_nombre, ref_row['dateGeneration'],
        treatment_date,
        total_calls, total_nombre,
        contacts_conformes, contacts_non_conformes,
        pct(contacts_conformes, contacts_total),
        decroches, pct(decroches, contacts_conformes),
        ecoutes, pct(ecoutes, decroches),
        reecoutes, pct(reecoutes, decroches),
        total_interactions, pct(total_interactions, decroches),
        minutes_consommees,
        *button_counts,
    ]

    content = DELIM.join(header_fields) + '\n' + DELIM.join(str(f) for f in data_fields)

    # Filename
    if file_date_compact:
        date_compact = file_date_compact
    else:
        date_compact = ref_row['dateGeneration'].replace('.', '')
    filename = f"{filename_prefix}_IN_REPORT_{date_compact}.txt"

    return {
        'content': content,
        'filename': filename,
        'stats': {
            'totalCalls': total_calls,
            'contactsTotal': contacts_total,
            'contactsConformes': contacts_conformes,
            'contactsNonConformes': contacts_non_conformes,
            'decroches': decroches,
            'ecoutes': ecoutes,
            'reecoutes': reecoutes,
            'interactions': total_interactions,
            'minutesConsommees': minutes_consommees,
            'buttonCounts': button_counts,
        },
    }


# ─── IN_SVI Generation ──────────────────────────────────────────────────────

def generate_in_svi(
    out_svi_records: list[dict],
    call_report: list[dict],
    date_compact: str,
    filename_prefix: str,
    file_date_compact: str | None = None,
) -> dict:
    """Generate IN_SVI content (only contacts who pressed DTMF)."""
    # Date effective à utiliser (source si disponible, sinon paramètre)
    actual_date_compact = file_date_compact if file_date_compact else date_compact

    # Si OUT_SVI vide → IN_SVI vide aussi
    if not out_svi_records:
        filename = f"{filename_prefix}_IN_SVI_{actual_date_compact}.txt"
        return {
            'content': '',
            'filename': filename,
            'recordCount': 0,
        }

    # Build lookups
    cle_to_record = {rec['cleContact']: rec for rec in out_svi_records}

    phone_to_record = {}
    for rec in out_svi_records:
        for tel in [rec['telGsm'], rec['telDom'], rec['telPro']]:
            norm = normalize_to9(tel)
            if norm:
                phone_to_record[norm] = rec

    svi_lines = []
    seen_cles = set()

    for row in call_report:
        call_outcome = find_field(row, 'call outcome', 'call_outcome', 'outcome')
        answered_by = find_field(row, 'answered by', 'answered_by', 'answeredby')
        dtmf = find_dtmf_input(row)
        phone = find_field(row, 'phone', 'telephone', 'tel', 'to')

        is_answered = (
            str(call_outcome).lower() == 'completed' and
            str(answered_by).lower() == 'human'
        )

        if not is_answered or not dtmf:
            continue

        cle = find_field(row, 'cle_contact', 'clecontact')
        phone = find_field(row, 'phone', 'telephone', 'tel', 'to')
        phone_norm = normalize_to9(phone)

        if cle and cle in cle_to_record:
            original = cle_to_record[cle]
            if cle in seen_cles:
                continue
            seen_cles.add(cle)
        elif phone_norm and phone_norm in phone_to_record:
            original = phone_to_record[phone_norm]
            cle_dup = original['cleContact']
            if cle_dup in seen_cles:
                continue
            seen_cles.add(cle_dup)
        else:
            continue

        line = DELIM.join([
            original['cleContact'],
            original['telDom'],
            original['telPro'],
            original['telGsm'],
            original['campagne'],
            actual_date_compact,
        ])
        svi_lines.append(line)

    filename = f"{filename_prefix}_IN_SVI_{date_compact}.txt"

    return {
        'content': '\n'.join(svi_lines),
        'filename': filename,
        'recordCount': len(svi_lines),
    }


# ─── Helpers ─────────────────────────────────────────────────────────────────

def extract_filename_prefix(filename: str) -> str:
    """Extract prefix from OUT file name (e.g. 'CARTENACTIVE_OUT_SVI_20260215.txt' → 'CARTENACTIVE')."""
    base = re.sub(r'\.(txt|csv)$', '', filename, flags=re.IGNORECASE)
    match = re.match(r'^(.+?)_OUT_(?:SVI|REPORT)_', base, re.IGNORECASE)
    return match.group(1) if match else base


def normalize_to9(phone: str) -> str:
    """Normalize any phone format to last 9 digits for matching."""
    digits = re.sub(r'\D', '', str(phone))
    if len(digits) >= 9:
        return digits[-9:]
    return digits


def find_field(record: dict, *keys: str) -> str:
    """Find a value in a record by trying multiple column names (case-insensitive)."""
    for key in keys:
        norm = key.lower().strip()
        for k, v in record.items():
            if k.lower().strip() == norm and v is not None and str(v) != '':
                return str(v)
    return ''


def find_dtmf_input(record: dict) -> str:
    """Find DTMF input from call export columns.
    Recherche dans les colonnes nommées : input, dtmf, digit, key, touche, saisie, chiffre, reponse.
    Retourne tout DTMF non vide (chiffres 0-9)."""
    dtmf_keywords = ['input', 'dtmf', 'digit', 'key', 'touche', 'saisie', 'chiffre', 'reponse']
    for key, value in record.items():
        k = key.lower().replace(' ', '').replace('_', '').replace(':', '')
        if any(word in k for word in dtmf_keywords) and 'branch' not in k and value is not None and str(value) != '':
            val = str(value).strip()
            if any(ch.isdigit() for ch in val):
                return val
    return ''


def format_today_dotted() -> str:
    """Format today's date as YYYY.MM.DD."""
    d = datetime.now()
    return f"{d.year}.{d.month:02d}.{d.day:02d}"


# ─── Streamlit App BOA Adapter (Batch) ───────────────────────────────────────

def load_call_records(uploaded_file) -> list[dict]:
    """Charge un fichier d'appels GoMobile (CSV/Excel) en liste de dicts."""
    if uploaded_file.name.lower().endswith('.csv'):
        df = pd.read_csv(uploaded_file, dtype=str)
    else:
        df = pd.read_excel(uploaded_file, dtype=str)
    return df.to_dict('records')


def infer_out_prefix_from_call_filename(filename: str) -> str:
    """Devine le préfixe OUT depuis le nom du fichier d'appels (ex: WLCM11082026.csv → WELCOME)."""
    m = re.match(r'^([A-Za-z]+)', filename)
    pfx = m.group(1).upper() if m else ''
    return CALL_PREFIX_TO_OUT.get(pfx, pfx)


def app_boa_adapter_batch(audio_duration, treatment_date):
    """Mode BATCH : plusieurs campagnes → ZIP de tous les IN_REPORT + IN_SVI + tableau récap."""

    # ── Étape 1 : OUT fichiers ──
    st.subheader("📁 Étape 1 : Déposer tous les fichiers OUT (.txt)")
    uploaded_txt = st.file_uploader(
        "Glisser-déposer TOUS les OUT_SVI + OUT_REPORT",
        type=["txt"],
        accept_multiple_files=True,
        key="boa_batch_txt"
    )

    out_by_prefix = {}
    if uploaded_txt:
        for f in uploaded_txt:
            content = f.read().decode('utf-8', errors='replace')
            base = f.name
            prefix = extract_filename_prefix(base)
            out_by_prefix.setdefault(prefix, {})
            if 'OUT_SVI' in base.upper():
                out_by_prefix[prefix]['SVI'] = content
            elif 'OUT_REPORT' in base.upper():
                out_by_prefix[prefix]['REPORT'] = content

    # ── Étape 2 : Call exports ──
    st.subheader("📞 Étape 2 : Déposer les fichiers d'appels GoMobile")
    uploaded_calls = st.file_uploader(
        "Glisser-déposer TOUS les call exports (CSV/Excel)",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=True,
        key="boa_batch_calls"
    )

    call_by_prefix = {}
    if uploaded_calls:
        for f in uploaded_calls:
            try:
                out_prefix = infer_out_prefix_from_call_filename(f.name)
                call_by_prefix[out_prefix] = load_call_records(f)
            except Exception as e:
                st.error(f"❌ Erreur lecture {f.name} : {e}")

    # Campagnes prêtes (SVI + REPORT présents)
    ready = []
    for prefix, files in out_by_prefix.items():
        if 'SVI' in files and 'REPORT' in files:
            ready.append({
                'prefix': prefix,
                'svi_content': files['SVI'],
                'report_content': files['REPORT'],
                'calls': call_by_prefix.get(prefix, []),
            })
        else:
            st.warning(f"⚠️ Campagne {prefix} : manque OUT_SVI ou OUT_REPORT, ignorée.")

    st.caption(f"**{len(ready)}** campagne(s) appariée(s) · {sum(len(c['calls']) for c in ready)} appels chargés")

    # ── Étape 3 : Lancement ──
    if st.button("🚀 Lancer la génération batch", disabled=not ready, key="boa_batch_run"):
        with st.spinner("Traitement de toutes les campagnes..."):
            # Valider date traitement
            treatment_date_val = treatment_date.strip() if treatment_date else None
            if treatment_date_val and not re.match(r'^\d{4}\.\d{2}\.\d{2}$', treatment_date_val):
                st.warning("Format date invalide, utilisation de la date du jour.")
                treatment_date_val = None

            zip_buffer = io.BytesIO()
            summary_rows = []
            results = []

            # Date du ZIP = date des fichiers source (1er OUT_REPORT), pas la date du jour
            batch_date_compact = None

            # Première passe : récupérer la date de référence depuis les OUT_REPORT non vides
            reference_date_compact = None
            for c in ready:
                out_report_rows = parse_out_report(c['report_content'])
                if out_report_rows:
                    reference_date_compact = out_report_rows[0]['dateGeneration'].replace('.', '')
                    break
            
            # Fallback : si tous les OUT_REPORT sont vides, utiliser la date du jour
            if reference_date_compact is None:
                reference_date_compact = datetime.now().strftime('%Y%m%d')

            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                for c in ready:
                    prefix = c['prefix']
                    out_svi_records = parse_out_svi(c['svi_content'])
                    out_report_rows = parse_out_report(c['report_content'])
                    calls = c['calls']

                    # Date compacte : utiliser la date du OUT_REPORT s'il n'est pas vide, sinon la date de référence
                    if out_report_rows:
                        date_compact = out_report_rows[0]['dateGeneration'].replace('.', '')
                    else:
                        date_compact = reference_date_compact

                    if batch_date_compact is None:
                        batch_date_compact = date_compact

                    in_report = generate_in_report(
                        out_svi_records, out_report_rows, calls,
                        audio_duration, treatment_date_val, prefix, date_compact,
                    )
                    in_svi = generate_in_svi(
                        out_svi_records, calls, date_compact, prefix, date_compact,
                    )

                    zf.writestr(in_report['filename'], in_report['content'].encode('utf-8'))
                    zf.writestr(in_svi['filename'], in_svi['content'].encode('utf-8'))

                    s = in_report['stats']
                    summary_rows.append({
                        'Campagne': prefix,
                        'IN_SVI lignes DTMF': in_svi['recordCount'],
                        'Appels': s['totalCalls'],
                        'Contacts total': s['contactsTotal'],
                        'Décrochés': s['decroches'],
                        'Écoutes': s['ecoutes'],
                        'Interactions (DTMF)': s['interactions'],
                    })
                    results.append((prefix, in_report, in_svi))

            zip_buffer.seek(0)

        st.success("✅ Batch terminé avec succès !")

        # ── Tableau récapitulatif DTMF ──
        st.subheader("📊 Tableau récapitulatif (Campagne → lignes DTMF IN_SVI)")
        st.dataframe(pd.DataFrame(summary_rows), width='stretch')

        # ── Téléchargement ZIP global ──
        zip_file_date = batch_date_compact if batch_date_compact else datetime.now().strftime('%Y%m%d')
        st.download_button(
            "📦 Télécharger le ZIP (tous les IN_REPORT + IN_SVI)",
            data=zip_buffer,
            file_name=f"BOA_IN_FILES_{zip_file_date}.zip",
            mime="application/zip",
            key="boa_batch_zip"
        )

        # ── Aperçus par campagne ──
        st.subheader("👀 Aperçus par campagne")
        for prefix, in_report, in_svi in results:
            with st.expander(f"{prefix} — IN_SVI ({in_svi['recordCount']} lignes DTMF)"):
                st.markdown(f"**{in_report['filename']}** (IN_REPORT)")
                st.text(in_report['content'])
                st.markdown(f"**{in_svi['filename']}** (IN_SVI)")
                if in_svi['recordCount'] > 0:
                    st.text(in_svi['content'])
                else:
                    st.info("Aucun contact n'a pressé de touche DTMF.")


# ─── Streamlit App BOA Adapter ───────────────────────────────────────────────

def app_boa_adapter():
    st.title("🔗 BOA Adapter – Génération IN_REPORT & IN_SVI")
    

    # ── Paramètres dans la sidebar ──
    with st.sidebar:
        st.header("⚙️ Paramètres GoMobile")
        audio_duration = st.number_input(
            "Durée audio (secondes)",
            min_value=1.0, max_value=250.0, value=25.0, step=1.0,
            key="boa_audio_dur"
        )
        treatment_date = st.text_input(
            "Date traitement GoMobile (YYYY.MM.DD, laisser vide = aujourd'hui)",
            value="",
            key="boa_treatment_date"
        )
        st.caption("Format: 2026.05.21")
        st.markdown("---")
        boamode = st.radio(
            "Mode de traitement",
            ("Simple (1 campagne)", "Batch (multi-campagnes)"),
            key="boa_mode"
        )

    # ── Si mode Batch → délégation ──
    if boamode == "Batch (multi-campagnes)":
        app_boa_adapter_batch(audio_duration, treatment_date)
        return


    # ── Étape 1: Upload multi-fichiers .txt ──
    st.subheader("📁 Étape 1 : Déposer les fichiers BOA (.txt)")

    uploaded_files = st.file_uploader(
        "Glisser-déposer les fichiers .txt",
        type=["txt"],
        accept_multiple_files=True,
        key="boa_txt_upload"
    )

    out_svi_files = []
    out_report_files = []

    if uploaded_files is not None and len(uploaded_files) > 0:
        for uploaded_file in uploaded_files:
            content = uploaded_file.read().decode('utf-8', errors='replace')
            base = uploaded_file.name
            if 'OUT_SVI' in base.upper():
                out_svi_files.append({'name': base, 'content': content})
            elif 'OUT_REPORT' in base.upper():
                out_report_files.append({'name': base, 'content': content})

       

    st.subheader("📋 Étape 2 : Sélectionner les fichiers à traiter")

    selected_svi = None
    selected_report = None

    if out_svi_files and out_report_files:
        col_svi, col_report = st.columns(2)

        with col_svi:
            svi_names = [f["name"] for f in out_svi_files]
            selected_svi_name = st.selectbox(
                "BOA OUT_SVI File (.txt)",
                options=svi_names,
                key="boa_svi_select"
            )
            selected_svi = next(f for f in out_svi_files if f["name"] == selected_svi_name)

        with col_report:
            report_names = [f["name"] for f in out_report_files]
            selected_report_name = st.selectbox(
                "BOA OUT_REPORT File (.txt)",
                options=report_names,
                key="boa_report_select"
            )
            selected_report = next(f for f in out_report_files if f["name"] == selected_report_name)

    # ── Étape 3: GoMobile Call Export (CSV/Excel) ──
    st.subheader("📞 Étape 3 : GoMobile Call Export (CSV/Excel)")
    st.markdown("Fichier d'appels exporté depuis GoMobile (format CSV ou Excel).")

    call_export_file = st.file_uploader(
        "Glisser-déposer le fichier d'appels GoMobile",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=False,
        key="boa_call_export"
    )

    call_report_data = None
    if call_export_file is not None:
        try:
            if call_export_file.name.lower().endswith('.csv'):
                call_report_data = pd.read_csv(call_export_file, dtype=str).to_dict('records')
            else:
                call_report_data = pd.read_excel(call_export_file, dtype=str).to_dict('records')
            st.success(f"✅ {len(call_report_data)} appels chargés depuis {call_export_file.name}")
        except Exception as e:
            st.error(f"❌ Erreur lecture fichier d'appels : {e}")

    # Debug: afficher les colonnes du fichier d'appels (après chargement)
    if call_report_data and len(call_report_data) > 0:
        with st.expander("🔍 Debug : Colonnes du fichier d'appels GoMobile"):
            cols = list(call_report_data[0].keys())
            st.write("**Colonnes détectées :**")
            st.code('\n'.join(cols))
            first_row = call_report_data[0]
            dtmf_col = None
            for k in first_row:
                kl = k.lower().replace(' ', '').replace('_', '').replace(':', '')
                dtmf_keywords = ['input', 'dtmf', 'digit', 'key', 'touche', 'saisie', 'chiffre', 'reponse']
                if any(w in kl for w in dtmf_keywords) and 'branch' not in kl:
                    dtmf_col = k
                    break
            if dtmf_col:
                st.success(f"✅ Colonne DTMF détectée : **{dtmf_col}**")
                sample_val = first_row.get(dtmf_col, '')
                st.write(f"Valeur exemple : `{sample_val}`")
            else:
                st.error("❌ Aucune colonne DTMF détectée. Vérifie le nom des colonnes ci-dessus.")

    # ── Étape 4: Traitement ──
    st.subheader("🚀 Étape 4 : Génération des fichiers IN")

    # ── Détection fichiers vides ──
    svi_empty = False
    report_empty = False
    if selected_svi is not None:
        try:
            svi_test = parse_out_svi(selected_svi['content'])
            svi_empty = len(svi_test) == 0
        except Exception:
            svi_empty = True
    if selected_report is not None:
        try:
            report_test = parse_out_report(selected_report['content'])
            report_empty = len(report_test) == 0
        except Exception:
            report_empty = True

    both_empty = svi_empty and report_empty

    can_process = (
        selected_svi is not None and
        selected_report is not None and
        (call_report_data is not None or both_empty)
    )

    

    if both_empty:
        st.info("ℹ️ Les fichiers OUT_SVI et OUT_REPORT sont vides — la génération produira des fichiers IN vides (headers seuls).")

    if st.button("▶️ Lancer la génération IN_REPORT + IN_SVI", disabled=not can_process, key="boa_run"):
        try:
            with st.spinner("Traitement en cours..."):
                # Parse
                out_svi_records = parse_out_svi(selected_svi['content'])
                out_report_rows = parse_out_report(selected_report['content'])

                # Extract prefix
                filename_prefix = extract_filename_prefix(selected_svi['name'])

                # Extract date from source filename (e.g. WELCOME_OUT_SVI_20260521.txt → 20260521)
                source_date_match = re.search(r'(\d{8})', selected_svi['name'])
                source_date_compact = source_date_match.group(1) if source_date_match else None

                # Validate treatment date
                treatment_date_val = treatment_date.strip() if treatment_date else None
                if treatment_date_val and not re.match(r'^\d{4}\.\d{2}\.\d{2}$', treatment_date_val):
                    st.warning("Format date invalide, utilisation de la date du jour.")
                    treatment_date_val = None

                # Si pas de Call Export mais fichiers vides → utiliser liste vide
                call_data = call_report_data if call_report_data is not None else []

                # Generate IN_REPORT
                in_report = generate_in_report(
                    out_svi_records,
                    out_report_rows,
                    call_data,
                    audio_duration,
                    treatment_date_val,
                    filename_prefix,
                    source_date_compact,
                )

                # Generate IN_SVI
                if out_report_rows:
                    date_compact = out_report_rows[0]['dateGeneration'].replace('.', '')
                else:
                    date_match = re.search(r'(\d{8})', filename_prefix)
                    date_compact = date_match.group(1) if date_match else datetime.now().strftime('%Y%m%d')
                in_svi = generate_in_svi(
                    out_svi_records,
                    call_data,
                    date_compact,
                    filename_prefix,
                    source_date_compact,
                )

            # ── Affichage des résultats ──
            st.success("✅ Génération terminée avec succès !")

            # Stats
            st.subheader("📊 Statistiques IN_REPORT")
            stats = in_report['stats']

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Appels totaux", stats['totalCalls'])
            c2.metric("Contacts total", stats['contactsTotal'])
            c3.metric("Contacts conformes", stats['contactsConformes'])
            c4.metric("Contacts non conformes", stats['contactsNonConformes'])

            c5, c6, c7, c8 = st.columns(4)
            c5.metric("Décrochés", stats['decroches'])
            c6.metric("Écoutes", stats['ecoutes'])
            c7.metric("Réécoutes", stats['reecoutes'])
            c8.metric("Interactions (DTMF)", stats['interactions'])

            st.metric("Minutes consommées", stats['minutesConsommees'])

            # Buttons pressés
            st.subheader("🔘 Compteurs de touches DTMF")
            btn_cols = st.columns(9)
            for i, count in enumerate(stats['buttonCounts']):
                btn_cols[i].metric(f"Touche {i+1}", count)

            # ── Téléchargements ──
            st.subheader("📥 Téléchargements")

            col_dl1, col_dl2 = st.columns(2)

            with col_dl1:
                st.download_button(
                    label=f"⬇️ Télécharger {in_report['filename']}",
                    data=in_report['content'].encode('utf-8'),
                    file_name=in_report['filename'],
                    mime="text/plain",
                    key="boa_dl_in_report"
                )

            with col_dl2:
                st.download_button(
                    label=f"⬇️ Télécharger {in_svi['filename']}",
                    data=in_svi['content'].encode('utf-8'),
                    file_name=in_svi['filename'],
                    mime="text/plain",
                    key="boa_dl_in_svi"
                )

            # ── Aperçus ──
            with st.expander("👀 Aperçu IN_REPORT"):
                st.text(in_report['content'])

            with st.expander(f"👀 Aperçu IN_SVI ({in_svi['recordCount']} lignes)"):
                if in_svi['recordCount'] > 0:
                    preview_lines = in_svi['content'].split('\n')[:10]
                    st.text('\n'.join(preview_lines))
                    if in_svi['recordCount'] > 10:
                        st.caption(f"... et {in_svi['recordCount'] - 10} lignes supplémentaires")
                else:
                    st.info("Aucun contact n'a pressé de touche DTMF.")

        except Exception as e:
            st.error(f"❌ Erreur lors du traitement : {e}")
            st.exception(e)

    st.markdown("---")
    st.caption(
        "Logique : parse OUT_SVI/OUT_REPORT → croisement avec Call Export GoMobile → "
        "KPIs (décrochés, écoutes, réécoutes, DTMF) → génération IN_REPORT + IN_SVI."
    )


# ==============================
# ROUTEUR PRINCIPAL (MENU)
# ==============================

# Injecter le style global (thème light bleu/blanc)
inject_css()

LOGO_URL = (
    "https://media.licdn.com/dms/image/v2/D4E0BAQHd1vQ5srIY4w/company-logo_200_200/"
    "B4EZZdUUMCHMAI-/0/1745322330921/gomobile_africa_logo?e=2147483647&v=beta&t=h3LJTeBhImOortH_t5PmBxDMwzEi3vyIelylBx9lKuU"
)

# Bannière de marque en haut de la sidebar
st.sidebar.markdown(
    f"""
    <div class="side-brand">
        <img src="{LOGO_URL}" alt="logo" onerror="this.style.display='none'">
        <div class="name">GoMobile Tools</div>
        <div class="sub">Suite Téléphonie &amp; Marketing</div>
    </div>
    """,
    unsafe_allow_html=True,
)

app_choice = st.sidebar.radio(
    "Choisir l'application :",
    (
        "EQDOM_MARKETING",
        "TRAITEMENT_AGENCE",
        "BOA_MARKETING",
        "BOA_REPORT_GENERATOR",
    ),
    key="main_app_choice"
)

if app_choice == "EQDOM_MARKETING":
    app_hero("EQDOM · Marketing", "Normalisation & déduplication de fichiers Excel marketing")
    app_eqdom_marketing()
elif app_choice == "TRAITEMENT_AGENCE":
    app_hero("Traitement Fichiers Agence", "Détection automatique agence & mise en forme ASSURCALL")
    app_traitement_agence()
elif app_choice == "BOA_MARKETING":
    app_hero("AVT → APT · Nettoyage", "Normalisation de fichiers GoMobile (AVT vers APT)")
    app_avt_to_apt()
else:
    app_hero("Génération IN_REPORT & IN_SVI", "Croisement des exports GoMobile pour générer les fichiers IN")
    app_boa_adapter()
