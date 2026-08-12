# GoMobile Tools — Suite Outils Téléphonie / Marketing

Application **Streamlit** regroupant 4 outils sous le même menu :

1. **EQDOM · Marketing** — Normalisation & déduplication de fichiers Excel marketing (`ID_CLIENT_INTERNE` → `user_identify`, `CLIENT_TEL_VALIDE` → `telephone`, filtre numéros `05/06/07`).
2. **AVT → APT · Nettoyage** — Normalisation de fichiers GoMobile (remplissage `TEL_GSM`, `DOM/PRO` → `NA`, filtres Maroc, déduplication, export cadrillé).
3. **BOA · Génération IN_REPORT & IN_SVI** — Croisement des exports GoMobile (`OUT_SVI` + `OUT_REPORT` + Call Export) pour générer les fichiers `IN_REPORT` / `IN_SVI`, en mode **Simple (1 campagne)** ou **Batch (multi-campagnes → ZIP + tableau DTMF)**.
4. **Correcteur de Numéros** — Ajout du préfixe `0` pour WS MRKG.

---

## 🚀 Déploiement sur Render

L'application est prête pour **Render** (Blueprint). Deux options :

### Option A — Deploy automatique (Blueprprint)

1. Pousse ce dépôt sur **GitHub**.
2. Dans [render.com](https://render.com) → **New → Blueprint**.
3. Connecte le dépôt → Render lit `render.yaml` et déploie automatiquement.

### Option B — Deploy Web Service (manuel)

1. **New → Web Service** → connecter le repo GitHub.
2. **Build Command**:
   ```
   pip install -r requirements.txt
   ```
3. **Start Command**:
   ```
   streamlit run gomobile.py --server.port 10000 --server.address 0.0.0.0
   ```
4. Choisir une instance **Free** (ou Starter), puis **Deploy**.

> 💡 Note : sur le plan Free, Render met le service en veille après inactivité — il se réveille sur la 1ère visite (quelques secondes).

---

## 🧪 Exécution locale

```bash
pip install -r requirements.txt
streamlit run gomobile.py
```

Ouvre ensuite [http://localhost:8501](http://localhost:8501).

---

## 📁 Structure

```
.
├── gomobile.py            # Application principale (4 outils)
├── requirements.txt       # Dépendances Python
├── render.yaml            # Config Render (Blueprint)
├── .streamlit/config.toml # Thème (light bleu/blanc)
└── .gitignore
```