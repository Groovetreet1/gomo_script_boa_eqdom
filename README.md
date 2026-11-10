# GoMobile Tools — Suite Outils Téléphonie / Marketing

Application **Streamlit** regroupant 5 outils sous le même menu :

1. **EQDOM · Marketing** — Normalisation & déduplication de fichiers Excel marketing (`ID_CLIENT_INTERNE` → `user_identify`, `CLIENT_TEL_VALIDE` → `telephone`, filtre numéros `05/06/07`).
2. **AVT → APT · Nettoyage** — Normalisation de fichiers GoMobile (remplissage `TEL_GSM`, `DOM/PRO` → `NA`, filtres Maroc, déduplication, export cadrillé).
3. **BOA · Génération IN_REPORT & IN_SVI** — Croisement des exports GoMobile (`OUT_SVI` + `OUT_REPORT` + Call Export) pour générer les fichiers `IN_REPORT` / `IN_SVI`, en mode **Simple (1 campagne)** ou **Batch (multi-campagnes → ZIP + tableau DTMF)**.
4. **Correcteur de Numéros** — Ajout du préfixe `0` pour WS MRKG.
5. **AUTOMATISATION_CAMPAGNES** — Création des audiences, import des contacts `TEL_GSM` et lancement des campagnes GoMobile (T1 10:00, T2 14:00, T3 16:30 le jour du traitement).

---

## 🚀 Automatisation des campagnes GoMobile

L'onglet **AUTOMATISATION_CAMPAGNES** (ou **BOA_MARKETING → 🚀 Audiences & Campagnes GoMobile**) permet de :

- se connecter avec une **clé API GoMobile** (`X-API-Key`) — la clé reste en mémoire de session et n'est jamais écrite sur disque ;
- créer ou réutiliser les audiences nommées `NOM_LISTE JJMMAAAA` (date de traitement modifiable) ;
- importer les contacts via la colonne **`TEL_GSM`** ;
- créer et activer les campagnes portant le même nom que l'audience, avec 3 tentatives le jour du traitement : **10:00**, **14:00** et **16:30**.

Les relances T2/T3 ne sont planifiées que pour les flows **Voice**. Les listes traitées dans l'onglet **BOA_MARKETING** sont directement réutilisables dans l'onglet d'automatisation.

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
├── gomobile.py            # Application principale (5 outils)
├── gomobile_api.py        # Client API GoMobile + helpers campagnes
├── requirements.txt       # Dépendances Python
├── render.yaml            # Config Render (Blueprint)
├── .streamlit/config.toml # Thème (light bleu/blanc)
└── .gitignore
```