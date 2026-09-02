# ============================================================
# VISION EDITION — streamlit_app.py
# Version 2.0 — Multi-dossiers / Accès dirigeant / Rôles
# © 2025 Nicolas CUISSET — Mémoire d'expertise comptable
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import json
import hashlib
import datetime as _dt
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go
import qrcode
from PIL import Image
import anthropic

try:
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_ALIGN_VERTICAL
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

APP_URL = "https://outilaccompagnementmaisonsedition-lyvgltfbwtqo4m9tdmzofu.streamlit.app/"

# ============================================================
# HELPERS FORMAT FR
# ============================================================
def fmt_fr(x, decimals=0):
    try:
        s = f"{float(x):,.{decimals}f}"
    except (ValueError, TypeError):
        return str(x)
    return s.replace(",", "§").replace(".", ",").replace("§", " ")

# ============================================================
# GESTION DES UTILISATEURS ET DOSSIERS (JSON session)
# ============================================================
# Structure en session_state["cabinet"] :
# {
#   "ec_users": { "login": {"pw_hash": ..., "name": ..., "role": "ec"} },
#   "dossiers": {
#     "id_dossier": {
#       "nom": "Editions de l'Argonaute",
#       "dirigeant": "Thomas Bernard",
#       "exercice": "2025/2026",
#       "param_comptes": {...},
#       "mapping": {...},
#       "labels_indirect": {...},
#       "dirigeant_users": { "login": {"pw_hash": ..., "name": ...} }
#     }
#   }
# }

def init_cabinet():
    """Initialise la structure cabinet si absente."""
    if "cabinet" not in st.session_state:
        st.session_state["cabinet"] = {
            "ec_users": {
                "aurore": {
                    "pw_hash": _hash("12345"),
                    "name": "Aurore Demoulin",
                    "role": "ec"
                }
            },
            "dossiers": {}
        }

def _hash(pw):
    return hashlib.sha256(pw.encode()).hexdigest()

def verifier_login(username, password, role_attendu=None):
    """Vérifie un login EC ou dirigeant selon le dossier actif."""
    init_cabinet()
    cab = st.session_state["cabinet"]
    # Vérif EC
    ec = cab["ec_users"].get(username)
    if ec and ec["pw_hash"] == _hash(password):
        if role_attendu in (None, "ec"):
            return {"ok": True, "role": "ec", "name": ec["name"], "dossier_id": None}
    # Vérif dirigeant dans tous les dossiers
    for did, dos in cab["dossiers"].items():
        du = dos.get("dirigeant_users", {}).get(username)
        if du and du["pw_hash"] == _hash(password):
            return {"ok": True, "role": "dirigeant", "name": du["name"],
                    "dossier_id": did, "dossier_nom": dos["nom"]}
    return {"ok": False}

def get_dossier(dossier_id):
    init_cabinet()
    return st.session_state["cabinet"]["dossiers"].get(dossier_id, {})

def save_dossier(dossier_id, data):
    init_cabinet()
    st.session_state["cabinet"]["dossiers"][dossier_id] = data

# ============================================================
# CONFIG PAGE
# ============================================================
st.set_page_config(
    page_title="VISION EDITION",
    page_icon="📚",
    layout="wide"
)

# ============================================================
# INITIALISATION SESSION
# ============================================================
init_cabinet()
for k, v in [
    ("login", False), ("role", None), ("username", ""), ("name", ""),
    ("dossier_id", None), ("messages_agent", []), ("mode_anonyme", False)
]:
    if k not in st.session_state:
        st.session_state[k] = v

# ============================================================
# AUTHENTIFICATION
# ============================================================
if not st.session_state["login"]:
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("## 📚 VISION EDITION")
        st.markdown("*Pilotage analytique — Maisons d'édition indépendantes*")
        st.divider()
        username_input = st.text_input("Identifiant", placeholder="Votre identifiant")
        password_input = st.text_input("Mot de passe", type="password")
        if st.button("🔑 Connexion", use_container_width=True, type="primary"):
            res = verifier_login(username_input, password_input)
            if res["ok"]:
                st.session_state["login"] = True
                st.session_state["role"] = res["role"]
                st.session_state["username"] = username_input
                st.session_state["name"] = res["name"]
                st.session_state["dossier_id"] = res.get("dossier_id")
                st.rerun()
            else:
                st.error("❌ Identifiants incorrects")
        st.divider()
        st.markdown("**Accès mobile**")
        qr = qrcode.QRCode(version=1, box_size=4, border=2)
        qr.add_data(APP_URL)
        qr.make(fit=True)
        img_qr = qr.make_image(fill_color="black", back_color="white")
        buf = BytesIO(); img_qr.save(buf, format="PNG")
        st.image(buf.getvalue(), width=120)
    st.stop()

# ============================================================
# SIDEBAR SELON RÔLE
# ============================================================
role = st.session_state["role"]

with st.sidebar:
    st.markdown(f"👤 **{st.session_state['name']}**")
    if role == "ec":
        st.caption("Expert-comptable")
    elif role == "dirigeant":
        dossier_actif = get_dossier(st.session_state["dossier_id"])
        st.caption(f"Dirigeant — {dossier_actif.get('nom', '')}")
    st.divider()

    if role == "ec":
        # Sélection dossier
        cab = st.session_state["cabinet"]
        dossiers = cab["dossiers"]
        if dossiers:
            options_dos = {did: d["nom"] for did, d in dossiers.items()}
            did_sel = st.selectbox(
                "📁 Dossier client",
                options=list(options_dos.keys()),
                format_func=lambda x: options_dos[x],
                key="sidebar_dossier_sel"
            )
            if st.session_state["dossier_id"] != did_sel:
                st.session_state["dossier_id"] = did_sel
                # Réinitialiser les données du dossier précédent
                for k in ["df_pivot", "df_pivot_brut", "df_comptables", "param_comptes",
                           "df_source_mappe", "repartition_active"]:
                    st.session_state.pop(k, None)
                st.rerun()
        else:
            st.info("Aucun dossier. Créez-en un.")
            st.session_state["dossier_id"] = None

        pages_ec = [
            "🏠 Accueil",
            "👥 Gestion des dossiers",
            "📂 Import des données",
            "⚙️ Paramétrage analytique",
            "📈 Tableau de bord éditorial",
            "📖 Analyse par titre",
            "🎯 Simulateur de rentabilité",
            "💰 Trésorerie prévisionnelle",
            "✍️ Droits d'auteurs",
            "📦 Retours & Remises",
            "📊 Synthèse financière",
            "🤖 Assistant IA",
            "📄 Rapport de pilotage"
        ]
        page = st.selectbox("Navigation", pages_ec)
        st.divider()
        st.session_state["mode_anonyme"] = st.checkbox(
            "🕶️ Mode démonstration", value=st.session_state.get("mode_anonyme", False)
        )

    elif role == "dirigeant":
        pages_dir = [
            "🏠 Mon tableau de bord",
            "📖 Mes titres",
            "💰 Ma trésorerie",
            "📦 Retours & Remises",
            "✍️ Droits d'auteurs",
            "🤖 Mon assistant IA"
        ]
        page = st.selectbox("Navigation", pages_dir)

    st.divider()
    if st.button("🚪 Déconnexion", use_container_width=True):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

# ============================================================
# HELPERS ANALYTIQUE (identiques à v1 — extraits clés)
# ============================================================
COMPTE_CHARGES_INDIRECTES_REPARTIES = "CHARGES INDIRECTES REPARTIES"
COMPTE_PRODUITS_INDIRECTS_REPARTIS  = "PRODUITS INDIRECTS REPARTIS"

def filtrer_isbn_reels(df):
    if "Code_Analytique" not in df.columns:
        return df.iloc[0:0]
    label_ci = st.session_state.get("labels_indirect", {}).get("charges", "CHARGES INDIRECTES")
    label_pi = st.session_state.get("labels_indirect", {}).get("produits", "PRODUITS INDIRECTS")
    labels_exclus = {label_ci.upper(), label_pi.upper(), ""}
    code = df["Code_Analytique"].astype(str)
    mask = (~code.str.upper().isin(labels_exclus)) & (code.str.strip() != "")
    if "Famille_Analytique" in df.columns and df["Famille_Analytique"].astype(str).str.upper().eq("EDITION").any():
        mask = mask & (df["Famille_Analytique"].astype(str).str.upper() == "EDITION")
    return df[mask]

def mask_retours(df_scope, params):
    prefixes_retours = tuple(params.get("retours") or [])
    if not prefixes_retours:
        return pd.Series(False, index=df_scope.index)
    mask = df_scope["Compte"].astype(str).str.startswith(prefixes_retours)
    prefixes_remises = tuple(params.get("remises") or [])
    if prefixes_remises:
        mask = mask & (~df_scope["Compte"].astype(str).str.startswith(prefixes_remises))
    return mask

def mask_remises(df_scope, params):
    prefixes_remises = tuple(params.get("remises") or [])
    if not prefixes_remises:
        return pd.Series(False, index=df_scope.index)
    return df_scope["Compte"].astype(str).str.startswith(prefixes_remises)

def mask_provisions_reprises(df_scope, params):
    prefixes = tuple(params.get("provisions_reprises") or [])
    if not prefixes:
        return pd.Series(False, index=df_scope.index)
    return df_scope["Compte"].astype(str).str.startswith(prefixes)

def mask_ventes(df_scope, params):
    prefixes_ventes = tuple(params.get("ventes") or [])
    mask_conf = (df_scope["Compte"].astype(str).str.startswith(prefixes_ventes)
                 if prefixes_ventes else pd.Series(False, index=df_scope.index))
    mask_rep  = df_scope["Compte"].astype(str) == COMPTE_PRODUITS_INDIRECTS_REPARTIS
    is_7 = df_scope["Compte"].astype(str).str.startswith("7")
    mask_autres = (is_7 & (~mask_retours(df_scope, params))
                   & (~mask_remises(df_scope, params))
                   & (~mask_provisions_reprises(df_scope, params)))
    return mask_conf | mask_rep | mask_autres

def mask_charges(df_scope, params):
    prefixes = tuple(params.get("charges") or [])
    mask = (df_scope["Compte"].astype(str).str.startswith(prefixes)
            if prefixes else pd.Series(False, index=df_scope.index))
    return mask | (df_scope["Compte"].astype(str) == COMPTE_CHARGES_INDIRECTES_REPARTIES)

def normaliser_codes_ean(df, col="Code_Analytique"):
    if col not in df.columns:
        return df
    df = df.copy()
    codes = df[col].astype(str)
    label_ci = st.session_state.get("labels_indirect", {}).get("charges", "CHARGES INDIRECTES")
    label_pi = st.session_state.get("labels_indirect", {}).get("produits", "PRODUITS INDIRECTS")
    labels_reserves = {label_ci.upper(), label_pi.upper(), "", "NAN"}
    mask_ean = (~codes.str.strip().str.upper().isin(labels_reserves)) & codes.str.contains(" - ", regex=False)
    if not mask_ean.any():
        return df
    ean_num = codes[mask_ean].str.split(" - ").str[0].str.strip()
    canonique = codes[mask_ean].groupby(ean_num).agg(lambda s: max(s.unique(), key=len))
    df.loc[mask_ean, col] = ean_num.map(canonique)
    return df

def obtenir_mapping_anonymisation(df):
    titres = sorted(filtrer_isbn_reels(df)["Code_Analytique"].astype(str).unique().tolist())
    cache = st.session_state.get("anonymisation_cache")
    if cache and cache.get("titres_source") == titres:
        return cache["mapping"]
    mapping = {t: f"T{i+1}" for i, t in enumerate(titres)}
    st.session_state["anonymisation_cache"] = {"titres_source": titres, "mapping": mapping}
    return mapping

def label_affiche(code, df_pour_mapping=None):
    if not st.session_state.get("mode_anonyme"):
        return code
    mapping = st.session_state.get("anonymisation_cache", {}).get("mapping")
    if mapping is None and df_pour_mapping is not None:
        mapping = obtenir_mapping_anonymisation(df_pour_mapping)
    return (mapping or {}).get(code, code)

def calculer_indicateurs_titres(df, params, titres):
    df_i = df[df["Code_Analytique"].isin(titres)]
    def par_compte(prefix_list, col, exclude=None):
        if not prefix_list:
            return pd.Series(0.0, index=titres)
        mask = df_i["Compte"].astype(str).str.startswith(tuple(prefix_list))
        if exclude:
            mask = mask & (~df_i["Compte"].astype(str).str.startswith(tuple(exclude)))
        return df_i[mask].groupby("Code_Analytique")[col].sum().reindex(titres, fill_value=0.0)

    ventes  = par_compte(params["ventes"], "Crédit")
    ventes_d = par_compte(params.get("ventes_distributeur") or params["ventes"], "Crédit")
    retours = (par_compte(params["retours"], "Débit", exclude=params.get("remises"))
               - par_compte(params["retours"], "Crédit", exclude=params.get("remises")))
    remises = par_compte(params["remises"], "Débit") - par_compte(params["remises"], "Crédit")
    pstock  = params.get("stock") or ["603"]
    charges = (par_compte(params["charges"], "Débit", exclude=pstock)
               - par_compte(params["charges"], "Crédit", exclude=pstock))
    vstock  = par_compte(pstock, "Débit") - par_compte(pstock, "Crédit")
    prov_rep = params.get("provisions_reprises") or []
    net_prov = (par_compte(prov_rep, "Crédit") - par_compte(prov_rep, "Débit"))
    charges  = charges - net_prov
    mask_cfi = df_i["Compte"].astype(str) == COMPTE_CHARGES_INDIRECTES_REPARTIES
    cf = df_i[mask_cfi].groupby("Code_Analytique")["Débit"].sum().reindex(titres, fill_value=0.0)

    res = pd.DataFrame({"Code_Analytique": titres})
    res["Ventes HT"]  = ventes.values
    res["Retours"]    = retours.values
    res["Remises"]    = remises.values
    res["CA net"]     = res["Ventes HT"] - res["Retours"] - res["Remises"]
    res["Charges variables"] = charges.values
    res["Marge brute"] = res["CA net"] - res["Charges variables"]
    res["Variation de stock"] = vstock.values
    res["Charges fixes imputées"] = cf.values
    res["Résultat net"] = res["Marge brute"] - res["Variation de stock"] - res["Charges fixes imputées"]
    res["Taux retour (%)"] = np.where(ventes_d.values != 0, res["Retours"] / ventes_d.values * 100, 0)
    res["Taux remise (%)"] = np.where(ventes_d.values != 0, res["Remises"] / ventes_d.values * 100, 0)
    res["CA distributeur"] = ventes_d.values

    def _signal(row):
        if row["Résultat net"] > 0 and row["Taux retour (%)"] < 20: return "🟢"
        if row["Résultat net"] > 0 and row["Taux retour (%)"] < 35: return "🟡"
        return "🔴"
    res["Signal"] = res.apply(_signal, axis=1)
    return res

def get_client_ai():
    try:
        return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])
    except Exception:
        return None

SYSTEM_PROMPT_EC = """Tu es un expert-comptable spécialisé dans l'accompagnement des maisons d'édition indépendantes françaises.
Tu connais parfaitement les normes comptables françaises (PCG) appliquées à l'édition, les droits d'auteurs, retours éditeurs, provisions, distributeurs.
Tu assistes l'utilisateur dans l'analyse de ses données comptables analytiques via VISION EDITION.
Réponds toujours en français, de façon concise et orientée action."""

SYSTEM_PROMPT_DIRIGEANT = """Tu es un assistant de pilotage pour les maisons d'édition indépendantes.
Tu aides le dirigeant à comprendre ses données de gestion : chiffre d'affaires, taux de retour, rentabilité par titre, trésorerie.
Réponds de façon simple, claire et sans jargon comptable excessif.
Tu n'as accès qu'aux données de la maison d'édition de l'utilisateur.
Réponds toujours en français."""


# ============================================================
# PAGE : ACCUEIL (EC)
# ============================================================
# ============================================================
def get_df_dirigeant():
    """Récupère le pivot du dossier du dirigeant connecté."""
    did = st.session_state.get("dossier_id")
    if not did:
        st.warning("⚠️ Aucun dossier associé à votre compte.")
        st.stop()
    dos = get_dossier(did)
    if "df_pivot" not in st.session_state:
        st.warning("⚠️ Votre expert-comptable n'a pas encore généré les données de pilotage. "
                   "Contactez votre cabinet.")
        st.stop()
    return st.session_state["df_pivot"].copy(), st.session_state.get("param_comptes",{}), dos



# HELPER check_pivot (défini avant le bloc if/elif)
# ============================================================
def check_pivot():
    if not st.session_state.get("dossier_id"):
        st.warning("⚠️ Sélectionnez un dossier.")
        st.stop()
    if "df_pivot" not in st.session_state:
        st.warning("⚠️ Générez d'abord le socle analytique.")
        st.stop()
    dos = get_dossier(st.session_state["dossier_id"])
    st.info(f"📁 **{dos['nom']}** — {dos.get('exercice','')}")
    return st.session_state["df_pivot"].copy(), st.session_state["param_comptes"]

if role == "ec" and page == "🏠 Accueil":
    st.title("📚 VISION EDITION")
    st.markdown(f"*Bienvenue, **{st.session_state['name']}***")
    st.divider()

    col_main, col_qr = st.columns([3, 1])
    with col_main:
        st.markdown("### Comment démarrer ?")
        c1, c2, c3, c4 = st.columns(4)
        for col, num, icon, titre, desc in [
            (c1,"1","👥","Créer un dossier","Ajoutez votre client maison d'édition"),
            (c2,"2","📂","Importer","Chargez l'export comptable analytique"),
            (c3,"3","⚙️","Paramétrer","Mappez les colonnes et comptes"),
            (c4,"4","📈","Analyser","Explorez les tableaux de bord"),
        ]:
            with col:
                st.markdown(f"""<div style='text-align:center;padding:16px;background:#f8f9fa;
                border-radius:12px;border:1px solid #e0e0e0'>
                <div style='font-size:28px'>{icon}</div>
                <div style='font-size:11px;color:#888;margin:4px 0'>Étape {num}</div>
                <div style='font-weight:600;font-size:14px'>{titre}</div>
                <div style='font-size:12px;color:#666;margin-top:4px'>{desc}</div>
                </div>""", unsafe_allow_html=True)

        st.divider()
        st.markdown("### Outil Analytique Diffuseur")
        st.info("L'outil **Analytique Diffuseur** est une application distincte déployée séparément. "
                "Elle traite le relevé mensuel de votre diffuseur et génère les écritures comptables "
                "prêtes à importer dans votre logiciel avant utilisation de VISION EDITION.")

    with col_qr:
        st.markdown("### Accès mobile")
        qr2 = qrcode.QRCode(version=1, box_size=4, border=2)
        qr2.add_data(APP_URL); qr2.make(fit=True)
        img2 = qr2.make_image(fill_color="black", back_color="white")
        b2 = BytesIO(); img2.save(b2, format="PNG")
        st.image(b2.getvalue(), width=160)
        st.caption("Scanner pour accès mobile")

# ============================================================
# PAGE : GESTION DES DOSSIERS
# ============================================================
elif role == "ec" and page == "👥 Gestion des dossiers":
    st.header("👥 Gestion des dossiers clients")
    cab = st.session_state["cabinet"]

    tab_dossiers, tab_ec_users, tab_import_export = st.tabs([
        "📁 Dossiers", "👤 Utilisateurs EC", "💾 Import / Export config"
    ])

    # ── Onglet Dossiers ──
    with tab_dossiers:
        st.subheader("Créer un nouveau dossier")
        with st.form("form_nouveau_dossier"):
            c1, c2 = st.columns(2)
            nom_dos    = c1.text_input("Nom de la maison d'édition *", placeholder="Editions de l'Argonaute")
            dirigeant  = c1.text_input("Nom du dirigeant *", placeholder="Thomas Bernard")
            exercice   = c2.text_input("Exercice", value="2025/2026")
            periode_d  = c2.text_input("Début de période", value="01/04/2025")
            periode_f  = c2.text_input("Fin de période", value="31/03/2026")
            login_dir  = c1.text_input("Login dirigeant (pour accès client)", placeholder="argonaute")
            pw_dir     = c2.text_input("Mot de passe dirigeant", type="password")
            submitted = st.form_submit_button("➕ Créer ce dossier", type="primary")

            if submitted:
                if not nom_dos or not dirigeant:
                    st.error("Le nom de la maison d'édition et le dirigeant sont obligatoires.")
                else:
                    import uuid
                    did = str(uuid.uuid4())[:8]
                    dossier = {
                        "nom": nom_dos,
                        "dirigeant": dirigeant,
                        "exercice": exercice,
                        "periode_debut": periode_d,
                        "periode_fin": periode_f,
                        "param_comptes": {},
                        "mapping": {},
                        "labels_indirect": {"charges": "CHARGES INDIRECTES", "produits": "PRODUITS INDIRECTS"},
                        "dirigeant_users": {}
                    }
                    if login_dir and pw_dir:
                        dossier["dirigeant_users"][login_dir] = {
                            "pw_hash": _hash(pw_dir),
                            "name": dirigeant
                        }
                    save_dossier(did, dossier)
                    st.success(f"✅ Dossier **{nom_dos}** créé (ID : {did})")
                    if login_dir:
                        st.info(f"Accès dirigeant créé : login = **{login_dir}**")
                    st.rerun()

        st.divider()
        st.subheader("Dossiers existants")
        dossiers = cab["dossiers"]
        if not dossiers:
            st.info("Aucun dossier créé pour l'instant.")
        else:
            for did, dos in dossiers.items():
                with st.expander(f"📁 {dos['nom']} — {dos.get('exercice','')}"):
                    c1, c2, c3 = st.columns(3)
                    c1.write(f"**Dirigeant :** {dos.get('dirigeant','')}")
                    c2.write(f"**Période :** {dos.get('periode_debut','')} au {dos.get('periode_fin','')}")
                    c3.write(f"**ID :** {did}")

                    # Accès dirigeants
                    du = dos.get("dirigeant_users", {})
                    if du:
                        st.write("**Accès dirigeants :**", ", ".join(du.keys()))
                    else:
                        st.caption("Aucun accès dirigeant configuré.")

                    # Ajout accès dirigeant
                    with st.form(f"form_dirigeant_{did}"):
                        dc1, dc2 = st.columns(2)
                        new_login = dc1.text_input("Nouveau login dirigeant", key=f"nl_{did}")
                        new_pw    = dc2.text_input("Mot de passe", type="password", key=f"np_{did}")
                        new_name  = dc1.text_input("Nom affiché", value=dos.get("dirigeant",""), key=f"nn_{did}")
                        if st.form_submit_button("Ajouter accès dirigeant"):
                            if new_login and new_pw:
                                dos["dirigeant_users"][new_login] = {
                                    "pw_hash": _hash(new_pw),
                                    "name": new_name or dos.get("dirigeant","")
                                }
                                save_dossier(did, dos)
                                st.success(f"Accès {new_login} ajouté.")
                                st.rerun()

                    if st.button(f"🗑️ Supprimer ce dossier", key=f"del_{did}"):
                        del st.session_state["cabinet"]["dossiers"][did]
                        if st.session_state["dossier_id"] == did:
                            st.session_state["dossier_id"] = None
                        st.rerun()

    # ── Onglet Utilisateurs EC ──
    with tab_ec_users:
        st.subheader("Collaborateurs du cabinet")
        with st.form("form_ec_user"):
            c1, c2, c3 = st.columns(3)
            ec_login = c1.text_input("Login")
            ec_pw    = c2.text_input("Mot de passe", type="password")
            ec_name  = c3.text_input("Nom affiché")
            if st.form_submit_button("➕ Ajouter collaborateur"):
                if ec_login and ec_pw:
                    cab["ec_users"][ec_login] = {
                        "pw_hash": _hash(ec_pw),
                        "name": ec_name or ec_login,
                        "role": "ec"
                    }
                    st.success(f"Collaborateur {ec_login} ajouté.")
                    st.rerun()

        st.divider()
        for login_u, info in cab["ec_users"].items():
            col1, col2 = st.columns([4,1])
            col1.write(f"👤 **{info['name']}** (login : `{login_u}`)")
            if login_u != st.session_state["username"]:
                if col2.button("Supprimer", key=f"del_ec_{login_u}"):
                    del cab["ec_users"][login_u]
                    st.rerun()

    # ── Onglet Import/Export config ──
    with tab_import_export:
        st.subheader("Sauvegarder la configuration cabinet")
        st.caption("Exportez la configuration complète (dossiers + paramètres) en JSON. "
                   "Réimportez-la à la prochaine session pour retrouver tous vos dossiers.")

        config_export = {
            "ec_users": {
                k: {"name": v["name"], "role": v["role"], "pw_hash": v["pw_hash"]}
                for k, v in cab["ec_users"].items()
            },
            "dossiers": cab["dossiers"]
        }

        st.download_button(
            "💾 Exporter la configuration cabinet (JSON)",
            data=json.dumps(config_export, ensure_ascii=False, indent=2),
            file_name=f"config_cabinet_{_dt.date.today().isoformat()}.json",
            mime="application/json",
            use_container_width=True
        )

        st.divider()
        st.subheader("Recharger une configuration")
        fichier_config = st.file_uploader("Fichier JSON de configuration", type=["json"])
        if fichier_config:
            try:
                cfg = json.load(fichier_config)
                if st.button("📤 Appliquer cette configuration", type="primary"):
                    if "ec_users" in cfg:
                        cab["ec_users"].update(cfg["ec_users"])
                    if "dossiers" in cfg:
                        cab["dossiers"].update(cfg["dossiers"])
                    st.success("✅ Configuration chargée avec succès.")
                    st.rerun()
            except Exception as e:
                st.error(f"Fichier invalide : {e}")


# ============================================================
# PAGE : IMPORT DES DONNÉES (EC)
# ============================================================
elif role == "ec" and page == "📂 Import des données":
    st.header("📂 Import des données analytiques")

    if not st.session_state.get("dossier_id"):
        st.warning("⚠️ Sélectionnez ou créez un dossier client dans **👥 Gestion des dossiers**.")
        st.stop()

    dos = get_dossier(st.session_state["dossier_id"])
    st.info(f"📁 Dossier actif : **{dos['nom']}** — {dos.get('exercice','')}")

    tab1, tab2 = st.tabs(["📁 Importer mon fichier", "🎭 Données de démonstration"])

    with tab1:
        st.info("Importez votre export comptable analytique au format Excel (.xlsx). "
                "Ce fichier est le grand livre analytique exporté depuis votre logiciel comptable "
                "après import des écritures générées par **Analytique Diffuseur**.")
        fichier = st.file_uploader("Sélectionnez votre fichier Excel", type=["xlsx"])
        if fichier:
            try:
                df = pd.read_excel(fichier, header=0)
                df.columns = df.columns.str.strip()
                st.session_state["df_comptables"] = df
                st.success(f"✅ {df.shape[0]} lignes chargées — {df.shape[1]} colonnes")
                st.write("**Colonnes :**", list(df.columns))
                st.dataframe(df.head(10))
                st.info("➡️ Passez dans **⚙️ Paramétrage analytique** pour configurer les colonnes.")
            except Exception as e:
                st.error(f"❌ Erreur : {e}")

    with tab2:
        st.info("Données fictives pour explorer tous les modules.")
        if st.button("🎭 Charger les données de démonstration", type="primary"):
            np.random.seed(42)
            isbns = [f"978-2-{i:04d}-{j:04d}-{k}" for i,j,k in [
                (1234,1001,1),(1234,1002,8),(1234,1003,5),(1234,1004,2),
                (1234,1005,9),(1234,1006,6),(1234,1007,3),(1234,1008,0)]]
            titres_d = ["Le Dernier Manuscrit","Mémoires du Vent","Sous les Toits de Paris",
                        "L'Héritage Silencieux","Chroniques du Nord","La Lumière d'Août",
                        "Terres Inconnues","Les Mots du Soir"]
            rows = []
            dates = pd.date_range("2024-01-01","2024-12-31",freq="MS")
            for date in dates:
                for isbn,titre in zip(isbns,titres_d):
                    v = np.random.randint(500,8000)
                    rows.append({"Compte":"701100","Débit":0,"Crédit":round(v,2),
                                 "Code_Analytique":isbn,"Famille_Analytique":"EDITION",
                                 "Libellé":f"Ventes {titre}","Date":date})
                    if np.random.random()<0.6:
                        rows.append({"Compte":"709100","Débit":round(v*np.random.uniform(0.05,0.30),2),"Crédit":0,
                                     "Code_Analytique":isbn,"Famille_Analytique":"EDITION",
                                     "Libellé":f"Retours {titre}","Date":date})
                    rows.append({"Compte":"607100","Débit":round(v*np.random.uniform(0.15,0.40),2),"Crédit":0,
                                 "Code_Analytique":isbn,"Famille_Analytique":"EDITION",
                                 "Libellé":f"Charges {titre}","Date":date})
                for compte,libelle,montant in [("615000","Loyer",1200),("641000","Salaires",4500)]:
                    rows.append({"Compte":compte,"Débit":montant,"Crédit":0,
                                 "Code_Analytique":"CHARGES INDIRECTES","Famille_Analytique":"EDITION",
                                 "Libellé":libelle,"Date":date})
            df_demo = pd.DataFrame(rows)
            df_demo["Date"] = pd.to_datetime(df_demo["Date"])
            st.session_state["df_comptables"] = df_demo
            st.session_state["df_source_mappe"] = df_demo
            pivot = df_demo.groupby(["Compte","Famille_Analytique","Code_Analytique","Date","Libellé"],
                                    as_index=False).agg({"Débit":"sum","Crédit":"sum"})
            st.session_state["df_pivot"] = pivot
            st.session_state["df_pivot_brut"] = pivot.copy()
            st.session_state["param_comptes"] = {
                "ventes":["701"],"ventes_distributeur":["701"],"retours":["709"],
                "remises":["7091"],"charges":["6"],"provisions_reprises":["781"],
                "charges_imputees":"Oui"
            }
            st.session_state["labels_indirect"] = {"charges":"CHARGES INDIRECTES","produits":"PRODUITS INDIRECTS"}
            st.session_state["familles_cols"] = ["Famille_Analytique"]
            st.session_state["codes_cols"] = ["Code_Analytique"]
            st.session_state["noms_familles_actives"] = ["EDITION"]
            st.session_state["repartition_active"] = False
            st.success("✅ Données de démonstration chargées !")
            st.info("➡️ Accédez à **📈 Tableau de bord éditorial**.")

# ============================================================
# PAGE : PARAMÉTRAGE ANALYTIQUE (EC) — simplifié
# ============================================================
elif role == "ec" and page == "⚙️ Paramétrage analytique":
    st.header("⚙️ Paramétrage analytique")

    if not st.session_state.get("dossier_id"):
        st.warning("⚠️ Sélectionnez un dossier.")
        st.stop()

    dos = get_dossier(st.session_state["dossier_id"])
    st.info(f"📁 **{dos['nom']}** — {dos.get('exercice','')}")

    if "df_comptables" not in st.session_state:
        st.warning("⚠️ Importez d'abord vos données.")
        st.stop()

    df = st.session_state["df_comptables"].copy()
    columns = list(df.columns)

    # Recharger config dossier si disponible
    saved_params = dos.get("param_comptes", {})
    saved_mapping = dos.get("mapping", {})

    with st.expander("📥 Recharger une configuration JSON sauvegardée", expanded=False):
        fconfig = st.file_uploader("Fichier JSON de configuration", type=["json"], key="up_config_param")
        if fconfig:
            try:
                cfg = json.load(fconfig)
                if st.button("Appliquer"):
                    cfg_inv = {v:k for k,v in (cfg.get("mapping",{})).items()}
                    for k_ss, col_cfg in [
                        ("map_compte_col","Compte"),("map_debit_col","Débit"),
                        ("map_credit_col","Crédit"),("map_date_col","Date")]:
                        if cfg_inv.get(col_cfg) in columns:
                            st.session_state[k_ss] = cfg_inv[col_cfg]
                    p = cfg.get("param_comptes",{})
                    st.session_state["map_ventes_comptes"] = ",".join(p.get("ventes",["701"]))
                    st.session_state["map_retours_comptes"] = ",".join(p.get("retours",["709"]))
                    st.session_state["map_remises_comptes"] = ",".join(p.get("remises",["7091"]))
                    st.session_state["map_charges_comptes"] = ",".join(p.get("charges",["6"]))
                    st.session_state["map_provisions_reprises_comptes"] = ",".join(p.get("provisions_reprises",["781"]))
                    st.session_state["map_ventes_distributeur_comptes"] = ",".join(p.get("ventes_distributeur",["7011"]))
                    st.success("✅ Configuration rechargée.")
                    st.rerun()
            except Exception as e:
                st.error(f"Fichier invalide : {e}")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Mapping des colonnes")
        compte_col  = st.selectbox("Colonne Compte", columns, key="map_compte_col")
        debit_col   = st.selectbox("Colonne Débit", columns, key="map_debit_col")
        credit_col  = st.selectbox("Colonne Crédit", columns, key="map_credit_col")
        date_col    = st.selectbox("Colonne Date", columns, key="map_date_col")
        libelle_col = st.selectbox("Libellé (optionnel)", [""]+columns, key="map_libelle_col")
        journal_col = st.selectbox("Code journal (optionnel)", [""]+columns, key="map_journal_col")
    with col2:
        st.subheader("Comptes comptables")
        ventes_comptes  = st.text_input("Comptes ventes (CA large)", value=",".join(saved_params.get("ventes",["701"])), key="map_ventes_comptes")
        ventes_dist     = st.text_input("Comptes ventes distributeur", value=",".join(saved_params.get("ventes_distributeur",["7011"])), key="map_ventes_distributeur_comptes")
        retours_comptes = st.text_input("Comptes retours", value=",".join(saved_params.get("retours",["709"])), key="map_retours_comptes")
        remises_comptes = st.text_input("Comptes remises", value=",".join(saved_params.get("remises",["7091"])), key="map_remises_comptes")
        charges_comptes = st.text_input("Comptes charges", value=",".join(saved_params.get("charges",["6"])), key="map_charges_comptes")
        prov_comptes    = st.text_input("Comptes reprises provisions", value=",".join(saved_params.get("provisions_reprises",["781"])), key="map_provisions_reprises_comptes")

    st.subheader("Familles analytiques")
    nb_familles = st.number_input("Nombre de familles", min_value=1, max_value=4, value=1, step=1, key="map_nb_familles")
    familles_mapping = []
    noms_suggestion = ["EDITION","COMMUNICATION","Types de dépenses / revenus","AUTEUR"]
    for i in range(int(nb_familles)):
        fc1,fc2,fc3 = st.columns(3)
        nom_f = fc1.text_input(f"Famille {i+1}", value=noms_suggestion[i] if i<4 else "", key=f"nom_famille_{i}")
        fam_col = fc2.selectbox(f"Colonne famille {i+1}", [""]+columns, key=f"famille_col_{i}")
        cod_col = fc3.selectbox(f"Colonne catégorie {i+1}", [""]+columns, key=f"code_col_{i}")
        familles_mapping.append({"nom":nom_f,"famille_col":fam_col,"code_col":cod_col})

    col_li1,col_li2 = st.columns(2)
    label_ci = col_li1.text_input("Libellé charges indirectes", value="CHARGES INDIRECTES", key="map_label_ci")
    label_pi = col_li2.text_input("Libellé produits indirects", value="PRODUITS INDIRECTS", key="map_label_pi")

    if st.button("⚙️ Générer le socle analytique", type="primary"):
        mapping = {compte_col:"Compte",debit_col:"Débit",credit_col:"Crédit",date_col:"Date"}
        if libelle_col: mapping[libelle_col] = "Libellé"
        if journal_col: mapping[journal_col] = "Journal"
        familles_cols, codes_cols, noms_familles_actives = [], [], []
        for i,fam in enumerate(familles_mapping):
            suffix = "" if i==0 else f"_{i+1}"
            col_f_out = f"Famille_Analytique{suffix}"
            col_c_out = f"Code_Analytique{suffix}"
            if fam["famille_col"]: mapping[fam["famille_col"]] = col_f_out
            if fam["code_col"]:    mapping[fam["code_col"]]    = col_c_out
            familles_cols.append(col_f_out); codes_cols.append(col_c_out)
            noms_familles_actives.append(fam["nom"] or f"Famille {i+1}")

        df.rename(columns=mapping, inplace=True)
        for col in familles_cols+codes_cols+["Libellé","Journal"]:
            if col not in df.columns: df[col] = ""
            else: df[col] = df[col].fillna("")
        df["Date"]   = pd.to_datetime(df["Date"], errors="coerce")
        df["Débit"]  = pd.to_numeric(df["Débit"], errors="coerce").fillna(0)
        df["Crédit"] = pd.to_numeric(df["Crédit"], errors="coerce").fillna(0)
        df["Compte"] = df["Compte"].astype(str).str.strip()

        def sc(s): return [c.strip() for c in s.split(",") if c.strip()]
        params = {
            "ventes":  sc(ventes_comptes),
            "ventes_distributeur": sc(ventes_dist) or sc(ventes_comptes),
            "retours": sc(retours_comptes),
            "remises": sc(remises_comptes),
            "charges": sc(charges_comptes),
            "provisions_reprises": sc(prov_comptes),
            "charges_imputees": "Oui",
        }

        group_cols = ["Compte"]+familles_cols+codes_cols+["Date"]
        if "Libellé" in df.columns: group_cols.append("Libellé")
        if "Journal" in df.columns: group_cols.append("Journal")
        pivot = df.groupby(group_cols, as_index=False).agg({"Débit":"sum","Crédit":"sum"})
        pivot = normaliser_codes_ean(pivot)

        st.session_state["df_source_mappe"] = df
        st.session_state["df_pivot"] = pivot
        st.session_state["df_pivot_brut"] = pivot.copy()
        st.session_state["param_comptes"] = params
        st.session_state["labels_indirect"] = {"charges":label_ci.strip(),"produits":label_pi.strip()}
        st.session_state["familles_cols"] = familles_cols
        st.session_state["codes_cols"] = codes_cols
        st.session_state["noms_familles_actives"] = noms_familles_actives
        st.session_state["repartition_active"] = False

        # Sauvegarder dans le dossier
        dos["param_comptes"] = params
        dos["mapping"] = mapping
        dos["labels_indirect"] = st.session_state["labels_indirect"]
        save_dossier(st.session_state["dossier_id"], dos)

        st.success("✅ Socle analytique généré et paramètres sauvegardés dans le dossier.")

        config_out = {"mapping":mapping,"param_comptes":params,
                      "familles_cols":familles_cols,"codes_cols":codes_cols,
                      "noms_familles_actives":noms_familles_actives,
                      "labels_indirect":st.session_state["labels_indirect"]}
        st.download_button("💾 Sauvegarder la configuration (JSON)",
            data=json.dumps(config_out,ensure_ascii=False,indent=2),
            file_name=f"config_{dos['nom'][:20].replace(' ','_')}.json",
            mime="application/json")
        st.dataframe(pivot.head(15))



# ── TABLEAU DE BORD ──
elif role == "ec" and page == "📈 Tableau de bord éditorial":
    df, params = check_pivot()
    dos = get_dossier(st.session_state["dossier_id"])
    st.header(f"📈 Tableau de bord — {dos['nom']}")
    if st.session_state.get("repartition_active"):
        st.caption("ℹ️ Charges indirectes réparties sur les titres actifs.")

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    years = sorted(df["Date"].dt.year.dropna().unique().tolist())
    annee = st.selectbox("Année", ["Toutes"]+[str(y) for y in years])
    if annee != "Toutes":
        df = df[df["Date"].dt.year == int(annee)]

    df_date_ok = df.dropna(subset=["Date"]).copy()
    df_date_ok["Mois"] = df_date_ok["Date"].dt.to_period("M").astype(str)

    df_v   = df[mask_ventes(df, params)]
    df_r   = df[mask_retours(df, params)]
    df_rem = df[mask_remises(df, params)]
    df_c   = df[mask_charges(df, params)]
    pref_dist = tuple(params.get("ventes_distributeur") or params["ventes"])
    df_dist = df[df["Compte"].astype(str).str.startswith(pref_dist)]
    ca_dist = df_dist["Crédit"].sum()

    df_prov = df[mask_provisions_reprises(df, params)]
    net_prov = df_prov["Crédit"].sum() - df_prov["Débit"].sum()

    ca_brut       = df_v["Crédit"].sum()
    corrections   = df_v["Débit"].sum()
    total_retours = df_r["Débit"].sum() - df_r["Crédit"].sum()
    total_remises = df_rem["Débit"].sum() - df_rem["Crédit"].sum()
    ca_net        = ca_brut - total_retours - total_remises - corrections
    charges_tot   = (df_c["Débit"].sum() - df_c["Crédit"].sum()) - net_prov
    resultat      = ca_net - charges_tot
    taux_ret      = (total_retours / ca_dist * 100) if ca_dist else 0
    taux_rem      = (total_remises / ca_dist * 100) if ca_dist else 0

    k1,k2,k3,k4,k5 = st.columns(5)
    k1.metric("CA brut", f"{fmt_fr(ca_brut)} €")
    k2.metric("CA net", f"{fmt_fr(ca_net)} €", delta=f"-{fmt_fr(total_retours+total_remises+corrections)} €")
    k3.metric("Taux de retour", f"{taux_ret:.1f} %",
              delta_color="inverse", delta="⚠️ Élevé" if taux_ret > 25 else "✅ Normal")
    k4.metric("Charges totales", f"{fmt_fr(charges_tot)} €")
    k5.metric("Résultat net", f"{fmt_fr(resultat)} €")

    col_g1, col_g2 = st.columns(2)
    with col_g1:
        st.subheader("Évolution mensuelle")
        df_v_d = df_date_ok[mask_ventes(df_date_ok, params)]
        df_r_d = df_date_ok[mask_retours(df_date_ok, params)]
        tv = df_v_d.groupby("Mois")["Crédit"].sum().reset_index().rename(columns={"Crédit":"CA brut"})
        tr = df_r_d.groupby("Mois")["Débit"].sum().reset_index().rename(columns={"Débit":"Retours"})
        trend = tv.merge(tr, on="Mois", how="left").fillna(0)
        trend["CA net"] = trend["CA brut"] - trend["Retours"]
        fig = go.Figure()
        fig.add_trace(go.Bar(x=trend["Mois"], y=trend["CA brut"], name="CA brut", marker_color="#3B82F6"))
        fig.add_trace(go.Bar(x=trend["Mois"], y=trend["Retours"], name="Retours", marker_color="#EF4444"))
        fig.add_trace(go.Scatter(x=trend["Mois"], y=trend["CA net"], name="CA net",
                                 mode="lines+markers", line=dict(color="#10B981",width=2)))
        fig.update_layout(barmode="overlay", height=320, margin=dict(t=20), legend=dict(orientation="h"))
        st.plotly_chart(fig, use_container_width=True)

    with col_g2:
        st.subheader("Top 10 titres")
        df_isbn = filtrer_isbn_reels(df)
        top = df_isbn.groupby("Code_Analytique", as_index=False).agg({"Crédit":"sum","Débit":"sum"})
        top["Résultat"] = top["Crédit"] - top["Débit"]
        top10 = top.nlargest(10,"Résultat").copy()
        top10["Titre"] = top10["Code_Analytique"].apply(lambda c: label_affiche(c, df))
        fig2 = px.bar(top10, x="Titre", y="Résultat", color="Résultat",
                      color_continuous_scale=["#EF4444","#F59E0B","#10B981"], height=320)
        fig2.update_layout(showlegend=False, margin=dict(t=20))
        st.plotly_chart(fig2, use_container_width=True)

# ── ANALYSE PAR TITRE ──
elif role == "ec" and page == "📖 Analyse par titre":
    df, params = check_pivot()
    dos = get_dossier(st.session_state["dossier_id"])
    st.header(f"📖 Analyse par titre — {dos['nom']}")

    titres = sorted(filtrer_isbn_reels(df)["Code_Analytique"].astype(str).unique().tolist())
    if not titres:
        st.warning("Aucun ISBN détecté.")
        st.stop()

    indic = calculer_indicateurs_titres(df, params, titres)
    col_s, col_t, col_f = st.columns(3)

    def _liste(container, sous_titre, df_tri, col_montant):
        container.markdown(f"**{sous_titre}**")
        for _, row in df_tri.iterrows():
            container.write(f"{row['Signal']} **{label_affiche(row['Code_Analytique'],df)}** — {fmt_fr(row[col_montant])} €")

    _liste(col_s, "📊 Plus significatifs", indic.sort_values("Ventes HT",ascending=False).head(5), "Ventes HT")
    _liste(col_t, "🏆 Plus rentables", indic.sort_values("Résultat net",ascending=False).head(5), "Résultat net")
    _liste(col_f, "⚠️ Plus difficiles", indic.sort_values("Résultat net",ascending=True).head(5), "Résultat net")

    st.divider()
    isbn_sel = st.selectbox("Sélectionner un titre", titres, format_func=lambda c: label_affiche(c,df))
    df_t = df[df["Code_Analytique"] == isbn_sel]
    df_v_ = df_t[df_t["Compte"].astype(str).str.startswith(tuple(params["ventes"]))]
    df_r_ = df_t[mask_retours(df_t, params)]
    df_rem_ = df_t[mask_remises(df_t, params)]
    pstock = tuple(params.get("stock") or ["603"])
    df_c_ = df_t[df_t["Compte"].astype(str).str.startswith(tuple(params["charges"]))
                  & (~df_t["Compte"].astype(str).str.startswith(pstock))]
    df_cfi = df_t[df_t["Compte"].astype(str) == COMPTE_CHARGES_INDIRECTES_REPARTIES]

    ventes_ht = df_v_["Crédit"].sum()
    retours_m = df_r_["Débit"].sum() - df_r_["Crédit"].sum()
    remises_m = df_rem_["Débit"].sum() - df_rem_["Crédit"].sum()
    ca_net_t  = ventes_ht - retours_m - remises_m
    charges_v = df_c_["Débit"].sum() - df_c_["Crédit"].sum()
    marge_b   = ca_net_t - charges_v
    cf        = df_cfi["Débit"].sum()
    res_net   = marge_b - cf

    k1,k2,k3,k4 = st.columns(4)
    k1.metric("CA net", f"{fmt_fr(ca_net_t)} €")
    k2.metric("Marge brute", f"{fmt_fr(marge_b)} €")
    k3.metric("Charges fixes imputées", f"{fmt_fr(cf)} €")
    k4.metric("Résultat net", f"{fmt_fr(res_net)} €")

    rows_sig = [
        ("Ventes HT", ventes_ht, "base"),
        ("− Retours", -retours_m, "deduction"),
        ("− Remises", -remises_m, "deduction"),
        ("= CA net", ca_net_t, "subtotal"),
        ("− Charges variables", -charges_v, "deduction"),
        ("= Marge brute", marge_b, "subtotal"),
        ("− Charges fixes imputées", -cf, "deduction"),
        ("= Résultat net", res_net, "total"),
    ]
    measures = [{"base":"absolute","deduction":"relative","subtotal":"total","total":"total"}[r[2]] for r in rows_sig]
    fig3 = go.Figure(go.Waterfall(
        orientation="v", measure=measures,
        x=[r[0] for r in rows_sig], y=[r[1] for r in rows_sig],
        decreasing={"marker":{"color":"#EF4444"}},
        increasing={"marker":{"color":"#10B981"}},
        totals={"marker":{"color":"#3B82F6"}}
    ))
    fig3.update_layout(height=300, margin=dict(t=10))
    st.plotly_chart(fig3, use_container_width=True)

# ── SIMULATEUR ──
elif role == "ec" and page == "🎯 Simulateur de rentabilité":
    df, params = check_pivot()
    dos = get_dossier(st.session_state["dossier_id"])
    st.header(f"🎯 Simulateur de rentabilité — {dos['nom']}")
    titres = sorted(filtrer_isbn_reels(df)["Code_Analytique"].astype(str).unique().tolist())
    if not titres:
        st.warning("Aucun ISBN détecté.")
        st.stop()
    isbn_sim = st.selectbox("Titre à simuler", titres, format_func=lambda c: label_affiche(c,df))
    df_t = df[df["Code_Analytique"]==isbn_sim]
    df_v_ = df_t[df_t["Compte"].astype(str).str.startswith(tuple(params["ventes"]))]
    df_vd = df_t[df_t["Compte"].astype(str).str.startswith(tuple(params.get("ventes_distributeur") or params["ventes"]))]
    df_r_ = df_t[mask_retours(df_t,params)]
    df_rem_ = df_t[mask_remises(df_t,params)]
    pstock = tuple(params.get("stock") or ["603"])
    df_c_ = df_t[df_t["Compte"].astype(str).str.startswith(tuple(params["charges"]))
                  & (~df_t["Compte"].astype(str).str.startswith(pstock))]
    df_cfi = df_t[df_t["Compte"].astype(str)==COMPTE_CHARGES_INDIRECTES_REPARTIES]

    vb  = df_v_["Crédit"].sum(); vd = df_vd["Crédit"].sum()
    rb  = df_r_["Débit"].sum()-df_r_["Crédit"].sum()
    rem = df_rem_["Débit"].sum()-df_rem_["Crédit"].sum()
    cv  = df_c_["Débit"].sum()-df_c_["Crédit"].sum()
    cf  = df_cfi["Débit"].sum()
    tr  = (rb/vd*100) if vd else 0

    col1,col2 = st.columns(2)
    with col1:
        tr_hyp = st.slider("Taux de retour hypothétique (%)", 0.0, 100.0, float(round(tr,1)), 0.5)
        var_ch = st.slider("Variation charges variables (%)", -50, 50, 0, 1)
    with col2:
        nb_t = (st.session_state.get("repartition_detail") or {}).get("nb_titres_actifs", len(titres))
        nb_t_h = st.slider("Nombre titres actifs (clé répartition)", 1, max(nb_t*2,nb_t+5), int(nb_t))
        total_ci = (st.session_state.get("repartition_detail") or {}).get("part_charge",0)*nb_t

    ret_h = vd*tr_hyp/100; cv_h = cv*(1+var_ch/100)
    ca_h  = vb-ret_h-rem; mb_h = ca_h-cv_h
    cf_h  = (total_ci/nb_t_h) if (nb_t_h and st.session_state.get("repartition_active")) else cf
    res_h = mb_h-cf_h

    st.divider()
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("CA net (simulé)",f"{fmt_fr(ca_h)} €",delta=f"{fmt_fr(ca_h-(vb-rb-rem))} €")
    c2.metric("Marge brute (simulée)",f"{fmt_fr(mb_h)} €",delta=f"{fmt_fr(mb_h-(vb-rb-rem-cv))} €")
    c3.metric("Résultat net (simulé)",f"{fmt_fr(res_h)} €",delta=f"{fmt_fr(res_h-(vb-rb-rem-cv-cf))} €")
    c4.metric("Charges fixes (simulées)",f"{fmt_fr(cf_h)} €")

    if vd:
        seuil = (vb-rem-cv_h-cf_h)/vd*100
        seuil = max(0,seuil)
        marge_s = seuil-tr_hyp
        if marge_s>=0:
            st.success(f"✅ Seuil de rentabilité : taux de retour max **{seuil:.1f}%** (marge de {marge_s:.1f} pts)")
        else:
            st.error(f"❌ Déficitaire — seuil d'équilibre à **{seuil:.1f}%** (vs {tr_hyp:.1f}% simulé)")


# ── TRÉSORERIE (EC) ──
elif role == "ec" and page == "💰 Trésorerie prévisionnelle":
    df, params = check_pivot()
    dos = get_dossier(st.session_state["dossier_id"])
    st.header(f"💰 Trésorerie prévisionnelle — {dos['nom']}")
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Débit"] = pd.to_numeric(df["Débit"], errors="coerce").fillna(0)
    df["Crédit"] = pd.to_numeric(df["Crédit"], errors="coerce").fillna(0)
    if "Journal" not in df.columns: df["Journal"] = ""
    df = df.dropna(subset=["Date"])
    if df.empty:
        st.warning("Aucune écriture datée.")
        st.stop()

    col1, col2 = st.columns(2)
    with col1:
        date_debut = st.date_input("Date de départ", df["Date"].min())
        tresorerie_ouv = st.number_input("Trésorerie à l'ouverture (€)", value=0.0, step=100.0)
    with col2:
        horizon = st.slider("Horizon projection (mois)", 0, 24, 6)
        comptes_banque = tuple(x.strip() for x in st.text_input("Comptes trésorerie","512,530,580").split(",") if x.strip())
        journaux_exclus = tuple(x.strip() for x in st.text_input("Journaux exclus (AN)","AN").split(",") if x.strip())

    mask_an = df["Journal"].isin(journaux_exclus) if journaux_exclus else pd.Series(False, index=df.index)
    df_flux = df[~mask_an & (df["Date"]>=pd.to_datetime(date_debut))].copy()
    df_flux["Mois"] = df_flux["Date"].dt.to_period("M")

    if df_flux.empty:
        st.warning("Aucune écriture après la date de départ.")
        st.stop()

    mois = sorted(df_flux["Mois"].unique())
    encaissements = df_flux[df_flux["Compte"].astype(str).str.startswith("41")].groupby("Mois")["Crédit"].sum()
    decaissements = df_flux[df_flux["Compte"].astype(str).str.startswith(("40","42","43","44"))].groupby("Mois")["Débit"].sum()
    flux_net = encaissements.sub(decaissements, fill_value=0).reindex(mois, fill_value=0)
    treso = tresorerie_ouv + flux_net.cumsum()

    st.session_state["treso_real"] = treso
    st.session_state["treso_ouverture"] = tresorerie_ouv

    m1,m2,m3 = st.columns(3)
    m1.metric("Trésorerie ouverture", f"{fmt_fr(tresorerie_ouv)} €")
    m2.metric("Trésorerie clôture (réalisé)", f"{fmt_fr(treso.iloc[-1])} €")
    m3.metric("Flux net généré", f"{fmt_fr(flux_net.sum())} €")

    _MOIS_FR = {1:"Jan",2:"Fév",3:"Mar",4:"Avr",5:"Mai",6:"Jui",7:"Jul",8:"Aoû",9:"Sep",10:"Oct",11:"Nov",12:"Déc"}
    def ml(p): return f"{_MOIS_FR.get(p.month,str(p.month))} {p.year}"

    fig_t = go.Figure()
    fig_t.add_trace(go.Scatter(x=[ml(m) for m in treso.index], y=treso.values,
                               name="Réalisé", line=dict(color="#111827",width=2.5)))

    if horizon > 0:
        base_enc = encaissements.iloc[-3:].mean() if len(encaissements)>=3 else (encaissements.mean() if len(encaissements) else 0)
        base_dec = decaissements.iloc[-3:].mean() if len(decaissements)>=3 else (decaissements.mean() if len(decaissements) else 0)
        col_sc1,col_sc2,col_sc3,col_sc4 = st.columns(4)
        t_opt  = col_sc1.number_input("Croissance encaissements optimiste (%/mois)", value=4.0, step=0.5)/100
        t_cent = col_sc2.number_input("Croissance encaissements central (%/mois)", value=2.0, step=0.5)/100
        t_pess = col_sc3.number_input("Croissance encaissements pessimiste (%/mois)", value=0.0, step=0.5)/100
        t_ch   = col_sc4.number_input("Évolution charges (%/mois)", value=1.0, step=0.5)/100
        for nom, tc, col_hex in [("Optimiste",t_opt,"#10B981"),("Central",t_cent,"#3B82F6"),("Pessimiste",t_pess,"#EF4444")]:
            futurs = [mois[-1]+i for i in range(1,horizon+1)]
            enc_f = base_enc; dec_f = base_dec; proj = []
            for m_f in futurs:
                enc_f *= (1+tc); dec_f *= (1+t_ch); proj.append(enc_f-dec_f)
            treso_proj = treso.iloc[-1] + pd.Series(proj, index=futurs).cumsum()
            treso_full = pd.concat([treso.iloc[[-1]], treso_proj])
            fig_t.add_trace(go.Scatter(x=[ml(m) for m in treso_full.index], y=treso_full.values,
                                       name=nom, line=dict(color=col_hex,width=2,dash="dot")))

    fig_t.add_hline(y=0, line_dash="dash", line_color="gray")
    fig_t.update_layout(height=380, margin=dict(t=20), xaxis_title="", yaxis_title="€",
                        legend=dict(orientation="h"))
    st.plotly_chart(fig_t, use_container_width=True)

# ── DROITS D'AUTEURS (EC) ──
elif role == "ec" and page == "✍️ Droits d'auteurs":
    df, params = check_pivot()
    dos = get_dossier(st.session_state["dossier_id"])
    st.header(f"✍️ Droits d'auteurs — {dos['nom']}")
    st.info("Module complet disponible — configurez les comptes ci-dessous.")
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    col1,col2 = st.columns(2)
    with col1:
        date_debut_d = st.date_input("Période du", df["Date"].dropna().min().date() if not df["Date"].dropna().empty else _dt.date.today(), key="da_d")
        compte_db = st.text_input("Droits bruts (charge)", value="604300000")
        compte_urs = st.text_input("URSSAF à payer", value="438106")
    with col2:
        date_fin_d = st.date_input("Au", df["Date"].dropna().max().date() if not df["Date"].dropna().empty else _dt.date.today(), key="da_f")
        compte_diff = st.text_input("Contribution diffuseur", value="645106")
        compte_net  = st.text_input("Droits à payer (net)", value="408106")

    mask_p = (df["Date"]>=pd.to_datetime(date_debut_d)) & (df["Date"]<=pd.to_datetime(date_fin_d))
    df_p = df[mask_p]

    def par_isbn(cpt, sens):
        m = df_p["Compte"].astype(str).str.strip()==str(cpt).strip()
        if not m.any(): return pd.Series(dtype=float)
        g = df_p[m].groupby("Code_Analytique")
        return (g["Débit"].sum()-g["Crédit"].sum()) if sens=="debit" else (g["Crédit"].sum()-g["Débit"].sum())

    db_s = par_isbn(compte_db,"debit"); urs_s = par_isbn(compte_urs,"credit")
    diff_s = par_isbn(compte_diff,"debit"); net_s = par_isbn(compte_net,"credit")
    av_s = par_isbn("409600","debit")

    isbns_d = sorted(set(db_s.index)|set(urs_s.index)|set(diff_s.index)|set(net_s.index))
    isbns_d = [i for i in isbns_d if str(i).strip() not in ("","CHARGES INDIRECTES","PRODUITS INDIRECTS")]

    if not isbns_d:
        st.info("Aucune écriture trouvée sur ces comptes pour la période.")
    else:
        lignes = []
        for isbn in isbns_d:
            lignes.append({
                "ISBN": isbn,
                "Droits bruts (€)": round(float(db_s.get(isbn,0)),2),
                "Contribution diffuseur (€)": round(float(diff_s.get(isbn,0)),2),
                "Précompte URSSAF (€)": round(float(urs_s.get(isbn,0)),2),
                "Net dû auteur (€)": round(float(net_s.get(isbn,0)),2),
                "À-valoir restant (€)": round(float(av_s.get(isbn,0)),2),
            })
        df_d = pd.DataFrame(lignes)
        cols_m = ["Droits bruts (€)","Contribution diffuseur (€)","Précompte URSSAF (€)","Net dû auteur (€)","À-valoir restant (€)"]
        st.dataframe(df_d.style.format({c:(lambda x: f"{fmt_fr(x,2)} €") for c in cols_m}), use_container_width=True)

        total_db = df_d["Droits bruts (€)"].sum()
        total_urs = df_d["Précompte URSSAF (€)"].sum()
        total_diff = df_d["Contribution diffuseur (€)"].sum()
        total_net = df_d["Net dû auteur (€)"].sum()

        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Droits bruts", f"{fmt_fr(total_db,2)} €")
        c2.metric("Précompte URSSAF", f"{fmt_fr(total_urs,2)} €")
        c3.metric("Contribution diffuseur", f"{fmt_fr(total_diff,2)} €")
        c4.metric("Net dû aux auteurs", f"{fmt_fr(total_net,2)} €")

        st.markdown(f"**Total à reverser URSSAF : {fmt_fr(total_urs+total_diff,2)} €**")

        buf_d = BytesIO()
        with pd.ExcelWriter(buf_d, engine="openpyxl") as writer:
            df_d.to_excel(writer, index=False, sheet_name="Droits_auteurs")
        buf_d.seek(0)
        st.download_button("📥 Exporter (Excel)", buf_d,
            file_name=f"Droits_auteurs_{date_debut_d}_{date_fin_d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ── RETOURS & REMISES (EC) ──
elif role == "ec" and page == "📦 Retours & Remises":
    df, params = check_pivot()
    dos = get_dossier(st.session_state["dossier_id"])
    st.header(f"📦 Retours & Remises — {dos['nom']}")
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Mois"] = df["Date"].dt.strftime("%Y-%m")
    seuil = st.sidebar.number_input("Seuil alerte (%)", value=25, step=5)

    df_v_  = df[df["Compte"].astype(str).str.startswith(tuple(params["ventes"]))]
    df_r_  = df[mask_retours(df,params)]
    df_rem_= df[mask_remises(df,params)]
    pref_d = tuple(params.get("ventes_distributeur") or params["ventes"])
    df_vd_ = df[df["Compte"].astype(str).str.startswith(pref_d)]

    tv = df_v_["Crédit"].sum(); vd = df_vd_["Crédit"].sum()
    tr = abs(df_r_["Débit"].sum()-df_r_["Crédit"].sum()) if not df_r_.empty else 0
    rem= abs(df_rem_["Débit"].sum()-df_rem_["Crédit"].sum()) if not df_rem_.empty else 0
    taux_r = (tr/vd*100) if vd else 0; taux_rem_pct = (rem/vd*100) if vd else 0

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("CA brut",f"{fmt_fr(tv)} €")
    c2.metric("Total retours",f"{fmt_fr(tr)} €")
    c3.metric("Taux de retour",f"{taux_r:.1f} %",
              delta="⚠️ Dépasse le seuil" if taux_r>seuil else "✅ Normal",
              delta_color="inverse" if taux_r>seuil else "normal")
    c4.metric("Taux de remise",f"{taux_rem_pct:.1f} %")

    if taux_r > seuil:
        st.error(f"🚨 Taux de retour ({taux_r:.1f}%) dépasse votre seuil de {seuil}% !")

    if not df_r_.empty:
        trend_r = df_r_.groupby("Mois")["Débit"].sum().abs().reset_index()
        fig_r = px.bar(trend_r, x="Mois", y="Débit", title="Retours mensuels (€)",
                       color_discrete_sequence=["#EF4444"], height=300)
        st.plotly_chart(fig_r, use_container_width=True)

        df_ri = filtrer_isbn_reels(df_r_)
        if not df_ri.empty:
            ret_isbn = df_ri.groupby("Code_Analytique")["Débit"].sum().abs().reset_index().sort_values("Débit",ascending=False)
            st.subheader("Retours par titre")
            st.dataframe(ret_isbn, hide_index=True)

# ── SYNTHÈSE FINANCIÈRE (EC) ──
elif role == "ec" and page == "📊 Synthèse financière":
    df, params = check_pivot()
    dos = get_dossier(st.session_state["dossier_id"])
    st.header(f"📊 Synthèse financière — {dos['nom']}")

    df_v_   = df[mask_ventes(df,params)]
    df_r_   = df[mask_retours(df,params)]
    df_rem_ = df[mask_remises(df,params)]
    df_c_   = df[mask_charges(df,params)]
    df_prov = df[mask_provisions_reprises(df,params)]
    net_p   = df_prov["Crédit"].sum()-df_prov["Débit"].sum()

    ca_b  = df_v_["Crédit"].sum(); corr = df_v_["Débit"].sum()
    tr_   = df_r_["Débit"].sum()-df_r_["Crédit"].sum()
    rem_  = df_rem_["Débit"].sum()-df_rem_["Crédit"].sum()
    ca_n  = ca_b-tr_-rem_-corr
    ch_t  = (df_c_["Débit"].sum()-df_c_["Crédit"].sum())-net_p
    res_  = ca_n-ch_t
    mb_p  = (res_/ca_b*100) if ca_b else 0

    soldes   = [ca_b,-tr_,-rem_,-corr,ca_n,-ch_t,res_]
    libelles = ["CA brut","Retours","Remises","Corrections","CA net","Charges","Résultat net"]

    col1,col2 = st.columns([1.2,1])
    with col1:
        st.subheader("Compte de résultat synthétique")
        df_sum = pd.DataFrame({"Poste":libelles,"Montant (€)":soldes})
        st.dataframe(df_sum.style.format({"Montant (€)":(lambda x: fmt_fr(x))}), hide_index=True)
        st.metric("Taux de marge nette",f"{mb_p:.1f} %")
    with col2:
        fig_w = go.Figure(go.Waterfall(
            orientation="v",
            measure=["absolute","relative","relative","relative","total","relative","total"],
            x=libelles, y=soldes,
            decreasing={"marker":{"color":"#EF4444"}},
            increasing={"marker":{"color":"#10B981"}},
            totals={"marker":{"color":"#3B82F6"}}
        ))
        fig_w.update_layout(height=360, margin=dict(t=20))
        st.plotly_chart(fig_w, use_container_width=True)

    buf_s = BytesIO()
    with pd.ExcelWriter(buf_s, engine="openpyxl") as writer:
        df_sum.to_excel(writer, index=False, sheet_name="Synthese")
    buf_s.seek(0)
    st.download_button("📥 Exporter la synthèse (Excel)", buf_s,
        file_name="Synthese_Financiere.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ── ASSISTANT IA (EC) ──
elif role == "ec" and page == "🤖 Assistant IA":
    df_ia, params_ia = check_pivot()
    dos = get_dossier(st.session_state["dossier_id"])
    st.header(f"🤖 Assistant IA — {dos['nom']}")
    client_ai = get_client_ai()
    if client_ai is None:
        st.error("⚠️ Clé API Anthropic non configurée (ANTHROPIC_API_KEY dans les secrets Streamlit).")
        st.stop()

    # Contexte données
    df_v_ = df_ia[mask_ventes(df_ia,params_ia)]
    df_r_ = df_ia[mask_retours(df_ia,params_ia)]
    ca_b_ = df_v_["Crédit"].sum(); tr_ = df_r_["Débit"].sum()-df_r_["Crédit"].sum()
    ca_n_ = ca_b_ - tr_
    taux_r_ = (tr_/ca_b_*100) if ca_b_ else 0
    nb_isbn_ = filtrer_isbn_reels(df_ia)["Code_Analytique"].nunique()

    data_ctx = f"""
DONNÉES ANALYTIQUES — {dos['nom']} — {dos.get('exercice','')}
CA brut : {fmt_fr(ca_b_)} EUR
Total retours : {fmt_fr(tr_)} EUR
CA net : {fmt_fr(ca_n_)} EUR
Taux de retour : {taux_r_:.1f} %
Nombre de titres actifs : {nb_isbn_}
"""
    sys_prompt = SYSTEM_PROMPT_EC + f"\n\nDONNÉES DU DOSSIER :\n{data_ctx}"

    for msg in st.session_state["messages_agent"]:
        with st.chat_message(msg["role"], avatar="🧑" if msg["role"]=="user" else "🤖"):
            st.markdown(msg["content"])

    if prompt_u := st.chat_input("Posez votre question..."):
        st.session_state["messages_agent"].append({"role":"user","content":prompt_u})
        with st.chat_message("user", avatar="🧑"):
            st.markdown(prompt_u)
        with st.chat_message("assistant", avatar="🤖"):
            with st.spinner("Analyse en cours..."):
                resp = client_ai.messages.create(
                    model="claude-sonnet-4-6", max_tokens=800,
                    system=sys_prompt, messages=st.session_state["messages_agent"]
                )
                answer = resp.content[0].text
                st.markdown(answer)
                st.session_state["messages_agent"].append({"role":"assistant","content":answer})

    if st.button("🗑️ Effacer"):
        st.session_state["messages_agent"] = []
        st.rerun()

# ── RAPPORT DE PILOTAGE (EC) ──
elif role == "ec" and page == "📄 Rapport de pilotage":
    df, params = check_pivot()
    dos = get_dossier(st.session_state["dossier_id"])
    st.header("📄 Rapport de pilotage analytique")
    st.markdown(f"*{dos['nom']} — {dos.get('exercice','')}*")

    if not DOCX_AVAILABLE:
        st.error("❌ python-docx non installé. Ajoutez `python-docx` dans requirements.txt.")
        st.stop()

    col1,col2 = st.columns(2)
    with col1:
        nom_editeur = st.text_input("Maison d'édition", value=dos.get("nom",""))
        nom_dirigeant = st.text_input("Nom du dirigeant", value=dos.get("dirigeant",""))
        exercice = st.text_input("Exercice", value=dos.get("exercice","2025/2026"))
    with col2:
        date_rapport = st.date_input("Date du rapport", value=_dt.date.today())
        periode_debut = st.text_input("Début de période", value=dos.get("periode_debut","01/04/2025"))
        periode_fin   = st.text_input("Fin de période",   value=dos.get("periode_fin","31/03/2026"))

    mode_anon = st.checkbox("🕶️ Anonymiser les titres", value=False)
    nb_top = st.slider("Nombre de titres Top / Flop", 3, 10, 5)

    titres_r = sorted(filtrer_isbn_reels(df)["Code_Analytique"].astype(str).unique().tolist())
    if not titres_r:
        st.error("Aucun EAN détecté.")
        st.stop()

    res_r = calculer_indicateurs_titres(df, params, titres_r)
    mapping_anon = obtenir_mapping_anonymisation(df)
    res_r["Titre"] = res_r["Code_Analytique"].apply(lambda c: mapping_anon.get(c,c) if mode_anon else c)

    df_v_ = df[mask_ventes(df,params)]; df_c_ = df[mask_charges(df,params)]
    df_r_ = df[mask_retours(df,params)]; df_rem_ = df[mask_remises(df,params)]
    df_prov = df[mask_provisions_reprises(df,params)]
    net_p = df_prov["Crédit"].sum()-df_prov["Débit"].sum()
    ca_b_ = df_v_["Crédit"].sum()-df_v_["Débit"].sum()
    tr_   = df_r_["Débit"].sum()-df_r_["Crédit"].sum()
    rem_  = df_rem_["Débit"].sum()-df_rem_["Crédit"].sum()
    ca_n_ = ca_b_-tr_-rem_
    ch_t_ = (df_c_["Débit"].sum()-df_c_["Crédit"].sum())-net_p
    mn_   = ca_n_-ch_t_
    mb_   = res_r["Marge brute"].sum()
    taux_mb_ = (mb_/ca_b_*100) if ca_b_ else 0
    n_pos_ = (res_r["Marge brute"]>0).sum()
    n_neg_ = (res_r["Marge brute"]<=0).sum()
    nb_eans_ = len(titres_r)
    res_sorted_ = res_r.sort_values("Marge brute",ascending=False).reset_index(drop=True)

    st.divider()
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("CA net", f"{fmt_fr(ca_n_)} €")
    m2.metric("Marge brute", f"{fmt_fr(mb_)} €", f"{taux_mb_:.1f} %")
    m3.metric("Résultat net", f"{fmt_fr(mn_)} €")
    m4.metric("Titres déficitaires", f"{n_neg_} / {nb_eans_}")

    if st.button("🖨️ Générer le rapport Word", type="primary", use_container_width=True):
        with st.spinner("Génération..."):
            try:
                _MR = RGBColor(0x1F,0x4E,0x79); _TR = RGBColor(0xC0,0x52,0x2A)
                _WH = RGBColor(0xFF,0xFF,0xFF)

                def _shd(cell, h):
                    tc=cell._tc; p=tc.get_or_add_tcPr()
                    for o in p.findall(qn("w:shd")): p.remove(o)
                    s=OxmlElement("w:shd"); s.set(qn("w:val"),"clear")
                    s.set(qn("w:color"),"auto"); s.set(qn("w:fill"),h); p.append(s)

                def _brd(cell,c="CCCCCC",z="4"):
                    tc=cell._tc; p=tc.get_or_add_tcPr()
                    for o in p.findall(qn("w:tcBorders")): p.remove(o)
                    b=OxmlElement("w:tcBorders")
                    for e in ["top","left","bottom","right"]:
                        x=OxmlElement(f"w:{e}"); x.set(qn("w:val"),"single")
                        x.set(qn("w:sz"),z); x.set(qn("w:color"),c); b.append(x)
                    p.append(b)

                def _wc(cell,text,bold=False,italic=False,color=None,size=9,center=False,fill=None):
                    if fill: _shd(cell,fill)
                    _brd(cell); cell.vertical_alignment=WD_ALIGN_VERTICAL.CENTER
                    p=cell.paragraphs[0]; p.clear()
                    p.alignment=WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
                    p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
                    r=p.add_run(str(text)); r.bold=bold; r.italic=italic
                    r.font.size=Pt(size); r.font.name="Arial"
                    r.font.color.rgb=color if color else RGBColor(0,0,0)

                def _P(d,text,bold=False,italic=False,size=10,color=None,align="left",sb=6,sa=6):
                    p=d.add_paragraph()
                    p.paragraph_format.space_before=Pt(sb); p.paragraph_format.space_after=Pt(sa)
                    p.alignment={"left":WD_ALIGN_PARAGRAPH.LEFT,"center":WD_ALIGN_PARAGRAPH.CENTER,
                                 "justify":WD_ALIGN_PARAGRAPH.JUSTIFY}.get(align,WD_ALIGN_PARAGRAPH.LEFT)
                    r=p.add_run(text); r.bold=bold; r.italic=italic
                    r.font.size=Pt(size); r.font.name="Arial"
                    if color: r.font.color.rgb=color
                    return p

                def _sep(d):
                    p=d.add_paragraph()
                    p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(6)
                    pr=p._p.get_or_add_pPr(); pb=OxmlElement("w:pBdr"); bt=OxmlElement("w:bottom")
                    bt.set(qn("w:val"),"single"); bt.set(qn("w:sz"),"6"); bt.set(qn("w:color"),"C0522A")
                    pb.append(bt); pr.append(pb)

                def _fn(n,sign=False):
                    return ("+" if (sign and n>=0) else "")+f"{n:,.0f} EUR".replace(",","  ")

                top_list_ = [res_sorted_.iloc[i] for i in range(min(nb_top,len(res_sorted_)))]
                flop_list_= [res_sorted_.iloc[-(i+1)] for i in range(min(nb_top,len(res_sorted_)))]

                d=Document()
                for s in d.sections:
                    s.top_margin=Cm(2); s.bottom_margin=Cm(2)
                    s.left_margin=Cm(2.5); s.right_margin=Cm(2)

                _P(d,"CAB EDITION",bold=True,size=14,color=_MR,sb=0,sa=2)
                _P(d,"Expert-comptable — Maisons d'édition indépendantes",size=10,sb=0,sa=4)

                _mois_fr_={1:"janvier",2:"février",3:"mars",4:"avril",5:"mai",6:"juin",
                            7:"juillet",8:"août",9:"septembre",10:"octobre",11:"novembre",12:"décembre"}
                date_fr_=f"{date_rapport.day} {_mois_fr_[date_rapport.month]} {date_rapport.year}"
                pr_=d.add_paragraph(); pr_.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                pr_.paragraph_format.space_before=Pt(0); pr_.paragraph_format.space_after=Pt(4)
                rr_=pr_.add_run(f"Lille, le {date_fr_}"); rr_.font.size=Pt(10); rr_.font.name="Arial"
                _sep(d)

                _P(d,f"{nom_dirigeant}",bold=True,size=11,sb=4,sa=2)
                _P(d,nom_editeur,size=11,sb=0,sa=10)

                to_=d.add_table(rows=1,cols=1); to_.style="Table Grid"
                co_=to_.rows[0].cells[0]; _shd(co_,"D6E4F0"); _brd(co_)
                po_=co_.paragraphs[0]; po_.paragraph_format.space_before=Pt(4); po_.paragraph_format.space_after=Pt(4)
                r1_=po_.add_run("Objet : "); r1_.bold=True; r1_.font.size=Pt(10); r1_.font.name="Arial"; r1_.font.color.rgb=_MR
                r2_=po_.add_run(f"Rapport de pilotage analytique — Exercice {exercice}"); r2_.font.size=Pt(10); r2_.font.name="Arial"

                _P(d,"",sb=10,sa=0)
                _P(d,f"{nom_dirigeant},",bold=True,size=11,sb=0,sa=6)
                _P(d,f"Conformément aux diligences prévues dans notre mission d'accompagnement au pilotage analytique, "
                   f"nous avons l'honneur de vous adresser le présent rapport de pilotage relatif à l'exercice {exercice}, "
                   f"pour la période du {periode_debut} au {periode_fin}.",
                   size=10,align="justify",sb=0,sa=6)
                _P(d,f"Ce rapport porte sur l'ensemble des {nb_eans_} ISBN actifs de votre catalogue, "
                   f"traités par notre outil Python VISION EDITION.",
                   size=10,align="justify",sb=0,sa=10)

                # Section I
                _P(d,"I.   SYNTHÈSE DE LA PERFORMANCE ANALYTIQUE",bold=True,size=12,color=_MR,sb=8,sa=4)
                _sep(d)
                ts_=d.add_table(rows=5,cols=3); ts_.style="Table Grid"
                ts_.rows[0].cells[0]._tc; _wc(ts_.rows[0].cells[0],"Indicateur",bold=True,color=_WH,size=9,fill="1F4E79")
                _wc(ts_.rows[0].cells[1],"Montant",bold=True,color=_WH,size=9,center=True,fill="1F4E79")
                _wc(ts_.rows[0].cells[2],"Commentaire",bold=True,color=_WH,size=9,fill="1F4E79")
                for i,(lbl,val,cmt,fl) in enumerate([
                    ("CA net éditeur",_fn(ca_n_),f"{nb_eans_} ISBN actifs","D6E4F0"),
                    ("Marge brute sur ventes",_fn(mb_,True),f"Taux : {taux_mb_:.1f}%","E2EFDA"),
                    (f"Titres déficitaires",f"{n_neg_} / {nb_eans_}","Action corrective requise","FFDDD9"),
                    ("Résultat analytique net",_fn(mn_,True),"Après charges de structure","FFDDD9" if mn_<0 else "E2EFDA"),
                ]):
                    row_=ts_.rows[i+1]
                    _wc(row_.cells[0],lbl,bold=True,size=9,fill=fl)
                    _wc(row_.cells[1],val,bold=True,center=True,size=10,fill=fl)
                    _wc(row_.cells[2],cmt,italic=True,size=8,fill=fl)

                _P(d,"",sb=10,sa=0)
                # Section II — Top & Flop
                _P(d,"II.   ANALYSE PAR TITRE",bold=True,size=12,color=_MR,sb=8,sa=4)
                _sep(d)
                _P(d,f"Top {nb_top} — titres les plus contributifs",bold=True,size=10,color=_MR,sb=4,sa=4)
                tt_=d.add_table(rows=nb_top+1,cols=4); tt_.style="Table Grid"
                for c_,h_ in enumerate(["Rg","Titre","CA net (EUR)","Marge brute (EUR)"]):
                    _wc(tt_.rows[0].cells[c_],h_,bold=True,color=_WH,size=9,fill="1F4E79")
                for i,rd_ in enumerate(top_list_):
                    bg_=("F5E6DF" if i%2==0 else "FFFFFF")
                    _wc(tt_.rows[i+1].cells[0],str(i+1),bold=True,center=True,size=9,fill=bg_)
                    _wc(tt_.rows[i+1].cells[1],str(rd_["Titre"])[:44],bold=True,size=9,fill=bg_)
                    _wc(tt_.rows[i+1].cells[2],_fn(rd_["CA net"]),center=True,size=9,fill=bg_)
                    _wc(tt_.rows[i+1].cells[3],_fn(rd_["Marge brute"],True),bold=True,center=True,size=10,
                        color=RGBColor(0x21,0x73,0x46),fill="E2EFDA")

                _P(d,"",sb=8,sa=0)
                _P(d,f"Flop {nb_top} — titres déficitaires",bold=True,size=10,color=RGBColor(0xC0,0,0),sb=4,sa=4)
                tf_=d.add_table(rows=nb_top+1,cols=4); tf_.style="Table Grid"
                for c_,h_ in enumerate(["Rg","Titre","CA net (EUR)","Marge brute (EUR)"]):
                    _wc(tf_.rows[0].cells[c_],h_,bold=True,color=_WH,size=9,fill="C00000")
                for i,rd_ in enumerate(flop_list_):
                    bg_=("FFF5F5" if i%2==0 else "FFFFFF")
                    _wc(tf_.rows[i+1].cells[0],str(i+1),bold=True,center=True,size=9,fill=bg_)
                    _wc(tf_.rows[i+1].cells[1],str(rd_["Titre"])[:44],bold=True,size=9,fill=bg_)
                    _wc(tf_.rows[i+1].cells[2],_fn(rd_["CA net"]),center=True,size=9,fill=bg_)
                    _wc(tf_.rows[i+1].cells[3],_fn(rd_["Marge brute"],True),bold=True,center=True,size=10,
                        color=RGBColor(0xC0,0,0),fill="FFDDD9")

                _P(d,"",sb=10,sa=0)
                # Conclusion
                _P(d,"III.   RECOMMANDATIONS",bold=True,size=12,color=_MR,sb=8,sa=4)
                _sep(d)
                _P(d,"Au vu de l'analyse ci-avant, nous vous recommandons d'engager rapidement une analyse "
                   "approfondie des titres déficitaires et de leur impact sur la structure de votre catalogue. "
                   "Nous demeurons à votre entière disposition pour en discuter lors de notre prochaine réunion trimestrielle.",
                   size=10,align="justify",sb=0,sa=16)
                _P(d,"Veuillez agréer, Monsieur, l'expression de nos salutations distinguées.",size=10,sb=0,sa=20)
                _P(d,"CAB EDITION — L'Expert-comptable",bold=True,size=10,color=_MR,sb=0,sa=4)
                _P(d,f"Date : {date_fr_}",size=9,sb=0,sa=0)

                # Note méthodologique
                _P(d,"",sb=10,sa=0)
                pn_=d.add_paragraph(); pn_.paragraph_format.space_before=Pt(4); pn_.paragraph_format.space_after=Pt(4)
                rn1_=pn_.add_run("Note méthodologique : "); rn1_.bold=True; rn1_.font.size=Pt(8); rn1_.font.name="Arial"
                rn1_.font.color.rgb=RGBColor(0x55,0x55,0x55)
                sfx=" (titres anonymisés)" if mode_anon else ""
                rn2_=pn_.add_run(f"Rapport généré par Python VISION EDITION — {periode_debut} au {periode_fin} — "
                                  f"{nb_eans_} ISBN actifs{sfx} — données provisoires.")
                rn2_.italic=True; rn2_.font.size=Pt(8); rn2_.font.name="Arial"
                rn2_.font.color.rgb=RGBColor(0x77,0x77,0x77)

                buf_w=BytesIO(); d.save(buf_w); buf_w.seek(0)
                fn_w=f"Rapport_{nom_editeur[:20].replace(' ','_')}_{exercice.replace('/','_')}.docx"
                st.success("✅ Rapport généré !")
                st.download_button("⬇️ Télécharger le rapport Word", data=buf_w, file_name=fn_w,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True, type="primary")

            except Exception as e:
                st.error(f"❌ Erreur : {e}")
                st.exception(e)


# ── TABLEAU DE BORD DIRIGEANT ──
elif role == "dirigeant" and page == "🏠 Mon tableau de bord":
    df, params, dos = get_df_dirigeant()
    st.title(f"📚 {dos['nom']}")
    st.markdown(f"*Tableau de bord de pilotage — Exercice {dos.get('exercice','')}*")
    st.divider()

    df_v_ = df[mask_ventes(df,params)]; df_r_ = df[mask_retours(df,params)]
    df_rem_ = df[mask_remises(df,params)]; df_c_ = df[mask_charges(df,params)]
    pref_d_ = tuple(params.get("ventes_distributeur") or params.get("ventes",[]))
    df_vd_ = df[df["Compte"].astype(str).str.startswith(pref_d_)] if pref_d_ else df_v_
    ca_d_ = df_vd_["Crédit"].sum()
    ca_b_ = df_v_["Crédit"].sum()
    tr_   = df_r_["Débit"].sum()-df_r_["Crédit"].sum()
    rem_  = df_rem_["Débit"].sum()-df_rem_["Crédit"].sum()
    ca_n_ = ca_b_-tr_-rem_
    ch_t_ = df_c_["Débit"].sum()-df_c_["Crédit"].sum()
    res_  = ca_n_-ch_t_
    taux_r_ = (tr_/ca_d_*100) if ca_d_ else 0

    st.markdown("### Vos indicateurs clés")
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("💶 Chiffre d'affaires net", f"{fmt_fr(ca_n_)} €")
    c2.metric("📊 Résultat net", f"{fmt_fr(res_)} €",
              delta="✅ Bénéficiaire" if res_>0 else "⚠️ Déficitaire",
              delta_color="normal" if res_>0 else "inverse")
    c3.metric("🔁 Taux de retour", f"{taux_r_:.1f} %",
              delta="✅ Normal" if taux_r_<25 else "⚠️ Élevé",
              delta_color="normal" if taux_r_<25 else "inverse")
    titres_d = filtrer_isbn_reels(df)["Code_Analytique"].nunique()
    c4.metric("📚 Titres actifs", str(titres_d))

    st.divider()
    # Évolution mensuelle simplifiée
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df_ok = df.dropna(subset=["Date"]).copy()
    df_ok["Mois"] = df_ok["Date"].dt.to_period("M").astype(str)
    df_v_ok = df_ok[mask_ventes(df_ok,params)]
    df_r_ok = df_ok[mask_retours(df_ok,params)]
    tv_ = df_v_ok.groupby("Mois")["Crédit"].sum().reset_index().rename(columns={"Crédit":"CA"})
    tr2_ = df_r_ok.groupby("Mois")["Débit"].sum().reset_index().rename(columns={"Débit":"Retours"})
    trend_ = tv_.merge(tr2_,on="Mois",how="left").fillna(0)
    trend_["CA net"] = trend_["CA"]-trend_["Retours"]

    st.markdown("### Évolution de votre chiffre d'affaires")
    fig_d = px.line(trend_, x="Mois", y=["CA","CA net"], markers=True, height=300,
                    labels={"value":"€","variable":""},
                    color_discrete_map={"CA":"#3B82F6","CA net":"#10B981"})
    fig_d.update_layout(margin=dict(t=10), legend=dict(orientation="h"))
    st.plotly_chart(fig_d, use_container_width=True)

    # Message de bienvenue personnalisé
    st.info(f"💡 Pour analyser vos titres en détail, rendez-vous dans **📖 Mes titres**. "
            f"Pour poser une question, utilisez **🤖 Mon assistant IA**.")

# ── MES TITRES (DIRIGEANT) ──
elif role == "dirigeant" and page == "📖 Mes titres":
    df, params, dos = get_df_dirigeant()
    st.header(f"📖 Mes titres — {dos['nom']}")

    titres_dir = sorted(filtrer_isbn_reels(df)["Code_Analytique"].astype(str).unique().tolist())
    if not titres_dir:
        st.warning("Aucun titre détecté dans vos données.")
        st.stop()

    indic_dir = calculer_indicateurs_titres(df, params, titres_dir)

    # Vue simplifiée pour le dirigeant
    st.markdown("### Classement de vos titres")
    indic_dir_aff = indic_dir[["Code_Analytique","Signal","CA net","Marge brute","Taux retour (%)","Résultat net"]].copy()
    indic_dir_aff = indic_dir_aff.sort_values("Marge brute", ascending=False)
    indic_dir_aff.columns = ["ISBN / Titre","Signal","CA net (€)","Marge brute (€)","Taux retour (%)","Résultat net (€)"]

    st.dataframe(
        indic_dir_aff.style.format({
            "CA net (€)":(lambda x: fmt_fr(x)),
            "Marge brute (€)":(lambda x: fmt_fr(x)),
            "Taux retour (%)":(lambda x: f"{x:.1f} %"),
            "Résultat net (€)":(lambda x: fmt_fr(x)),
        }),
        use_container_width=True, hide_index=True
    )

    st.divider()
    st.markdown("### Focus sur un titre")
    isbn_dir = st.selectbox("Choisir un titre", titres_dir)
    row_dir = indic_dir[indic_dir["Code_Analytique"]==isbn_dir].iloc[0]

    c1,c2,c3 = st.columns(3)
    c1.metric("CA net", f"{fmt_fr(row_dir['CA net'])} €")
    c2.metric("Marge brute", f"{fmt_fr(row_dir['Marge brute'])} €")
    c3.metric("Taux de retour", f"{row_dir['Taux retour (%)']:.1f} %")
    st.markdown(f"**Signal : {row_dir['Signal']}**")
    if row_dir["Signal"] == "🔴":
        st.warning("Ce titre présente des difficultés. Parlez-en à votre expert-comptable.")
    elif row_dir["Signal"] == "🟡":
        st.info("Ce titre est à surveiller. Le taux de retour mérite attention.")
    else:
        st.success("Ce titre est rentable et performant.")

# ── MA TRÉSORERIE (DIRIGEANT) ──
elif role == "dirigeant" and page == "💰 Ma trésorerie":
    df, params, dos = get_df_dirigeant()
    st.header(f"💰 Ma trésorerie — {dos['nom']}")

    if "treso_real" not in st.session_state:
        st.info("Votre expert-comptable n'a pas encore généré les projections de trésorerie. "
                "Elles seront disponibles après la prochaine mise à jour de vos données.")
        st.stop()

    treso_ = st.session_state["treso_real"]
    _MOIS = {1:"Jan",2:"Fév",3:"Mar",4:"Avr",5:"Mai",6:"Jui",7:"Jul",8:"Aoû",9:"Sep",10:"Oct",11:"Nov",12:"Déc"}
    def _ml(p): return f"{_MOIS.get(p.month,str(p.month))} {p.year}"

    st.metric("Trésorerie actuelle", f"{fmt_fr(treso_.iloc[-1])} €",
              delta="✅ Positive" if treso_.iloc[-1]>0 else "⚠️ Négative",
              delta_color="normal" if treso_.iloc[-1]>0 else "inverse")

    fig_tr = go.Figure()
    fig_tr.add_trace(go.Scatter(x=[_ml(m) for m in treso_.index], y=treso_.values,
                                fill="tozeroy", name="Trésorerie",
                                line=dict(color="#3B82F6",width=2)))
    fig_tr.add_hline(y=0, line_dash="dash", line_color="red", annotation_text="Seuil zéro")
    fig_tr.update_layout(height=320, margin=dict(t=10), xaxis_title="", yaxis_title="€")
    st.plotly_chart(fig_tr, use_container_width=True)
    st.caption("Trésorerie reconstituée à partir de votre comptabilité analytique par votre cabinet.")

# ── RETOURS (DIRIGEANT) ──
elif role == "dirigeant" and page == "📦 Retours & Remises":
    df, params, dos = get_df_dirigeant()
    st.header(f"📦 Retours & Remises — {dos['nom']}")

    df_r_ = df[mask_retours(df,params)]; df_v_ = df[mask_ventes(df,params)]
    pref_d_ = tuple(params.get("ventes_distributeur") or params.get("ventes",[]))
    df_vd_ = df[df["Compte"].astype(str).str.startswith(pref_d_)] if pref_d_ else df_v_
    ca_d_ = df_vd_["Crédit"].sum()
    tr_ = df_r_["Débit"].sum()-df_r_["Crédit"].sum() if not df_r_.empty else 0
    taux_r_ = (tr_/ca_d_*100) if ca_d_ else 0

    c1,c2 = st.columns(2)
    c1.metric("Total retours", f"{fmt_fr(tr_)} €")
    c2.metric("Taux de retour", f"{taux_r_:.1f} %",
              delta="✅ Normal" if taux_r_<25 else "⚠️ Élevé",
              delta_color="normal" if taux_r_<25 else "inverse")

    if taux_r_ > 30:
        st.error("⚠️ Votre taux de retour est élevé. Votre expert-comptable vous contactera pour en discuter.")
    elif taux_r_ > 25:
        st.warning("Votre taux de retour mérite attention. Parlez-en lors de votre prochaine réunion.")

    if not df_r_.empty:
        df_r_["Date"] = pd.to_datetime(df_r_["Date"], errors="coerce")
        df_r_["Mois"] = df_r_["Date"].dt.strftime("%Y-%m")
        trend_r_ = df_r_.groupby("Mois")["Débit"].sum().abs().reset_index()
        fig_r_ = px.bar(trend_r_, x="Mois", y="Débit", title="Retours par mois",
                        color_discrete_sequence=["#EF4444"], height=280)
        st.plotly_chart(fig_r_, use_container_width=True)

# ── DROITS D'AUTEURS (DIRIGEANT) ──
elif role == "dirigeant" and page == "✍️ Droits d'auteurs":
    df, params, dos = get_df_dirigeant()
    st.header(f"✍️ Droits d'auteurs — {dos['nom']}")
    st.info("Cette section affiche vos droits d'auteurs comptabilisés. "
            "Pour un relevé détaillé par auteur, contactez votre cabinet.")

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    mask_db = df["Compte"].astype(str).str.strip()=="604300000"
    mask_av = df["Compte"].astype(str).str.strip()=="409600"

    total_db = df[mask_db]["Débit"].sum()-df[mask_db]["Crédit"].sum()
    av_d = df[mask_av]["Débit"].sum(); av_c = df[mask_av]["Crédit"].sum()
    av_restant = av_d-av_c

    c1,c2 = st.columns(2)
    c1.metric("Droits d'auteurs comptabilisés", f"{fmt_fr(total_db,2)} €")
    c2.metric("À-valoirs restant à amortir", f"{fmt_fr(av_restant,2)} €",
              delta="⚠️ À surveiller" if av_restant>0 else "✅ Amortis",
              delta_color="off")

    if av_restant > 0:
        st.info(f"Un à-valoir de {fmt_fr(av_restant,2)} € reste à amortir sur votre catalogue. "
                f"Votre expert-comptable suit cet indicateur pour vous alerter si nécessaire.")

# ── ASSISTANT IA DIRIGEANT ──
elif role == "dirigeant" and page == "🤖 Mon assistant IA":
    df, params, dos = get_df_dirigeant()
    st.header(f"🤖 Mon assistant — {dos['nom']}")
    st.markdown("*Posez vos questions sur votre activité éditoriale en langage naturel.*")

    client_ai = get_client_ai()
    if client_ai is None:
        st.error("⚠️ Assistant IA non disponible. Contactez votre cabinet.")
        st.stop()

    # Contexte simplifié pour le dirigeant
    df_v_ = df[mask_ventes(df,params)]; df_r_ = df[mask_retours(df,params)]
    ca_b_ = df_v_["Crédit"].sum()
    tr_   = df_r_["Débit"].sum()-df_r_["Crédit"].sum()
    ca_n_ = ca_b_-tr_
    taux_r_ = (tr_/ca_b_*100) if ca_b_ else 0
    titres_dir_ = sorted(filtrer_isbn_reels(df)["Code_Analytique"].astype(str).unique().tolist())
    indic_dir_ = calculer_indicateurs_titres(df,params,titres_dir_) if titres_dir_ else pd.DataFrame()

    top5_ = ""
    if not indic_dir_.empty:
        top5_rows = indic_dir_.sort_values("CA net",ascending=False).head(5)
        top5_ = "\n".join([f"- {r['Code_Analytique']} : CA net {fmt_fr(r['CA net'])} EUR, "
                           f"marge brute {fmt_fr(r['Marge brute'])} EUR, retour {r['Taux retour (%)']:.1f}%"
                           for _,r in top5_rows.iterrows()])

    data_ctx_dir = f"""
DONNÉES DE {dos['nom']} — {dos.get('exercice','')}
CA brut : {fmt_fr(ca_b_)} EUR
Total retours : {fmt_fr(tr_)} EUR  
CA net : {fmt_fr(ca_n_)} EUR
Taux de retour global : {taux_r_:.1f} %
Nombre de titres actifs : {len(titres_dir_)}

Top 5 titres par CA net :
{top5_}
"""
    sys_dir = SYSTEM_PROMPT_DIRIGEANT + f"\n\nDONNÉES :\n{data_ctx_dir}"

    st.markdown("**Questions fréquentes :**")
    col_q1, col_q2, col_q3 = st.columns(3)
    questions_dir = [
        ("Quel est mon CA par titre ?", col_q1),
        ("Quel est mon taux de retour ?", col_q2),
        ("Quels sont mes titres rentables ?", col_q3),
    ]
    for q_text, col_q in questions_dir:
        if col_q.button(q_text, use_container_width=True):
            with st.spinner("Réflexion..."):
                resp_q = client_ai.messages.create(
                    model="claude-sonnet-4-6", max_tokens=500,
                    system=sys_dir, messages=[{"role":"user","content":q_text}]
                )
                st.info(f"**{q_text}**")
                st.success(resp_q.content[0].text)

    st.divider()
    for msg in st.session_state["messages_agent"]:
        with st.chat_message(msg["role"], avatar="🧑" if msg["role"]=="user" else "📚"):
            st.markdown(msg["content"])

    if prompt_dir := st.chat_input("Posez votre question..."):
        st.session_state["messages_agent"].append({"role":"user","content":prompt_dir})
        with st.chat_message("user", avatar="🧑"):
            st.markdown(prompt_dir)
        with st.chat_message("assistant", avatar="📚"):
            with st.spinner("Je cherche la réponse..."):
                resp_dir = client_ai.messages.create(
                    model="claude-sonnet-4-6", max_tokens=600,
                    system=sys_dir, messages=st.session_state["messages_agent"]
                )
                ans_dir = resp_dir.content[0].text
                st.markdown(ans_dir)
                st.session_state["messages_agent"].append({"role":"assistant","content":ans_dir})

    if st.button("🗑️ Effacer la conversation"):
        st.session_state["messages_agent"] = []
        st.rerun()

# ============================================================
# FOOTER
# ============================================================
st.divider()
st.markdown(
    "<div style='text-align:center;color:#888;font-size:12px'>"
    "© 2025 Nicolas CUISSET — VISION EDITION — Mémoire d'expertise comptable"
    "</div>",
    unsafe_allow_html=True
)
