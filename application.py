import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go
import qrcode
from PIL import Image
import anthropic
import json
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

# =====================
# FORMAT NUMERIQUE FRANCAIS
# =====================
def fmt_fr(x, decimals=0):
    """Formate un nombre selon la convention francaise : espace pour le
    separateur de milliers, virgule pour le separateur decimal (au lieu du
    format Python par defaut, qui utilise la virgule pour les milliers et
    peut induire en erreur un lecteur francais sur les montants en milliers)."""
    try:
        s = f"{float(x):,.{decimals}f}"
    except (ValueError, TypeError):
        return str(x)
    return s.replace(",", "§").replace(".", ",").replace("§", " ")


# =====================
# CONFIGURATION PAGE
# =====================
st.set_page_config(
    page_title="Outil éditorial - Maisons d'édition indépendantes",
    page_icon="📚",
    layout="wide"
)

APP_URL = "https://outilaccompagnementmaisonsedition-lyvgltfbwtqo4m9tdmzofu.streamlit.app/"

# =====================
# AUTHENTIFICATION
# =====================
if "login" not in st.session_state:
    st.session_state["login"] = False
if "page" not in st.session_state:
    st.session_state["page"] = "🏠 Accueil"
if "messages_agent" not in st.session_state:
    st.session_state["messages_agent"] = []

def login(username, password):
    users = {
        "aurore": {"password": "12345", "name": "Aurore Demoulin"},
        "laure.froidefond": {"password": "Laure2019$", "name": "Laure Froidefond"},
        "Bruno": {"password": "Toto1963$", "name": "Toto El Gringo"}
    }
    if username in users and password == users[username]["password"]:
        st.session_state["login"] = True
        st.session_state["username"] = username
        st.session_state["name"] = users[username]["name"]
        st.session_state["page"] = "🏠 Accueil"
        st.rerun()
    else:
        st.error("❌ Identifiants incorrects")

if not st.session_state["login"]:
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("## 📚 Outil d'accompagnement éditorial")
        st.markdown("*Maisons d'édition indépendantes*")
        st.divider()
        username_input = st.text_input("Identifiant", placeholder="Votre identifiant")
        password_input = st.text_input("Mot de passe", type="password", placeholder="Votre mot de passe")
        if st.button("🔑 Connexion", use_container_width=True, type="primary"):
            login(username_input, password_input)
        st.divider()
        # QR code sur la page de login aussi
        st.markdown("**Accès mobile — scannez pour ouvrir l'app**")
        qr = qrcode.QRCode(version=1, box_size=4, border=2)
        qr.add_data(APP_URL)
        qr.make(fit=True)
        img_qr = qr.make_image(fill_color="black", back_color="white")
        buf = BytesIO()
        img_qr.save(buf, format="PNG")
        st.image(buf.getvalue(), width=140)
    st.stop()

# =====================
# SIDEBAR
# =====================
with st.sidebar:
    st.markdown(f"👤 **{st.session_state['name']}**")
    st.divider()
    pages = [
        "🏠 Accueil",
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
    page = st.selectbox("Navigation", pages)
    st.divider()
    st.session_state["mode_anonyme"] = st.checkbox(
        "🕶️ Mode démonstration (titres anonymisés)",
        value=st.session_state.get("mode_anonyme", False),
        help="Remplace les titres réels par des identifiants anonymes (T1, T2...) dans les "
             "graphiques et libellés affichés, sans modifier ni les données ni les calculs. "
             "Utile pour réaliser des captures d'écran ou de la documentation sans exposer "
             "le catalogue réel du client."
    )
    st.divider()
    if st.button("🚪 Déconnexion", use_container_width=True):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

# =====================
# HELPERS
# =====================
def get_client():
    """Retourne le client Anthropic — la clé est injectée par Streamlit Cloud via secrets."""
    try:
        return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])
    except Exception:
        return None

def filtrer_isbn_reels(df):
    """Ne garde que les lignes réellement rattachées à un titre (un vrai ISBN).

    Sans ce filtre, deux types de lignes viennent polluer les classements/aggrégations
    par Code_Analytique :
    - les lignes hors comptes de charges/produits (tiers, banque, TVA...) qui n'ont jamais
      reçu de code analytique et remontent avec un Code_Analytique vide ;
    - les lignes "CHARGES INDIRECTES" / "PRODUITS INDIRECTS" (libellés globaux, pas des ISBN),
      tant qu'elles n'ont pas été réparties sur les titres actifs.

    Si le jeu de données utilise plusieurs familles analytiques en parallèle (cas du grand
    livre réel : EDITION / COMMUNICATION / Types de dépenses), on restreint en plus à la
    famille EDITION pour ignorer les lignes des autres familles qui pourraient elles aussi
    porter un code non vide. Si aucune ligne ne porte littéralement "EDITION" (cas des jeux
    de données de démonstration, qui ne suivent pas cette convention), cette restriction
    supplémentaire est ignorée pour rester compatible.
    """
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
    """Lignes de retours, à l'exclusion des lignes de remises. Sans cette exclusion, si le
    compte remises (ex. 7091) est un sous-compte du compte retours (ex. 709 — cas fréquent
    du plan comptable, 709 = "Rabais, remises, ristournes" avec 7091 en sous-compte),
    startswith("709") capte AUSSI les lignes de remises : elles seraient alors soustraites
    deux fois du CA net (une fois comme "retours", une fois comme "remises")."""
    prefixes_retours = tuple(params.get("retours") or [])
    if not prefixes_retours:
        return pd.Series(False, index=df_scope.index)
    mask = df_scope["Compte"].astype(str).str.startswith(prefixes_retours)
    prefixes_remises = tuple(params.get("remises") or [])
    if prefixes_remises:
        mask = mask & (~df_scope["Compte"].astype(str).str.startswith(prefixes_remises))
    return mask

def mask_remises(df_scope, params):
    """Lignes de remises (comptes configurés dans params["remises"])."""
    prefixes_remises = tuple(params.get("remises") or [])
    if not prefixes_remises:
        return pd.Series(False, index=df_scope.index)
    return df_scope["Compte"].astype(str).str.startswith(prefixes_remises)

# Comptes "placeholder" (non numériques) générés par la répartition des charges/produits
# indirects sur les titres actifs (cf. ⚙️ Paramétrage analytique → Répartition). Une fois
# la répartition activée, les lignes "CHARGES INDIRECTES"/"PRODUITS INDIRECTS" d'origine
# (portées sur de vrais comptes 6xx/7xx) sont RETIRÉES du pivot et remplacées par ces
# quote-parts par titre — qui ne matchent plus un préfixe "6"/"701" classique. Sans en
# tenir compte, les totaux globaux (Tableau de bord, Synthèse financière) sous-évaluent
# les charges/produits exactement du montant réparti.
COMPTE_CHARGES_INDIRECTES_REPARTIES = "CHARGES INDIRECTES REPARTIES"
COMPTE_PRODUITS_INDIRECTS_REPARTIS = "PRODUITS INDIRECTS REPARTIS"

def mask_provisions_reprises(df_scope, params):
    """Reprises sur provisions (ex. 781/7810 = reprise de provision pour retour), configurées
    séparément (params["provisions_reprises"]) : à la demande du client, ces montants ne sont
    PAS traités comme du CA/produit — ils sont nettés directement contre les charges (en
    contrepartie de la dotation initiale, cf. la nature "Provision pour retour" du détail des
    charges) plutôt que de gonfler artificiellement le chiffre d'affaires."""
    prefixes = tuple(params.get("provisions_reprises") or [])
    if not prefixes:
        return pd.Series(False, index=df_scope.index)
    return df_scope["Compte"].astype(str).str.startswith(prefixes)

def mask_ventes(df_scope, params):
    """Lignes de ventes/produits : comptes configurés (params["ventes"]), la quote-part de
    produits indirects répartis sur les titres actifs si la répartition a été activée (cf.
    COMPTE_PRODUITS_INDIRECTS_REPARTIS ci-dessus), et — à la demande du client, pour ne pas
    introduire une catégorie "autres produits" distincte — tout autre produit comptabilisé
    sur un compte 7xx qui n'est ni un retour, ni une remise, ni une reprise sur provisions
    (cf. mask_provisions_reprises, nettée contre les charges plutôt que comptée en CA) :
    commissions libraires (706), produits divers de gestion courante (708/758), variation de
    stock de produits finis (7134), subventions (740), produits financiers (768)...

    Ces montants sont bien réels (près de 67 k€ sur le grand livre du cas d'étude, dont
    seulement ~8,8 k€ tagués "PRODUITS INDIRECTS" et donc déjà répartis) : sans les inclure
    ici, ils disparaissent purement et simplement du résultat net du Tableau de bord et de la
    Synthèse financière, qui ne se réconcilient alors plus avec le résultat comptable réel
    (somme de tous les comptes 6 et 7 du grand livre). Conformément au principe retenu, ils
    sont directement intégrés au CA (comme n'importe quelle vente), plutôt qu'isolés dans une
    ligne séparée."""
    prefixes_ventes = tuple(params.get("ventes") or [])
    mask_configuree = (df_scope["Compte"].astype(str).str.startswith(prefixes_ventes)
                        if prefixes_ventes else pd.Series(False, index=df_scope.index))
    mask_repartis = df_scope["Compte"].astype(str) == COMPTE_PRODUITS_INDIRECTS_REPARTIS
    is_7 = df_scope["Compte"].astype(str).str.startswith("7")
    mask_autres_7xx = (is_7 & (~mask_retours(df_scope, params)) & (~mask_remises(df_scope, params))
                        & (~mask_provisions_reprises(df_scope, params)))
    return mask_configuree | mask_repartis | mask_autres_7xx

def mask_charges(df_scope, params):
    """Lignes de charges : comptes configurés (params["charges"]), plus la quote-part de
    charges indirectes réparties sur les titres actifs si la répartition a été activée
    (cf. COMPTE_CHARGES_INDIRECTES_REPARTIES ci-dessus)."""
    prefixes_charges = tuple(params.get("charges") or [])
    mask = (df_scope["Compte"].astype(str).str.startswith(prefixes_charges)
            if prefixes_charges else pd.Series(False, index=df_scope.index))
    return mask | (df_scope["Compte"].astype(str) == COMPTE_CHARGES_INDIRECTES_REPARTIES)

def normaliser_codes_ean(df, col="Code_Analytique"):
    """Fusionne les lignes dont le code analytique commence par le même numéro EAN mais
    diffère par un libellé légèrement différent après le tiret (casse, troncature, variante
    linguistique — ex. "9782376801436 - Villa Cavrois" vs "9782376801436 - VILLA CAVROIS").
    Sans cette normalisation, un même titre peut être scindé en deux codes analytiques
    distincts du seul fait d'une incohérence de saisie, faussant tous les classements et
    fiches par titre qui s'appuient sur ce code (cf. le cas déjà rencontré du préfixe
    "ISBN " dupliqué). Les codes vides ou les libellés globaux (charges/produits indirects)
    ne sont pas concernés. Le libellé conservé est le plus long des libellés observés pour
    un même numéro EAN, par simple convention (généralement le moins tronqué)."""
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
    """Construit (et met en cache en session) un mapping stable code réel -> identifiant
    anonyme (T1, T2...), basé sur la liste triée des titres actifs. Recalculé uniquement si
    l'ensemble des titres change, pour que chaque titre garde le même identifiant d'un
    rafraîchissement à l'autre pendant la session."""
    titres = sorted(filtrer_isbn_reels(df)["Code_Analytique"].astype(str).unique().tolist())
    cache = st.session_state.get("anonymisation_cache")
    if cache and cache.get("titres_source") == titres:
        return cache["mapping"]
    mapping = {t: f"T{i+1}" for i, t in enumerate(titres)}
    st.session_state["anonymisation_cache"] = {"titres_source": titres, "mapping": mapping}
    return mapping

def label_affiche(code, df_pour_mapping=None):
    """Retourne le libellé à afficher pour un code analytique : anonymisé (Txxx) si le mode
    démonstration est actif dans la barre latérale, sinon le code réel tel quel. Ne modifie
    jamais les données ni les calculs sous-jacents — uniquement l'étiquette affichée à
    l'écran (graphiques, en-têtes, sélecteurs de titre)."""
    if not st.session_state.get("mode_anonyme"):
        return code
    mapping = st.session_state.get("anonymisation_cache", {}).get("mapping")
    if mapping is None and df_pour_mapping is not None:
        mapping = obtenir_mapping_anonymisation(df_pour_mapping)
    return (mapping or {}).get(code, code)

def _dialog_decorator(title, width="large"):
    """Retourne le décorateur st.dialog si disponible (Streamlit ≥ 1.31), sinon un
    expander déplié en repli pour rester compatible avec une version plus ancienne."""
    if hasattr(st, "dialog"):
        return st.dialog(title, width=width)
    def _wrap(func):
        def _inner(*args, **kwargs):
            with st.expander(f"🪟 {title}", expanded=True):
                return func(*args, **kwargs)
        return _inner
    return _wrap

@_dialog_decorator("📖 Fiche titre", width="large")
def afficher_fiche_titre(isbn_sel, df, params):
    """Fiche détaillée d'un titre (mini SIG) affichée dans une fenêtre modale dédiée."""
    df_t = df[df["Code_Analytique"] == isbn_sel]
    df_v   = df_t[df_t["Compte"].astype(str).str.startswith(tuple(params["ventes"]))]
    # CA distributeur (ex. compte BLDD 7011 pour ce cas d'étude — configurable, non limité à un
    # distributeur particulier) : base STRICTE du taux de retour/remise de ce titre — distincte
    # du CA net du titre (df_v ci-dessus, ventes larges 701) utilisé pour la marge/le résultat.
    df_v_distrib_t = df_t[df_t["Compte"].astype(str).str.startswith(tuple(params.get("ventes_distributeur") or params["ventes"]))]
    df_r   = df_t[mask_retours(df_t, params)]
    df_rem = df_t[mask_remises(df_t, params)]
    # Variation de stock (compte 603/713 par défaut, configurable via params["stock"]) : isolée
    # des charges variables "directes" et de la marge brute. Mélangé aux autres comptes 6xx, ce
    # compte peut faire apparaître un titre comme très rentable alors que le coût de ses invendus
    # est seulement différé (stock en hausse = crédit sur ce compte) et non absent — cf. anomalie
    # constatée sur "Les Couleurs de l'Exil" (marge brute et résultat net gonflés par un crédit de
    # stock de 7 625 € pour seulement 5 436 € de ventes).
    prefixes_stock = tuple(params.get("stock") or ["603"])
    df_c     = df_t[df_t["Compte"].astype(str).str.startswith(tuple(params["charges"]))
                     & (~df_t["Compte"].astype(str).str.startswith(prefixes_stock))]
    df_stock = df_t[df_t["Compte"].astype(str).str.startswith(prefixes_stock)]
    # Reprises sur provisions (ex. 781/7810 = reprise de provision pour retour, cf. ⚙️
    # Paramétrage analytique → "Comptes reprises sur provisions") : nettées ici directement
    # contre les charges directes de ce titre, en contrepartie de la dotation initiale (ex.
    # 6810, nature "Provision pour retour" ci-dessous) — uniquement si ces écritures portent
    # le code analytique de ce titre précis ; sinon (provision gérée au niveau société), le
    # montant reste à 0 pour ce titre et n'apparaît que dans les totaux globaux.
    prefixes_prov_reprises_t = tuple(params.get("provisions_reprises") or [])
    df_prov_reprises_t = (df_t[df_t["Compte"].astype(str).str.startswith(prefixes_prov_reprises_t)]
                           if prefixes_prov_reprises_t else df_t.iloc[0:0])
    net_prov_reprises_t = df_prov_reprises_t["Crédit"].sum() - df_prov_reprises_t["Débit"].sum()

    # Charges fixes imputées = quote-part de la répartition des charges indirectes sur ce
    # titre précis (compte "CHARGES INDIRECTES REPARTIES", généré si la répartition au nombre
    # de titres actifs a été activée dans ⚙️ Paramétrage analytique). Vaut 0 sinon.
    df_cfi = df_t[df_t["Compte"].astype(str) == "CHARGES INDIRECTES REPARTIES"]

    ventes_ht     = df_v["Crédit"].sum()
    retours_m     = df_r["Débit"].sum() - df_r["Crédit"].sum()
    remises_m     = df_rem["Débit"].sum() - df_rem["Crédit"].sum()
    ca_net        = ventes_ht - retours_m - remises_m
    # Net débit-crédit, hors variation de stock (isolée ci-dessus) : ne sommer que le débit
    # surestimerait la charge réelle lorsque le stock augmente sur la période (crédit de sens
    # contraire) ; l'inclure dans les charges variables gonflerait ou dégonflerait à tort la
    # marge brute et le résultat net du titre.
    charges_v       = (df_c["Débit"].sum() - df_c["Crédit"].sum()) - net_prov_reprises_t
    variation_stock = df_stock["Débit"].sum() - df_stock["Crédit"].sum()
    marge_brute   = ca_net - charges_v
    charges_fixes = df_cfi["Débit"].sum()
    # La variation de stock reste isolée de la marge brute (cf. anomalie "Les Couleurs de
    # l'Exil" ci-dessus), mais est désormais incluse dans le résultat net du titre — à la
    # demande du client — comme une ligne de charge/produit à part entière entre marge brute
    # et charges fixes imputées, plutôt que purement informative.
    resultat_net  = marge_brute - variation_stock - charges_fixes
    ventes_distrib_t = df_v_distrib_t["Crédit"].sum()
    taux_ret      = (retours_m / ventes_distrib_t * 100) if ventes_distrib_t else 0
    taux_rem      = (remises_m / ventes_distrib_t * 100) if ventes_distrib_t else 0

    if resultat_net > 0 and taux_ret < 20:
        signal, bg, fg = "🟢 Titre rentable", "#d1fae5", "#065f46"
    elif resultat_net > 0 and taux_ret < 35:
        signal, bg, fg = "🟡 Rentabilité à surveiller", "#fef3c7", "#92400e"
    else:
        signal, bg, fg = "🔴 Titre en difficulté", "#fee2e2", "#991b1b"

    st.markdown(f"""
    <div style='padding:14px 18px; border-radius:12px; background:{bg}; color:{fg};
                font-weight:600; font-size:16px; margin-bottom:14px'>
        {label_affiche(isbn_sel, df)} — {signal}
    </div>
    """, unsafe_allow_html=True)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Taux de retour", f"{taux_ret:.1f} %")
    k2.metric("Taux de remise", f"{taux_rem:.1f} %")
    k3.metric("Marge brute", f"{fmt_fr(marge_brute, 0)} €")
    k4.metric("Résultat net (charges fixes incl.)", f"{fmt_fr(resultat_net, 0)} €")
    if charges_fixes == 0 and not st.session_state.get("repartition_active"):
        st.caption("ℹ️ Aucune charge fixe imputée : la répartition des charges indirectes sur les "
                   "titres actifs n'a pas été activée dans ⚙️ Paramétrage analytique.")
    if abs(variation_stock) > 0.5:
        if variation_stock < 0:
            st.caption(f"ℹ️ Stock de ce titre en hausse sur la période (crédit net de "
                       f"{fmt_fr(-variation_stock, 0)} € sur le compte de variation de stock) — montant "
                       "**non inclus dans la marge brute** (qui resterait sinon artificiellement gonflée/"
                       "dégonflée par un simple mouvement de stock), mais **inclus dans le résultat net** "
                       "ci-dessus, comme une ligne à part entière.")
        else:
            st.caption(f"ℹ️ Stock de ce titre en baisse sur la période (débit net de "
                       f"{fmt_fr(variation_stock, 0)} € sur le compte de variation de stock) — montant "
                       "**non inclus dans la marge brute**, mais **inclus dans le résultat net** ci-dessus, "
                       "comme une ligne à part entière.")

    st.markdown("#### Mini SIG — Soldes intermédiaires de gestion")
    # Détail des charges par nature (variation de stock, droits d'auteur, commercialisation,
    # structure/gérant, dotations, contenu, fabrication) si configuré dans ⚙️ Paramétrage
    # analytique (section "Détail des charges par nature") ; sinon, on garde la ligne agrégée
    # "Charges variables" comme auparavant.
    # detail_par_poste : associe à chaque poste du mini SIG les écritures du grand livre qui
    # le composent, pour permettre le drill-down juste en dessous du graphique (section
    # "🔍 Détail des écritures par poste").
    detail_par_poste = {"Ventes HT": df_v, "Retours": df_r, "Remises": df_rem}

    detail_charges = params.get("detail_charges")
    # Compte(s) mixte(s) contenu/fabrication (cf. ⚙️ Paramétrage analytique) : un même compte
    # (ex. 604 "Achats d'études et prestations de services") peut regrouper indifféremment des
    # prestations de contenu et de fabrication. On scinde alors ses écritures par mot-clé
    # détecté dans le libellé/fournisseur (ex. "REPROGRAPHIE", "PRINT") plutôt que par une
    # répartition aléatoire, pour rester reproductible et justifiable.
    mixte = params.get("contenu_fabrication_mixte") or {}
    mixte_comptes = tuple(mixte.get("comptes") or [])
    mots_cles_fab = [m.strip().upper() for m in (mixte.get("mots_cles_fabrication") or []) if m.strip()]
    if detail_charges:
        charge_rows = []
        total_detail = 0.0
        comptes_detailles = []
        for nom, prefixes in detail_charges.items():
            # Le compte de variation de stock (603/713, cf. df_stock ci-dessus) est déjà isolé
            # et affiché à part (message d'information sous les indicateurs) : on ignore ici
            # toute nature qui le couvrirait, pour ne pas le compter deux fois et ne pas fausser
            # le calcul de "reste" ci-dessous (qui se réconcilie avec charges_v, lequel exclut
            # déjà ce compte).
            if prefixes and any(str(p) == str(sp) or str(p).startswith(str(sp)) or str(sp).startswith(str(p))
                                 for p in prefixes for sp in prefixes_stock):
                continue
            # La nature "Provision pour retour" inclut automatiquement le(s) compte(s) de
            # reprise sur provisions configuré(s) (params["provisions_reprises"], ex. 781/7810),
            # pour imputer la reprise à la dotation initiale sur cette même ligne, comme demandé.
            prefixes_effectifs = (list(prefixes) + list(params.get("provisions_reprises") or [])
                                   if nom == "Provision pour retour" else prefixes)
            if prefixes_effectifs:
                df_nat_prefixe = df_t[df_t["Compte"].astype(str).str.startswith(tuple(prefixes_effectifs))]
            else:
                df_nat_prefixe = df_t.iloc[0:0]
            df_nat_mixte = df_t.iloc[0:0]
            if nom in ("Contenu", "Fabrication") and mixte_comptes:
                df_mixte_t = df_t[df_t["Compte"].astype(str).str.startswith(mixte_comptes)]
                if not df_mixte_t.empty and mots_cles_fab:
                    lib_upper = df_mixte_t["Libellé"].astype(str).str.upper()
                    mask_fab = lib_upper.str.contains("|".join(mots_cles_fab), regex=True)
                else:
                    mask_fab = pd.Series(False, index=df_mixte_t.index)
                df_nat_mixte = df_mixte_t[mask_fab] if nom == "Fabrication" else df_mixte_t[~mask_fab]
                comptes_detailles.extend(mixte_comptes)
            if prefixes_effectifs or (nom in ("Contenu", "Fabrication") and mixte_comptes):
                df_nat = pd.concat([df_nat_prefixe, df_nat_mixte]) if not df_nat_mixte.empty else df_nat_prefixe
                val = df_nat["Débit"].sum() - df_nat["Crédit"].sum()
                comptes_detailles.extend(prefixes_effectifs)
            else:
                df_nat = df_t.iloc[0:0]
                val = 0.0
            charge_rows.append((f"− {nom}", -val, "deduction"))
            detail_par_poste[f"− {nom}"] = df_nat
            total_detail += val
        # Écart entre la somme des natures détaillées et le total réel des charges variables
        # (comptes non couverts par les 7 natures ci-dessus) : affiché explicitement pour ne
        # jamais rompre la réconciliation avec la marge brute.
        reste = charges_v - total_detail
        if abs(reste) > 0.5:
            label_reste = "− Autres charges directes (non détaillées)"
            charge_rows.append((label_reste, -reste, "deduction"))
            df_reste = (df_c[~df_c["Compte"].astype(str).str.startswith(tuple(comptes_detailles))]
                        if comptes_detailles else df_c)
            detail_par_poste[label_reste] = df_reste
    else:
        charge_rows = [("− Charges variables", -charges_v, "deduction")]
        detail_par_poste["− Charges variables"] = df_c

    if abs(variation_stock) > 0.5:
        detail_par_poste["− Variation de stock"] = df_stock
    if charges_fixes:
        detail_par_poste["− Charges fixes imputées"] = df_cfi

    rows_sig = (
        [
            ("Ventes HT", ventes_ht, "base"),
            ("− Retours", -retours_m, "deduction"),
            ("− Remises", -remises_m, "deduction"),
            ("= CA net",  ca_net,     "subtotal"),
        ]
        + charge_rows
        + [
            ("= Marge brute",              marge_brute,    "subtotal"),
        ]
        + ([("− Variation de stock", -variation_stock, "deduction")] if abs(variation_stock) > 0.5 else [])
        + [
            ("− Charges fixes imputées",   -charges_fixes, "deduction"),
            ("= Résultat net du titre",    resultat_net,   "total"),
        ]
    )
    html_rows = ""
    for libelle, montant, style in rows_sig:
        if style == "subtotal":
            row_style = "background:#eef2ff; font-weight:600;"
        elif style == "total":
            color = "#065f46" if montant >= 0 else "#991b1b"
            fill = "#d1fae5" if montant >= 0 else "#fee2e2"
            row_style = f"background:{fill}; font-weight:700; color:{color};"
        elif style == "deduction":
            row_style = "color:#b91c1c;"
        else:
            row_style = "font-weight:500;"
        html_rows += (f"<tr style='{row_style}'>"
                      f"<td style='padding:7px 12px; border-bottom:1px solid #eee'>{libelle}</td>"
                      f"<td style='padding:7px 12px; border-bottom:1px solid #eee; text-align:right'>{fmt_fr(montant, 0)} €</td>"
                      f"</tr>")
    st.markdown(f"""
    <table style='width:100%; border-collapse:collapse; font-size:14px; border-radius:8px; overflow:hidden'>
        {html_rows}
    </table>
    """, unsafe_allow_html=True)

    # Mesures du waterfall dérivées directement du style de chaque ligne de rows_sig (plutôt
    # que d'un décompte manuel, fragile dès que le nombre de lignes varie — détail des charges
    # par nature actif ou non, ligne "Variation de stock" présente ou non).
    _style_vers_mesure = {"base": "absolute", "deduction": "relative", "subtotal": "total", "total": "total"}
    measures = [_style_vers_mesure[r[2]] for r in rows_sig]
    fig_sig = go.Figure(go.Waterfall(
        orientation="v",
        measure=measures,
        x=[r[0] for r in rows_sig], y=[r[1] for r in rows_sig],
        text=[f"{fmt_fr(r[1], 0)} €" for r in rows_sig],
        textposition="outside",
        connector={"line": {"color": "gray", "width": 0.5}},
        decreasing={"marker": {"color": "#EF4444"}},
        increasing={"marker": {"color": "#10B981"}},
        totals={"marker": {"color": "#3B82F6"}}
    ))
    fig_sig.update_layout(height=340 if detail_charges else 300, margin=dict(t=10), yaxis_title="€",
                           separators=", ", xaxis_tickangle=-25 if detail_charges else 0)
    st.plotly_chart(fig_sig, use_container_width=True)

    # ── Détail des écritures par poste (drill-down) ──
    # Permet de retrouver, pour un poste du mini SIG choisi ci-dessus (Ventes, Retours,
    # chaque nature de charge...), les écritures analytiques individuelles qui le composent,
    # sans avoir à ressortir le grand livre complet.
    postes_dispo = [k for k, v in detail_par_poste.items() if not v.empty]
    if postes_dispo:
        st.markdown("#### 🔍 Détail des écritures par poste")
        poste_sel = st.selectbox(
            "Choisir un poste pour afficher les écritures analytiques correspondantes",
            postes_dispo, key=f"poste_sel_{isbn_sel}"
        )
        cols_detail = [c for c in ["Date", "Compte", "Libellé", "Débit", "Crédit"] if c in detail_par_poste[poste_sel].columns]
        df_detail_poste = detail_par_poste[poste_sel][cols_detail].sort_values("Date") if "Date" in cols_detail else detail_par_poste[poste_sel][cols_detail]
        formats_detail = {c: (lambda x: f"{fmt_fr(x, 2)} €") for c in ["Débit", "Crédit"] if c in cols_detail}
        st.dataframe(df_detail_poste.style.format(formats_detail), use_container_width=True, hide_index=True)
        st.caption(f"{len(df_detail_poste)} écriture(s) — somme Débit − Crédit : "
                   f"{fmt_fr(detail_par_poste[poste_sel]['Débit'].sum() - detail_par_poste[poste_sel]['Crédit'].sum(), 2)} €")

    df_t_evol = df_t.copy()
    df_t_evol["Mois"] = pd.to_datetime(df_t_evol["Date"], errors="coerce").dt.to_period("M").astype(str)
    evol_v = df_t_evol[df_t_evol["Compte"].astype(str).str.startswith(tuple(params["ventes"]))].groupby("Mois")["Crédit"].sum().reset_index()
    evol_r = df_t_evol[df_t_evol["Compte"].astype(str).str.startswith(tuple(params["retours"]))].groupby("Mois")["Débit"].sum().reset_index()
    evol = evol_v.merge(evol_r, on="Mois", how="left").fillna(0)
    evol.columns = ["Mois", "Ventes", "Retours"]
    if not evol.empty:
        st.markdown("#### Évolution mensuelle")
        fig = px.line(evol, x="Mois", y=["Ventes", "Retours"], markers=True, height=260)
        fig.update_layout(margin=dict(t=10), legend_title="")
        st.plotly_chart(fig, use_container_width=True)

    cr_data = {
        "Poste": [r[0] for r in rows_sig],
        "Montant (€)": [r[1] for r in rows_sig]
    }
    df_cr = pd.DataFrame(cr_data)
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_cr.to_excel(writer, index=False, sheet_name="Compte_Resultat")
        evol.to_excel(writer, index=False, sheet_name="Evolution_mensuelle")
    buffer.seek(0)
    st.download_button("📥 Exporter la fiche titre (Excel)", buffer,
                        file_name=f"Fiche_titre_{isbn_sel.replace('/', '-')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"export_fiche_{isbn_sel}")

def resoudre_mapping_auteurs(df):
    """Si une famille analytique nommée "AUTEUR" a été mappée dans ⚙️ Paramétrage analytique
    (en plus d'EDITION/COMMUNICATION/Types de dépenses), on en déduit directement la
    correspondance ISBN → auteur depuis la comptabilité elle-même, sans référentiel séparé.
    Retourne un dict {isbn: auteur} (vide si aucune famille de ce type n'est configurée)."""
    noms = st.session_state.get("noms_familles_actives", [])
    codes_cols = st.session_state.get("codes_cols", [])
    if not noms or not codes_cols or df is None or "Code_Analytique" not in df.columns:
        return {}
    idx_auteur = next((i for i, n in enumerate(noms) if "auteur" in n.lower()), None)
    if idx_auteur is None or idx_auteur >= len(codes_cols):
        return {}
    col_auteur = codes_cols[idx_auteur]
    if col_auteur not in df.columns or col_auteur == "Code_Analytique":
        return {}
    # .notna() est indispensable AVANT la conversion en str : un NaN authentique devient la
    # chaîne "nan" une fois converti, qui n'est pas vide et passerait donc à tort le filtre
    # suivant, laissant fuiter une valeur NaN (float) dans le mapping — mélangée à des noms
    # d'auteurs (str) ailleurs, elle casse tout tri/set ultérieur (comparaison float/str).
    sous = df[(df["Code_Analytique"].astype(str).str.strip() != "")
              & df[col_auteur].notna()
              & (df[col_auteur].astype(str).str.strip() != "")]
    if sous.empty:
        return {}
    # Auteur le plus fréquent par ISBN, au cas où une incohérence ponctuelle existerait.
    mapping = sous.groupby("Code_Analytique")[col_auteur].agg(lambda s: s.mode().iloc[0] if not s.mode().empty else s.iloc[0])
    return mapping.to_dict()

def obtenir_statut_fiscal(auteur):
    """Statut fiscal d'un auteur (BNC ou option pour les traitements et salaires) — un choix
    contractuel de l'auteur, non déductible de la seule comptabilité analytique, qui
    conditionne si le diffuseur doit précompter les cotisations sociales (CSG/CRDS, formation
    professionnelle, RAAP) en plus de la contribution diffuseur (due dans tous les cas).
    Cherche d'abord dans le référentiel manuel (📋 Référentiel), puis dans les statuts
    assignés aux auteurs détectés automatiquement (famille analytique AUTEUR) ; par défaut
    Traitements et salaires (préréglage demandé pour ce client, où la comptabilité montre un
    précompte sur la quasi-totalité des auteurs)."""
    for c in st.session_state.get("royalties_referentiel", []):
        if c.get("auteur") == auteur:
            return c.get("statut_fiscal", "Traitements et salaires (option assimilé)")
    return st.session_state.get("statut_fiscal_auto", {}).get(auteur, "Traitements et salaires (option assimilé)")

def calculer_indicateurs_titres(df, params, titres):
    """Calcule, pour chaque titre actif, CA brut/net, charges variables, charges fixes
    imputées (quote-part de répartition des charges indirectes sur ce titre, si activée)
    et résultat net — utilisé pour les repères rapides (titres significatifs / rentables /
    compliqués) de la page Analyse par titre."""
    df_i = df[df["Code_Analytique"].isin(titres)]

    def par_compte(prefix_list, col, exclude_prefix_list=None):
        if not prefix_list:
            return pd.Series(0.0, index=titres)
        mask = df_i["Compte"].astype(str).str.startswith(tuple(prefix_list))
        if exclude_prefix_list:
            mask = mask & (~df_i["Compte"].astype(str).str.startswith(tuple(exclude_prefix_list)))
        return df_i[mask].groupby("Code_Analytique")[col].sum().reindex(titres, fill_value=0.0)

    ventes  = par_compte(params["ventes"], "Crédit")
    # CA distributeur (compte configurable, ex. BLDD 7011 pour ce cas d'étude) : base STRICTE
    # du taux de retour — distinct de "Ventes HT" ci-dessus (périmètre large 701) utilisé pour
    # la marge/le résultat.
    ventes_distrib = par_compte(params.get("ventes_distributeur") or params["ventes"], "Crédit")
    # Exclusion des comptes remises du filtre retours : évite un double comptage quand le
    # compte remises (ex. 7091) est un sous-compte du compte retours (ex. 709).
    retours = (par_compte(params["retours"], "Débit", exclude_prefix_list=params.get("remises"))
               - par_compte(params["retours"], "Crédit", exclude_prefix_list=params.get("remises")))
    remises = par_compte(params["remises"], "Débit") - par_compte(params["remises"], "Crédit")
    # Variation de stock (compte 603/713 par défaut, configurable via params["stock"]) : exclue
    # des charges variables et affichée à part (cf. afficher_fiche_titre). Mélangée aux comptes
    # 6xx "classiques", elle peut faire apparaître un titre comme très rentable alors que le coût
    # de ses invendus est seulement différé (stock en hausse = crédit sur ce compte), pas absent.
    prefixes_stock = params.get("stock") or ["603"]
    charges = (par_compte(params["charges"], "Débit", exclude_prefix_list=prefixes_stock)
               - par_compte(params["charges"], "Crédit", exclude_prefix_list=prefixes_stock))
    variation_stock = par_compte(prefixes_stock, "Débit") - par_compte(prefixes_stock, "Crédit")
    # Reprises sur provisions (ex. 781/7810) nettées contre les charges directes du titre — cf.
    # afficher_fiche_titre, même logique. Reste à 0 pour les titres où la reprise n'est pas
    # taguée par ISBN (provision gérée au niveau société).
    prefixes_prov_reprises = params.get("provisions_reprises") or []
    net_prov_reprises_titre = (par_compte(prefixes_prov_reprises, "Crédit")
                                - par_compte(prefixes_prov_reprises, "Débit"))
    charges = charges - net_prov_reprises_titre
    # Charges fixes imputées = quote-part de la répartition des charges indirectes sur ce
    # titre (compte "CHARGES INDIRECTES REPARTIES", cf. module Paramétrage analytique).
    # Reste à 0 si la répartition n'a pas été activée.
    mask_cfi = df_i["Compte"].astype(str) == "CHARGES INDIRECTES REPARTIES"
    charges_fixes = df_i[mask_cfi].groupby("Code_Analytique")["Débit"].sum().reindex(titres, fill_value=0.0)

    res = pd.DataFrame({"Code_Analytique": titres})
    res["Ventes HT"]  = ventes.values
    res["Retours"]    = retours.values
    res["Remises"]    = remises.values
    res["CA net"]      = res["Ventes HT"] - res["Retours"] - res["Remises"]
    res["Charges variables"] = charges.values
    res["Marge brute"] = res["CA net"] - res["Charges variables"]
    res["Variation de stock (incluse au résultat, hors marge)"] = variation_stock.values
    res["Charges fixes imputées"] = charges_fixes.values
    res["Résultat net"] = res["Marge brute"] - res["Variation de stock (incluse au résultat, hors marge)"] - res["Charges fixes imputées"]
    res["Taux retour (%)"] = np.where(ventes_distrib.values != 0, res["Retours"] / ventes_distrib.values * 100, 0)
    res["Taux remise (%)"] = np.where(ventes_distrib.values != 0, res["Remises"] / ventes_distrib.values * 100, 0)
    # CA distributeur du titre (dénominateur du taux de retour/remise ci-dessus) : conservé pour
    # repérer les cas où un taux extrême (ex. > 100 %) vient d'un dénominateur très faible plutôt
    # que d'un vrai problème de retours (ex. titre vendu très peu en direct au distributeur mais
    # dont des retours d'un exercice antérieur remontent sur la période).
    res["CA distributeur"] = ventes_distrib.values

    def _signal(row):
        if row["Résultat net"] > 0 and row["Taux retour (%)"] < 20:
            return "🟢"
        elif row["Résultat net"] > 0 and row["Taux retour (%)"] < 35:
            return "🟡"
        return "🔴"
    res["Signal"] = res.apply(_signal, axis=1)
    return res

def build_data_summary():
    """Construit un résumé textuel du pivot pour l'envoyer à Claude."""
    if "df_pivot" not in st.session_state:
        return None
    df = st.session_state["df_pivot"].copy()
    params = st.session_state.get("param_comptes", {})
    ventes = params.get("ventes", ["701"])
    retours = params.get("retours", ["709"])
    df_v = df[df["Compte"].astype(str).str.startswith(tuple(ventes))]
    df_r = df[df["Compte"].astype(str).str.startswith(tuple(retours))]
    ca_brut = df_v["Crédit"].sum()
    total_retours = df_r["Débit"].sum()
    ca_net = ca_brut - total_retours
    taux_retour = (total_retours / ca_brut * 100) if ca_brut else 0
    df_isbn = filtrer_isbn_reels(df)
    top = df_isbn.groupby("Code_Analytique", as_index=False).agg({"Crédit": "sum", "Débit": "sum"})
    top["Résultat"] = top["Crédit"] - top["Débit"]
    top5 = top.nlargest(5, "Résultat")[["Code_Analytique", "Résultat"]].to_dict(orient="records")
    bot5 = top.nsmallest(5, "Résultat")[["Code_Analytique", "Résultat"]].to_dict(orient="records")
    summary = f"""
DONNÉES ANALYTIQUES — MAISON D'ÉDITION INDÉPENDANTE
====================================================
CA brut : {fmt_fr(ca_brut, 0)} €
Total retours : {fmt_fr(total_retours, 0)} €
CA net : {fmt_fr(ca_net, 0)} €
Taux de retour : {taux_retour:.1f} %
Nombre de titres (ISBN) : {top['Code_Analytique'].nunique()}

Top 5 titres les plus rentables :
{json.dumps(top5, ensure_ascii=False, indent=2)}

Top 5 titres les moins rentables :
{json.dumps(bot5, ensure_ascii=False, indent=2)}
"""
    return summary

SYSTEM_PROMPT = """Tu es un expert-comptable spécialisé dans l'accompagnement des maisons d'édition indépendantes françaises.

Tu connais parfaitement :
- Les normes comptables françaises (PCG) appliquées à l'édition
- Les spécificités du secteur : droits d'auteurs, retours éditeurs, provisions, distributeurs
- Les ratios clés : taux de retour, taux de remise, marge brute éditeur, couverture des droits
- Les logiciels comptables utilisés en édition (CEGID, Sage, EBP)

Tu assistes l'utilisateur dans l'analyse de ses données comptables analytiques via cet outil Streamlit dédié aux éditeurs indépendants.

L'outil comporte les modules suivants : Import des données, Paramétrage analytique, Tableau de bord éditorial, Analyse par titre, Trésorerie prévisionnelle, Droits d'auteurs, Retours & Remises, Synthèse financière.

Réponds toujours en français, de façon concise et orientée action. Si tu détectes une anomalie dans les données, signale-la clairement."""

def generate_qr():
    qr = qrcode.QRCode(version=1, box_size=6, border=3)
    qr.add_data(APP_URL)
    qr.make(fit=True)
    img = qr.make_image(fill_color="#1a1a2e", back_color="white")
    buf = BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

def load_demo_data():
    """Génère un jeu de données de démonstration réaliste pour une maison d'édition."""
    np.random.seed(42)
    isbns = [f"978-2-{i:04d}-{j:04d}-{k}" for i, j, k in [
        (1234, 1001, 1), (1234, 1002, 8), (1234, 1003, 5), (1234, 1004, 2),
        (1234, 1005, 9), (1234, 1006, 6), (1234, 1007, 3), (1234, 1008, 0)
    ]]
    titres = ["Le Dernier Manuscrit", "Mémoires du Vent", "Sous les Toits de Paris",
              "L'Héritage Silencieux", "Chroniques du Nord", "La Lumière d'Août",
              "Terres Inconnues", "Les Mots du Soir"]
    familles = ["Roman", "Roman", "Essai", "Roman", "Polar", "Essai", "Jeunesse", "Poésie"]
    rows = []
    dates = pd.date_range("2024-01-01", "2024-12-31", freq="MS")
    for date in dates:
        for isbn, titre, famille in zip(isbns, titres, familles):
            ventes = np.random.randint(500, 8000)
            rows.append({"Compte": "701100", "Débit": 0, "Crédit": round(ventes, 2),
                         "Code_Analytique": isbn, "Famille_Analytique": famille,
                         "Libellé": f"Ventes {titre}", "Date": date})
            if np.random.random() < 0.6:
                retours = round(ventes * np.random.uniform(0.05, 0.30), 2)
                rows.append({"Compte": "709100", "Débit": retours, "Crédit": 0,
                             "Code_Analytique": isbn, "Famille_Analytique": famille,
                             "Libellé": f"Retours {titre}", "Date": date})
            charges = round(ventes * np.random.uniform(0.15, 0.40), 2)
            rows.append({"Compte": "607100", "Débit": charges, "Crédit": 0,
                         "Code_Analytique": isbn, "Famille_Analytique": famille,
                         "Libellé": f"Charges {titre}", "Date": date})
        # Charges fixes mensuelles
        for compte, libelle, montant in [
            ("615000", "Loyer", 1200), ("641000", "Salaires", 4500),
            ("512000", "Banque", 0), ("512000", "Recettes banque", 0)
        ]:
            if compte == "512000":
                rows.append({"Compte": compte, "Débit": 0, "Crédit": np.random.randint(8000, 15000),
                             "Code_Analytique": "", "Famille_Analytique": "",
                             "Libellé": "Recettes banque", "Date": date})
            else:
                rows.append({"Compte": compte, "Débit": montant, "Crédit": 0,
                             "Code_Analytique": "", "Famille_Analytique": "",
                             "Libellé": libelle, "Date": date})
    df = pd.DataFrame(rows)
    df["Date"] = pd.to_datetime(df["Date"])
    return df
# =====================
# ACCUEIL
# =====================
if page == "🏠 Accueil":
    col_main, col_qr = st.columns([3, 1])
    with col_main:
        st.title("📚 Outil d'accompagnement éditorial")
        st.markdown(f"*Bienvenue, **{st.session_state['name']}** — Maisons d'édition indépendantes*")
        st.divider()

        # Stepper visuel
        st.markdown("### Comment démarrer ?")
        col1, col2, col3, col4 = st.columns(4)
        for col, num, icon, titre, desc in [
            (col1, "1", "📂", "Importer", "Chargez votre export comptable Excel"),
            (col2, "2", "⚙️", "Paramétrer", "Mappez vos colonnes et comptes"),
            (col3, "3", "📈", "Analyser", "Explorez vos tableaux de bord"),
            (col4, "4", "🤖", "Interroger", "Posez vos questions à l'IA"),
        ]:
            with col:
                st.markdown(f"""
                <div style='text-align:center; padding:16px; background:#f8f9fa; border-radius:12px; border:1px solid #e0e0e0'>
                    <div style='font-size:28px'>{icon}</div>
                    <div style='font-size:11px; color:#888; margin:4px 0'>Étape {num}</div>
                    <div style='font-weight:600; font-size:14px'>{titre}</div>
                    <div style='font-size:12px; color:#666; margin-top:4px'>{desc}</div>
                </div>
                """, unsafe_allow_html=True)

        st.divider()
        st.markdown("### Fonctionnalités disponibles")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("""
- 📂 **Import des données** — fichier Excel comptable analytique
- ⚙️ **Paramétrage analytique** — mapping colonnes, contrôles et répartition
- 📈 **Tableau de bord éditorial** — KPIs, tendances, répartitions
- 📖 **Analyse par titre** — compte de résultat par ISBN
            """)
        with c2:
            st.markdown("""
- 💰 **Trésorerie prévisionnelle** — projection multi-scénarios
- ✍️ **Droits d'auteurs** — calcul par paliers, suivi avances
- 📦 **Retours & Remises** — taux, tendances, alertes
- 🤖 **Assistant IA** — analyse et conseils en langage naturel
            """)
    with col_qr:
        st.markdown("### Accès mobile")
        st.image(generate_qr(), width=180)
        st.caption("Scannez pour ouvrir l'app sur votre téléphone")
        st.markdown(f"[🔗 Lien direct]({APP_URL})")
    st.stop()

# =====================
# IMPORT DES DONNÉES
# =====================
elif page == "📂 Import des données":
    st.header("📂 Import des données analytiques")
    tab1, tab2 = st.tabs(["📁 Importer mon fichier", "🎭 Données de démonstration"])

    with tab1:
        st.info("Importez votre export comptable analytique au format Excel (.xlsx).")
        fichier = st.file_uploader("Sélectionnez votre fichier Excel", type=["xlsx"])
        if fichier:
            try:
                df = pd.read_excel(fichier, header=0)
                df.columns = df.columns.str.strip()
                st.session_state["df_comptables"] = df
                st.success(f"✅ Fichier chargé — {df.shape[0]} lignes, {df.shape[1]} colonnes")
                st.write("**Colonnes détectées :**", list(df.columns))
                # Validation rapide — on ne signale que les colonnes structurelles évidentes
                # (numéro de compte, débit, crédit, date). Les colonnes analytiques par famille
                # (Famille/Catégorie/Code analytique : EDITION, COMMUNICATION, Types de dépenses...)
                # sont normalement vides sur la majorité des lignes : chaque écriture n'appartient
                # qu'à UNE seule famille à la fois, donc les colonnes des 2 autres familles sont
                # vides pour cette ligne. Les remonter ici comme "valeurs manquantes" est un faux
                # positif systématique — le vrai contrôle de cohérence analytique se fait plus loin,
                # dans ⚙️ Paramétrage analytique, où le mapping multi-familles est connu.
                # Exclusions : colonnes analytiques par famille (cf. ci-dessus), et colonnes
                # d'enrichissement/traçabilité optionnelles (ex. "Compte d'origine (avant
                # ventilation...)" ajoutée lors d'une scission manuelle de compte comme 604/605)
                # qui contiennent "compte" mais ne sont, par construction, renseignées que sur
                # un sous-ensemble des lignes — les y inclure produirait un faux positif
                # systématique sans rapport avec un vrai problème d'import.
                mots_exclusion = ["analytique", "catégorie", "categorie", "origine", "nature",
                                   "ventil", "traçabilité", "tracabilite"]
                colonnes_structurelles = [c for c in df.columns if any(
                    mot in c.lower() for mot in ["compte", "débit", "debit", "crédit", "credit", "date"]
                ) and not any(mot in c.lower() for mot in mots_exclusion)]
                if colonnes_structurelles:
                    missing = df[colonnes_structurelles].isnull().sum()
                    missing = missing[missing > 0]
                    if not missing.empty:
                        st.warning(f"⚠️ Valeurs manquantes sur des colonnes structurelles : {missing.to_dict()}")
                st.caption("ℹ️ Les colonnes analytiques par famille (EDITION, COMMUNICATION, Types de dépenses...) "
                           "sont normalement vides sur une bonne partie des lignes — chaque écriture n'appartient "
                           "qu'à une seule famille à la fois. Ce n'est pas signalé ici ; la cohérence analytique "
                           "réelle est vérifiée dans **⚙️ Paramétrage analytique**.")
                st.dataframe(df.head(10))
                st.info("➡️ Passez maintenant dans **⚙️ Paramétrage analytique** pour configurer les colonnes.")
            except Exception as e:
                st.error(f"❌ Erreur : {e}")
        else:
            st.warning("Aucun fichier sélectionné.")

    with tab2:
        st.info("Chargez un jeu de données fictif réaliste pour explorer tous les modules sans fichier réel.")
        if st.button("🎭 Charger les données de démonstration", type="primary"):
            df_demo = load_demo_data()
            st.session_state["df_comptables"] = df_demo
            st.session_state["df_source_mappe"] = df_demo
            # Générer automatiquement le pivot avec paramètres par défaut
            pivot = df_demo.groupby(
                ["Compte", "Famille_Analytique", "Code_Analytique", "Date", "Libellé"],
                as_index=False
            ).agg({"Débit": "sum", "Crédit": "sum"})
            st.session_state["df_pivot"] = pivot
            st.session_state["df_pivot_brut"] = pivot.copy()
            st.session_state["param_comptes"] = {
                "ventes": ["701"], "ventes_distributeur": ["701"], "retours": ["709"],
                "remises": ["7091"], "charges": ["6"], "provisions_reprises": ["781"],
                "charges_imputees": "Oui"
            }
            # Clés nécessaires au module Paramétrage (contrôles + répartition) pour rester
            # cohérent même quand on saute directement le mapping via les données de démo.
            st.session_state["labels_indirect"] = {"charges": "CHARGES INDIRECTES", "produits": "PRODUITS INDIRECTS"}
            st.session_state["familles_cols"] = ["Famille_Analytique"]
            st.session_state["codes_cols"] = ["Code_Analytique"]
            st.session_state["noms_familles_actives"] = ["EDITION"]
            st.session_state["repartition_active"] = False
            st.success("✅ Données de démonstration chargées — pivot analytique généré automatiquement !")
            st.dataframe(df_demo.head(15))
            st.info("➡️ Tous les modules sont maintenant accessibles. Commencez par **📈 Tableau de bord éditorial** !")

# =====================
# PARAMÉTRAGE ANALYTIQUE
# =====================
elif page == "⚙️ Paramétrage analytique":
    st.header("⚙️ Paramétrage analytique")
    if "df_comptables" not in st.session_state:
        st.warning("⚠️ Importez d'abord vos données via **📂 Import des données**.")
        st.stop()

    df = st.session_state["df_comptables"].copy()
    st.info("Associez les colonnes de votre fichier aux champs attendus, puis configurez vos comptes comptables.")

    columns = list(df.columns)

    # ── Recharger une configuration déjà enregistrée ──
    with st.expander("📥 Recharger une configuration déjà enregistrée", expanded=False):
        st.caption(
            "Réimportez le fichier JSON obtenu via « 💾 Sauvegarder la configuration » lors d'un précédent "
            "passage, pour pré-remplir automatiquement tout le mapping ci-dessous (colonnes, comptes, détail "
            "des charges par nature, familles analytiques, libellés indirects) — sans tout re-saisir à chaque "
            "import. Fonctionne si votre nouvel export utilise les mêmes noms de colonnes que celui d'origine."
        )
        fichier_config = st.file_uploader("Fichier de configuration (JSON)", type=["json"], key="uploader_config")
        if fichier_config is not None and st.button("📤 Appliquer cette configuration", key="btn_appliquer_config"):
            try:
                cfg = json.load(fichier_config)
                cfg_mapping = cfg.get("mapping", {})
                cfg_inv = {v: k for k, v in cfg_mapping.items()}
                cfg_params = cfg.get("param_comptes", {})
                cfg_detail = cfg_params.get("detail_charges") or {}
                cfg_labels = cfg.get("labels_indirect", {})
                cfg_familles_cols = cfg.get("familles_cols", [])
                cfg_codes_cols = cfg.get("codes_cols", [])
                cfg_noms_familles = cfg.get("noms_familles_actives", [])

                if cfg_inv.get("Compte") in columns:
                    st.session_state["map_compte_col"] = cfg_inv["Compte"]
                if cfg_inv.get("Débit") in columns:
                    st.session_state["map_debit_col"] = cfg_inv["Débit"]
                if cfg_inv.get("Crédit") in columns:
                    st.session_state["map_credit_col"] = cfg_inv["Crédit"]
                if cfg_inv.get("Date") in columns:
                    st.session_state["map_date_col"] = cfg_inv["Date"]
                if cfg_inv.get("Libellé", "") in ([""] + columns):
                    st.session_state["map_libelle_col"] = cfg_inv.get("Libellé", "")
                if cfg_inv.get("Journal", "") in ([""] + columns):
                    st.session_state["map_journal_col"] = cfg_inv.get("Journal", "")

                st.session_state["map_ventes_comptes"] = ",".join(cfg_params.get("ventes", ["701"]))
                st.session_state["map_ventes_distributeur_comptes"] = ",".join(
                    cfg_params.get("ventes_distributeur", cfg_params.get("ventes_bldd", cfg_params.get("ventes", ["7011"])))
                )
                st.session_state["map_retours_comptes"] = ",".join(cfg_params.get("retours", ["709"]))
                st.session_state["map_remises_comptes"] = ",".join(cfg_params.get("remises", ["7091"]))
                st.session_state["map_charges_comptes"] = ",".join(cfg_params.get("charges", ["6"]))
                st.session_state["map_provisions_reprises_comptes"] = ",".join(cfg_params.get("provisions_reprises", ["781"]))
                st.session_state["map_charges_imputees"] = cfg_params.get("charges_imputees", "Oui")

                st.session_state["cpt_variation_stock"] = ",".join(cfg_detail.get("Variation de stock", []))
                st.session_state["cpt_droits_auteur_detail"] = ",".join(cfg_detail.get("Droits d'auteur", []))
                st.session_state["cpt_commercialisation"] = ",".join(cfg_detail.get("Commercialisation", []))
                st.session_state["cpt_structure"] = ",".join(cfg_detail.get("Structure/gérant", []))
                st.session_state["cpt_dotations"] = ",".join(cfg_detail.get("Provision pour retour", cfg_detail.get("Dotations amort.", [])))
                st.session_state["cpt_contenu"] = ",".join(cfg_detail.get("Contenu", []))
                st.session_state["cpt_fabrication"] = ",".join(cfg_detail.get("Fabrication", []))
                cfg_mixte = cfg_params.get("contenu_fabrication_mixte") or {}
                st.session_state["cpt_mixte_contenu_fab"] = ",".join(cfg_mixte.get("comptes", []))
                st.session_state["cpt_mots_cles_fab"] = ",".join(cfg_mixte.get(
                    "mots_cles_fabrication",
                    ["REPROGRAPHIE", "IMPRIM", "PRINT", "FAÇONNAGE", "FACONNAGE", "ROTATIVE", "REPRO"]
                ))

                nb_fam = len(cfg_noms_familles) if cfg_noms_familles else 1
                st.session_state["map_nb_familles"] = nb_fam
                for i in range(nb_fam):
                    st.session_state[f"nom_famille_{i}"] = cfg_noms_familles[i] if i < len(cfg_noms_familles) else ""
                    _fc = cfg_inv.get(cfg_familles_cols[i]) if i < len(cfg_familles_cols) else None
                    _cc = cfg_inv.get(cfg_codes_cols[i]) if i < len(cfg_codes_cols) else None
                    if _fc in ([""] + columns):
                        st.session_state[f"famille_col_{i}"] = _fc or ""
                    if _cc in ([""] + columns):
                        st.session_state[f"code_col_{i}"] = _cc or ""

                st.session_state["map_label_ci"] = cfg_labels.get("charges", "CHARGES INDIRECTES")
                st.session_state["map_label_pi"] = cfg_labels.get("produits", "PRODUITS INDIRECTS")

                st.success("✅ Configuration chargée — les champs ci-dessous sont pré-remplis.")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Fichier de configuration invalide : {e}")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Mapping des colonnes de base")
        compte_col  = st.selectbox("Colonne Compte", columns, key="map_compte_col")
        debit_col   = st.selectbox("Colonne Débit", columns, key="map_debit_col")
        credit_col  = st.selectbox("Colonne Crédit", columns, key="map_credit_col")
        date_col    = st.selectbox("Colonne Date", columns, key="map_date_col")
        libelle_col = st.selectbox("Libellé (optionnel)", [""] + columns, key="map_libelle_col")
        journal_col = st.selectbox(
            "Code journal (optionnel)", [""] + columns, key="map_journal_col",
            help="Recommandé pour le module Trésorerie : permet d'exclure les écritures de report "
                 "à nouveau (reprise des soldes d'ouverture, souvent codées « AN ») des flux de "
                 "trésorerie reconstitués, afin de ne pas les compter comme des mouvements de la période."
        )
    with col2:
        st.subheader("Comptes comptables")
        ventes_comptes  = st.text_input("Comptes ventes (CA large)", value="701", key="map_ventes_comptes",
                                         help="Tous les sous-comptes de ventes (ex. 701 = 7010 + 7011 + ...), "
                                              "utilisés pour le CA brut/net affiché.")
        ventes_distributeur_comptes = st.text_input(
            "Compte(s) ventes distributeur (base taux de retour/remise)", value="7011",
            key="map_ventes_distributeur_comptes",
            help="Sous-compte(s) strictement dédié(s) au(x) distributeur(s)/diffuseur(s) dont émane le relevé "
                 "de retours/remises (ex. 7011 pour le distributeur BLDD de ce cas d'étude — à adapter pour "
                 "tout autre distributeur, ou pour en cumuler plusieurs séparés par une virgule). Distinct du "
                 "CA large ci-dessus. Les taux de retour et de remise sont calculés uniquement sur ce "
                 "périmètre — le relevé du distributeur ne couvre que ce(s) canal/canaux, pas les autres "
                 "sous-comptes de ventes (ex. 7010...) ni les autres produits inclus dans le CA élargi (708, "
                 "commissions, subventions...)."
        )
        retours_comptes = st.text_input("Comptes retours", value="709", key="map_retours_comptes")
        remises_comptes = st.text_input("Comptes remises", value="7091", key="map_remises_comptes")
        charges_comptes = st.text_input("Comptes charges", value="6", key="map_charges_comptes")
        provisions_reprises_comptes = st.text_input(
            "Comptes reprises sur provisions (ex. retour)", value="781", key="map_provisions_reprises_comptes",
            help="Reprises sur provisions (ex. 7810 = reprise de provision pour retour), en contrepartie d'une "
                 "dotation antérieure (cf. « Provision pour retour » dans le détail des charges ci-dessous). "
                 "Ces comptes ne sont PAS comptés en CA/produits : ils sont nettés directement contre les "
                 "charges, pour imputer la reprise à la provision initiale plutôt que de gonfler le CA."
        )
        charges_imputees = st.radio("Charges déjà imputées par section ?", ["Oui", "Non"], key="map_charges_imputees")

    with st.expander("📐 Détail des charges par nature (optionnel — mini SIG détaillé par titre)"):
        st.caption(
            "Renseignez ici les comptes correspondant à chaque nature de charge directe, conformément "
            "à la décomposition retenue pour le pilotage par titre (variation de stock, droits d'auteur, "
            "commercialisation, structure/gérant, provision pour retour, contenu, fabrication). "
            "Si cette section reste vide, la fiche titre (module **📖 Analyse par titre**) continue "
            "d'afficher une seule ligne agrégée « Charges variables », comme aujourd'hui. Dès qu'au moins "
            "une nature est renseignée, la fiche titre affiche un compte de résultat en cascade détaillé "
            "par nature de charge plutôt que la ligne agrégée."
        )
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            cpt_variation_stock  = st.text_input("Variation de stock", value="", key="cpt_variation_stock",
                                                  help="Ex. 603")
            cpt_droits_auteur    = st.text_input("Droits d'auteur", value="", key="cpt_droits_auteur_detail",
                                                  help="Ex. 6043")
            cpt_commercialisation = st.text_input("Commercialisation", value="", key="cpt_commercialisation",
                                                   help="Ex. 6228,645106")
            cpt_structure        = st.text_input("Structure / gérant", value="", key="cpt_structure",
                                                  help="Ex. 6411,6451,6453")
        with col_d2:
            cpt_dotations = st.text_input("Provision pour retour (dotation)", value="", key="cpt_dotations",
                                           help="Ex. 6810 — dotation aux provisions pour retour. La reprise "
                                                "correspondante (ex. 781/7810) est configurée séparément "
                                                "ci-dessus (« Comptes reprises sur provisions ») et nettée "
                                                "automatiquement contre les charges plutôt que comptée en CA.")
            cpt_contenu   = st.text_input("Contenu (préparation éditoriale, prépresse)", value="", key="cpt_contenu",
                                           help="Comptes dédiés exclusivement au contenu (si votre plan comptable "
                                                "distingue déjà contenu et fabrication sur des comptes séparés). "
                                                "Laissez vide si vous utilisez le compte mixte ci-dessous.")
            cpt_fabrication = st.text_input("Fabrication (impression, façonnage)", value="", key="cpt_fabrication",
                                             help="Comptes dédiés exclusivement à la fabrication. Ex. 605. "
                                                  "Laissez vide si vous utilisez le compte mixte ci-dessous.")
            cpt_mixte_contenu_fab = st.text_input(
                "Compte mixte contenu/fabrication (scindé par mot-clé)", value="", key="cpt_mixte_contenu_fab",
                help="Ex. 604000000 — à utiliser quand un même compte regroupe indifféremment des prestations "
                     "de contenu (traduction, iconographie, droits d'image...) et de fabrication (impression, "
                     "façonnage...), comme c'est souvent le cas sur un compte générique \"Achats d'études et "
                     "prestations de services\". Chaque écriture de ce compte est classée automatiquement selon "
                     "les mots-clés détectés dans son libellé (fournisseur) ci-dessous : les lignes qui matchent "
                     "un mot-clé vont en Fabrication, toutes les autres vont en Contenu. Ne pas indiquer ce même "
                     "compte dans les deux champs ci-dessus, pour éviter un double comptage.")
            mots_cles_fabrication = st.text_input(
                "Mots-clés « fabrication » (dans le libellé/fournisseur)", value="REPROGRAPHIE,IMPRIM,PRINT,FAÇONNAGE,FACONNAGE,ROTATIVE,REPRO",
                key="cpt_mots_cles_fab",
                help="Liste de mots-clés (insensible à la casse), séparés par des virgules. Toute écriture du "
                     "compte mixte dont le libellé contient un de ces mots est classée en Fabrication ; le reste "
                     "est classé en Contenu par défaut (plutôt qu'une répartition aléatoire, pour rester "
                     "reproductible et justifiable). Ajustez la liste selon les noms de vos imprimeurs/façonniers.")

    st.subheader("Familles analytiques")
    st.caption("Votre export peut contenir plusieurs familles analytiques en parallèle "
               "(ex. EDITION pour les ISBN, COMMUNICATION pour la création graphique, "
               "la famille native « Types de dépenses / revenus » de votre logiciel, et une "
               "éventuelle famille AUTEUR). "
               "Mappez ici chaque paire de colonnes Famille / Valeur analytique. La 1ère famille "
               "mappée sert de référence pour le pivot ISBN (EDITION) ; les suivantes ne servent "
               "qu'aux contrôles de cohérence, pour qu'une ligne déjà affectée dans une autre "
               "famille (ex. COMMUNICATION) ne soit pas signalée à tort comme non affectée.")
    st.info("""
    ✍️ **Famille AUTEUR (recommandé) :** si votre logiciel comptable le permet, créez une 4e famille
    analytique nommée **AUTEUR**, taguée sur les mêmes lignes que la famille EDITION (chaque écriture
    de droits d'auteur porte alors à la fois son ISBN et son auteur). Une fois mappée ici, le module
    **✍️ Droits d'auteurs → 📒 Réel (comptabilisé)** détecte automatiquement la correspondance ISBN → auteur
    directement depuis la comptabilité, sans référentiel à ressaisir séparément.
    """)
    st.warning("""
    ⚠️ **Attention à ne pas confondre deux colonnes qui se ressemblent, pour chaque famille :**
    - **« Catégorie : \\<famille\\> »** → contient la vraie valeur affectée (ex. l'ISBN, "CHARGES INDIRECTES"…).
      **C'est celle-ci qu'il faut choisir.**
    - **« Code analytique : \\<famille\\> »** → identifiant technique interne de votre logiciel,
      généralement **vide sur toutes les lignes**. La sélectionner par erreur fait remonter énormément
      de faux positifs au contrôle de cohérence.
    """)

    nb_familles = st.number_input("Nombre de familles analytiques à mapper", min_value=1, max_value=4, value=1, step=1, key="map_nb_familles")
    familles_mapping = []
    noms_suggestion = ["EDITION", "COMMUNICATION", "Types de dépenses / revenus", "AUTEUR"]
    for i in range(int(nb_familles)):
        fc1, fc2, fc3 = st.columns([1, 1, 1])
        with fc1:
            nom_famille = st.text_input(
                f"Nom de la famille {i+1}",
                value=noms_suggestion[i] if i < len(noms_suggestion) else "",
                key=f"nom_famille_{i}"
            )
        with fc2:
            famille_col_i = st.selectbox(f"Colonne « Famille de catégories » {i+1} (optionnel)", [""] + columns, key=f"famille_col_{i}")
        with fc3:
            code_col_i = st.selectbox(
                f"Colonne « Catégorie » (valeur analytique) {i+1} (optionnel)", [""] + columns, key=f"code_col_{i}",
                help="Choisissez la colonne « Catégorie : <famille> », PAS « Code analytique : <famille> » (souvent vide)."
            )
        familles_mapping.append({"nom": nom_famille, "famille_col": famille_col_i, "code_col": code_col_i})

    st.subheader("Libellés des lignes indirectes")
    col_li1, col_li2 = st.columns(2)
    label_charges_indirectes = col_li1.text_input("Libellé des charges indirectes", value="CHARGES INDIRECTES", key="map_label_ci")
    label_produits_indirects = col_li2.text_input("Libellé des produits indirects", value="PRODUITS INDIRECTS", key="map_label_pi")

    # Aperçu avant validation
    st.subheader("Aperçu du mapping")
    apercu = df.head(5)[[compte_col, debit_col, credit_col, date_col]].copy()
    apercu.columns = ["Compte", "Débit", "Crédit", "Date"]
    st.dataframe(apercu)

    if st.button("⚙️ Générer le socle analytique", type="primary"):
        mapping = {compte_col: "Compte", debit_col: "Débit", credit_col: "Crédit", date_col: "Date"}
        if libelle_col:
            mapping[libelle_col] = "Libellé"
        if journal_col:
            mapping[journal_col] = "Journal"

        # Chaque famille mappée génère une paire de colonnes Famille_Analytique[_i] / Code_Analytique[_i].
        # La famille n°1 (sans suffixe) reste la référence utilisée par les modules ISBN (EDITION).
        familles_cols, codes_cols, noms_familles_actives = [], [], []
        for i, fam in enumerate(familles_mapping):
            suffix = "" if i == 0 else f"_{i+1}"
            col_famille_out = f"Famille_Analytique{suffix}"
            col_code_out = f"Code_Analytique{suffix}"
            if fam["famille_col"]:
                mapping[fam["famille_col"]] = col_famille_out
            if fam["code_col"]:
                mapping[fam["code_col"]] = col_code_out
            familles_cols.append(col_famille_out)
            codes_cols.append(col_code_out)
            noms_familles_actives.append(fam["nom"] or f"Famille {i+1}")

        df.rename(columns=mapping, inplace=True)
        for col in familles_cols + codes_cols + ["Libellé", "Journal"]:
            if col not in df.columns:
                df[col] = ""
            else:
                df[col] = df[col].fillna("")
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df["Débit"] = pd.to_numeric(df["Débit"], errors="coerce").fillna(0)
        df["Crédit"] = pd.to_numeric(df["Crédit"], errors="coerce").fillna(0)
        df["Compte"] = df["Compte"].astype(str).str.strip()

        st.session_state["df_source_mappe"] = df
        st.session_state["labels_indirect"] = {
            "charges": label_charges_indirectes.strip(),
            "produits": label_produits_indirects.strip(),
        }
        st.session_state["familles_cols"] = familles_cols
        st.session_state["codes_cols"] = codes_cols
        st.session_state["noms_familles_actives"] = noms_familles_actives

        # ================================================
        # CONTRÔLES 1 ET 2 (cf. mémoire — 3 contrôles ; le 3e est plus bas, hors du bouton)
        # ================================================
        st.subheader("🔎 Contrôles de cohérence")
        alerte_globale = False

        # --- Contrôle 1 : contrôle de l'import ---
        champs_attendus = ["Compte", "Débit", "Crédit", "Date"]
        champs_manquants = [c for c in champs_attendus if c not in df.columns]
        lignes_compte_vide = df["Compte"].isna().sum() + (df["Compte"] == "").sum()
        lignes_date_invalide = df["Date"].isna().sum()

        if champs_manquants or lignes_compte_vide > 0 or lignes_date_invalide > 0:
            alerte_globale = True
            st.error("❌ Contrôle de l'import : anomalies détectées")
            if champs_manquants:
                st.write(f"- Champs manquants dans l'export : {champs_manquants}")
            if lignes_compte_vide > 0:
                st.write(f"- {lignes_compte_vide} ligne(s) sans numéro de compte")
            if lignes_date_invalide > 0:
                st.write(f"- {lignes_date_invalide} ligne(s) avec une date invalide")
        else:
            st.success("✅ Contrôle de l'import : tous les champs attendus sont présents et renseignés.")

        # --- Contrôle 2 : cohérence charges générales / charges analytiques ---
        # Une ligne compte comme affectée dès qu'elle porte un code dans AU MOINS UNE des
        # familles mappées (EDITION, COMMUNICATION, Types de dépenses...).
        df_pl = df[df["Compte"].str.startswith(("6", "7"))].copy()
        has_any_code = pd.Series(False, index=df_pl.index)
        for c in codes_cols:
            has_any_code = has_any_code | (df_pl[c].astype(str).str.strip() != "")
        df_pl_non_code = df_pl[~has_any_code]
        total_general_pl = (df_pl["Crédit"] - df_pl["Débit"]).sum()
        total_analytique_pl = (df_pl[has_any_code]["Crédit"] - df_pl[has_any_code]["Débit"]).sum()
        ecart_pl = round(total_general_pl - total_analytique_pl, 2)

        if len(df_pl_non_code) > 0 or abs(ecart_pl) > 0.01:
            alerte_globale = True
            st.error("❌ Contrôle de cohérence charges/produits analytiques : anomalies détectées")
            st.write(f"- {len(df_pl_non_code)} ligne(s) de charge ou de produit sans code analytique dans "
                     f"aucune des familles mappées ({', '.join(noms_familles_actives)}) "
                     f"(montant net non affecté : {fmt_fr(round((df_pl_non_code['Crédit'] - df_pl_non_code['Débit']).sum(), 2), 2)} €)")
            st.write(f"- Écart total général / total analytique : {fmt_fr(ecart_pl, 2)} €")
            st.dataframe(df_pl_non_code[["Compte", "Libellé", "Date", "Débit", "Crédit"]].head(50))
        else:
            st.success(f"✅ Contrôle de cohérence charges/produits analytiques : toutes les lignes 6xx/7xx sont "
                       f"affectées (toutes familles confondues : {', '.join(noms_familles_actives)}), "
                       f"total général = total analytique.")

        if alerte_globale:
            st.warning("⚠️ Des anomalies ont été détectées ci-dessus (contrôles 1 et 2). Vous pouvez tout de "
                       "même générer le socle, mais il est recommandé de corriger l'affectation analytique "
                       "en amont avant de valider le rapport de pilotage.")
        else:
            st.success("✅ Contrôles 1 et 2 passés avec succès.")
        st.caption("Le contrôle 3 (cohérence des volumes BLDD) et la répartition des charges/produits "
                   "indirects sont disponibles juste en dessous, une fois le socle généré ci-dessous — "
                   "ils restent accessibles sans qu'il soit nécessaire de re-générer le socle.")

        # ================================================
        # GÉNÉRATION DU PIVOT
        # ================================================
        group_cols = ["Compte"] + familles_cols + codes_cols + ["Date"]
        if "Libellé" in df.columns:
            group_cols.append("Libellé")
        if "Journal" in df.columns:
            group_cols.append("Journal")
        pivot = df.groupby(group_cols, as_index=False).agg({"Débit": "sum", "Crédit": "sum"})
        # Fusionne les codes EAN dupliqués par un libellé légèrement différent (casse,
        # troncature...) avant de figer le socle, pour que chaque titre ne remonte jamais
        # sous deux codes analytiques distincts dans les classements et fiches par titre.
        nb_avant = pivot["Code_Analytique"].astype(str).nunique() if "Code_Analytique" in pivot.columns else 0
        pivot = normaliser_codes_ean(pivot, "Code_Analytique")
        nb_apres = pivot["Code_Analytique"].astype(str).nunique() if "Code_Analytique" in pivot.columns else 0
        if nb_avant and nb_apres < nb_avant:
            st.info(f"ℹ️ {nb_avant - nb_apres} code(s) analytique(s) fusionné(s) automatiquement "
                    f"(même numéro EAN, libellé légèrement différent d'une ligne à l'autre).")

        st.session_state["df_pivot"] = pivot
        st.session_state["df_pivot_brut"] = pivot.copy()
        def _split_comptes(s):
            return [c.strip() for c in s.split(",") if c.strip()]

        detail_charges = {
            "Variation de stock":  _split_comptes(cpt_variation_stock),
            "Droits d'auteur":     _split_comptes(cpt_droits_auteur),
            "Commercialisation":   _split_comptes(cpt_commercialisation),
            "Structure/gérant":    _split_comptes(cpt_structure),
            "Provision pour retour": _split_comptes(cpt_dotations),
            "Contenu":             _split_comptes(cpt_contenu),
            "Fabrication":         _split_comptes(cpt_fabrication),
        }
        # Compte(s) mixte(s) où contenu et fabrication ne sont pas distingués par compte
        # distinct (ex. un unique compte 604 "Achats d'études et prestations de services"
        # utilisé aussi bien pour un imprimeur que pour un traducteur) : on scinde alors les
        # écritures de ce compte par mot-clé détecté dans le libellé/fournisseur, plutôt que
        # par une répartition aléatoire, pour rester reproductible et justifiable dans le
        # mémoire (cf. ⚙️ Paramétrage analytique → Détail des charges par nature).
        mixte_comptes = _split_comptes(cpt_mixte_contenu_fab)
        mots_cles_fab = _split_comptes(mots_cles_fabrication)
        contenu_fabrication_mixte = (
            {"comptes": mixte_comptes, "mots_cles_fabrication": mots_cles_fab}
            if mixte_comptes else None
        )
        # Ne conserver le détail que si au moins une nature (ou le compte mixte) a été
        # renseignée — sinon la fiche titre continue d'afficher la ligne agrégée "Charges
        # variables".
        detail_charges_actif = any(v for v in detail_charges.values()) or bool(mixte_comptes)

        st.session_state["param_comptes"] = {
            "ventes":  [c.strip() for c in ventes_comptes.split(",")],
            "ventes_distributeur": [c.strip() for c in ventes_distributeur_comptes.split(",")] if ventes_distributeur_comptes.strip() else [c.strip() for c in ventes_comptes.split(",")],
            "retours": [c.strip() for c in retours_comptes.split(",")],
            "remises": [c.strip() for c in remises_comptes.split(",")],
            "charges": [c.strip() for c in charges_comptes.split(",")],
            "provisions_reprises": [c.strip() for c in provisions_reprises_comptes.split(",") if c.strip()],
            "charges_imputees": charges_imputees,
            "detail_charges": detail_charges if detail_charges_actif else None,
            "contenu_fabrication_mixte": contenu_fabrication_mixte,
        }

        # Export configuration JSON
        config = {
            "mapping": mapping,
            "param_comptes": st.session_state["param_comptes"],
            "familles_cols": familles_cols,
            "codes_cols": codes_cols,
            "noms_familles_actives": noms_familles_actives,
            "labels_indirect": st.session_state["labels_indirect"],
        }
        st.success("✅ Socle analytique généré avec succès !")
        st.caption("💡 Conservez ce fichier de configuration : vous pourrez le réimporter la prochaine fois "
                   "via « 📥 Recharger une configuration déjà enregistrée » ci-dessus, pour éviter de "
                   "refaire tout le mapping.")
        st.download_button(
            "💾 Sauvegarder la configuration (JSON)",
            data=json.dumps(config, ensure_ascii=False, indent=2),
            file_name="config_edition.json", mime="application/json"
        )
        st.dataframe(pivot.head(15))

    # ================================================
    # CONTRÔLE 3 : cohérence des volumes BLDD
    # ================================================
    # Placé HORS du bloc "if st.button(...)" : dans Streamlit, tout widget interactif
    # (uploader, radio...) déclenche un nouveau passage du script où st.button() redevient
    # False. S'il était imbriqué dans le bouton, interagir avec lui faisait disparaître tout
    # le socle généré. Ce bloc lit donc les données depuis st.session_state et reste actif
    # tant que le socle a été généré au moins une fois.
    if "df_source_mappe" in st.session_state:
        st.subheader("🔎 Contrôle de cohérence des volumes BLDD")
        st.caption("Redéposez le relevé BLDD du mois pour vérifier que les écritures de CA, remises et commissions "
                   "générées dans l'export analytique correspondent bien au relevé du diffuseur.")
        fichier_bldd = st.file_uploader("Relevé BLDD (optionnel)", type=["xlsx", "csv"], key="controle_bldd")

        if fichier_bldd:
            try:
                df_source = st.session_state["df_source_mappe"]
                param_ctrl = st.session_state["param_comptes"]
                if fichier_bldd.name.endswith(".csv"):
                    df_bldd = pd.read_csv(fichier_bldd)
                else:
                    df_bldd = pd.read_excel(fichier_bldd)
                df_bldd.columns = df_bldd.columns.str.strip()
                total_bldd = 0
                if "Débit" in df_bldd.columns and "Crédit" in df_bldd.columns:
                    total_bldd = (df_bldd["Crédit"] - df_bldd["Débit"]).sum()
                comptes_ctrl = tuple(param_ctrl["ventes"] + param_ctrl["retours"] + param_ctrl["remises"])
                mask_ctrl = df_source["Compte"].astype(str).str.startswith(comptes_ctrl)
                total_export_bldd = (df_source[mask_ctrl]["Crédit"] - df_source[mask_ctrl]["Débit"]).sum()
                ecart_bldd = round(total_bldd - total_export_bldd, 2)
                if abs(ecart_bldd) > 0.01:
                    st.error(f"❌ Écart entre le relevé BLDD ({fmt_fr(total_bldd, 2)} €) et l'export analytique "
                             f"({fmt_fr(total_export_bldd, 2)} €) : {fmt_fr(ecart_bldd, 2)} €")
                else:
                    st.success(f"✅ Le relevé BLDD ({fmt_fr(total_bldd, 2)} €) correspond à l'export analytique.")
            except Exception as e:
                st.warning(f"Impossible de lire le relevé BLDD : {e}")
        else:
            st.info("Aucun relevé BLDD déposé — ce contrôle sera ignoré pour cette génération.")

    # ================================================
    # RÉPARTITION DES CHARGES ET PRODUITS INDIRECTS
    # ================================================
    # Également hors du bouton, pour la même raison : le radio ci-dessous doit rester
    # interactif sans faire disparaître le socle déjà généré.
    if "df_pivot_brut" in st.session_state:
        st.subheader("📐 Répartition des charges et produits indirects")
        st.markdown("""
        Les charges de structure et certains produits (subventions, prestations non identifiables)
        n'ont pas pu être affectés à un titre en particulier et ont été codés comme
        **charges/produits indirects**. Conformément à la méthodologie retenue, il est proposé
        de les répartir au **nombre de titres actifs** plutôt qu'au chiffre d'affaires, afin de ne pas
        masquer les titres non rentables derrière les titres porteurs.
        """)

        pivot_brut = st.session_state["df_pivot_brut"]
        label_ci = st.session_state["labels_indirect"]["charges"]
        label_pi = st.session_state["labels_indirect"]["produits"]

        masque_ci = pivot_brut["Code_Analytique"].astype(str).str.strip() == label_ci
        masque_pi = pivot_brut["Code_Analytique"].astype(str).str.strip() == label_pi
        total_charges_indirectes = (pivot_brut[masque_ci]["Débit"] - pivot_brut[masque_ci]["Crédit"]).sum()
        total_produits_indirects = (pivot_brut[masque_pi]["Crédit"] - pivot_brut[masque_pi]["Débit"]).sum()

        titres_actifs = sorted(filtrer_isbn_reels(pivot_brut)["Code_Analytique"].astype(str).unique().tolist())
        nb_titres_actifs = len(titres_actifs)

        col_r1, col_r2, col_r3 = st.columns(3)
        col_r1.metric("Charges indirectes détectées", f"{fmt_fr(total_charges_indirectes, 2)} €")
        col_r2.metric("Produits indirects détectés", f"{fmt_fr(total_produits_indirects, 2)} €")
        col_r3.metric("Nombre de titres actifs", nb_titres_actifs)

        repartir = st.radio(
            "Souhaitez-vous répartir les charges et produits indirects sur les titres actifs ?",
            ["Non, je garde une ligne 'indirecte' globale", "Oui, répartir sur les titres actifs"],
            index=0,
            key="repartir_radio"
        )

        st.selectbox(
            "Clé de répartition (inducteur)",
            ["Nombre de titres actifs"],
            help="Seul l'inducteur 'nombre de titres actifs' est disponible pour le moment, "
                 "conformément à la méthodologie retenue (le CA n'est volontairement pas proposé "
                 "car il compenserait les titres non rentables par les titres porteurs)."
        )

        if repartir.startswith("Oui"):
            if nb_titres_actifs == 0:
                st.error("❌ Aucun titre actif détecté — répartition impossible.")
            else:
                pivot_reparti = pivot_brut[~masque_ci & ~masque_pi].copy()
                nouvelles_lignes = []

                # On répartit le Débit et le Crédit bruts séparément (et non le seul solde
                # net Débit-Crédit) : certaines lignes indirectes portent à la fois du débit
                # et du crédit (ex. régularisations/subventions avec contrepartie), et ne
                # garder qu'une quote-part nette ferait disparaître la partie brute de ces
                # montants des indicateurs "bruts" (ex. CA brut du Tableau de bord, qui est
                # une somme de Crédit seul, pas un solde). Le résultat net global n'est pas
                # affecté (Débit et Crédit sont toujours répartis dans les mêmes proportions,
                # donc leur différence — le net réparti — reste inchangée), mais les totaux
                # bruts restent réconciliables avec le grand livre d'origine.
                total_ci_debit = pivot_brut[masque_ci]["Débit"].sum()
                total_ci_credit = pivot_brut[masque_ci]["Crédit"].sum()
                if total_ci_debit != 0 or total_ci_credit != 0:
                    part_charge_debit = round(total_ci_debit / nb_titres_actifs, 2)
                    part_charge_credit = round(total_ci_credit / nb_titres_actifs, 2)
                    for isbn in titres_actifs:
                        nouvelles_lignes.append({
                            "Compte": "CHARGES INDIRECTES REPARTIES",
                            "Famille_Analytique": "EDITION",
                            "Code_Analytique": isbn,
                            "Date": pd.NaT,
                            "Libellé": f"Quote-part charges indirectes ({nb_titres_actifs} titres actifs)",
                            "Débit": part_charge_debit,
                            "Crédit": part_charge_credit,
                        })

                total_pi_debit = pivot_brut[masque_pi]["Débit"].sum()
                total_pi_credit = pivot_brut[masque_pi]["Crédit"].sum()
                if total_pi_debit != 0 or total_pi_credit != 0:
                    part_produit_debit = round(total_pi_debit / nb_titres_actifs, 2)
                    part_produit_credit = round(total_pi_credit / nb_titres_actifs, 2)
                    for isbn in titres_actifs:
                        nouvelles_lignes.append({
                            "Compte": "PRODUITS INDIRECTS REPARTIS",
                            "Famille_Analytique": "EDITION",
                            "Code_Analytique": isbn,
                            "Date": pd.NaT,
                            "Libellé": f"Quote-part produits indirects ({nb_titres_actifs} titres actifs)",
                            "Débit": part_produit_debit,
                            "Crédit": part_produit_credit,
                        })

                if nouvelles_lignes:
                    df_nouvelles = pd.DataFrame(nouvelles_lignes)
                    for c in pivot_reparti.columns:
                        if c not in df_nouvelles.columns:
                            df_nouvelles[c] = None
                    pivot_reparti = pd.concat([pivot_reparti, df_nouvelles[pivot_reparti.columns]], ignore_index=True)

                st.session_state["df_pivot"] = pivot_reparti
                st.session_state["repartition_active"] = True
                st.session_state["repartition_detail"] = {
                    "nb_titres_actifs": nb_titres_actifs,
                    "part_charge": round(total_charges_indirectes / nb_titres_actifs, 2) if nb_titres_actifs else 0,
                    "part_produit": round(total_produits_indirects / nb_titres_actifs, 2) if nb_titres_actifs else 0,
                }

                st.success(
                    f"✅ Répartition effectuée sur {nb_titres_actifs} titres actifs : "
                    f"{fmt_fr(round(total_charges_indirectes / nb_titres_actifs, 2), 2)} € de charges et "
                    f"{fmt_fr(round(total_produits_indirects / nb_titres_actifs, 2), 2)} € de produits par titre."
                )
                st.dataframe(pivot_reparti[
                    pivot_reparti["Compte"].isin(["CHARGES INDIRECTES REPARTIES", "PRODUITS INDIRECTS REPARTIS"])
                ].head(20))
        else:
            st.session_state["df_pivot"] = pivot_brut
            st.session_state["repartition_active"] = False
            st.info("Les charges et produits indirects restent regroupés sur une ligne globale "
                    "(non répartie sur les titres). Vous pourrez revenir sur ce choix à tout moment ci-dessus, "
                    "sans avoir besoin de régénérer le socle.")

        # ================================================
        # CONTRÔLE DE COUVERTURE DES CHARGES/PRODUITS
        # ================================================
        # Contrôle permanent (rejouable à tout moment, sans regénérer le socle) qui répond à
        # la question « est-ce que TOUT est bien récupéré quelque part ? ». Chaque écriture des
        # comptes 6xx/7xx doit tomber dans l'une de ces trois catégories : affectée à un vrai
        # titre, répartie comme indirecte (cf. section répartition ci-dessus), ou orpheline —
        # auquel cas elle disparaît silencieusement des indicateurs (Tableau de bord, Synthèse
        # financière) sans qu'aucune erreur ne le signale ailleurs. Complète le Contrôle 2
        # (import) qui ne s'affiche qu'une fois, au moment de générer le socle : ici on peut
        # revérifier à tout moment, y compris après avoir changé le mapping ou réimporté un
        # fichier corrigé.
        st.markdown("---")
        st.subheader("🔎 Contrôle de couverture des charges/produits")
        st.caption(
            "Vérifie que chaque écriture des comptes configurés est bien récupérée quelque part : "
            "affectée à un titre réel, répartie comme indirecte, ou orpheline (non affectée nulle "
            "part, donc absente des indicateurs sans avertissement). Rejouable à tout moment."
        )

        df_cc = st.session_state["df_pivot"]
        params_cc = st.session_state.get("param_comptes", {}) or {}
        label_ci_cc = st.session_state.get("labels_indirect", {}).get("charges", "CHARGES INDIRECTES")
        label_pi_cc = st.session_state.get("labels_indirect", {}).get("produits", "PRODUITS INDIRECTS")
        # Familles analytiques mappées AUTRES que la famille de référence (EDITION) — cf.
        # ⚙️ Paramétrage analytique → Familles analytiques. Une ligne peut être légitimement
        # vide en Code_Analytique (EDITION) tout en étant taguée dans une autre famille (ex.
        # COMMUNICATION) : elle n'est alors pas "perdue", simplement hors du périmètre édition
        # (titres/ISBN) sur lequel portent la fiche titre et le Tableau de bord éditorial. On
        # la distingue d'une vraie ligne orpheline pour ne pas déclencher une fausse alerte.
        autres_codes_cols = [c for c in (st.session_state.get("codes_cols") or [])
                              if c != "Code_Analytique" and c in df_cc.columns]
        noms_familles_cc = st.session_state.get("noms_familles_actives", [])

        compte_cc = df_cc["Compte"].astype(str)
        code_cc = (df_cc["Code_Analytique"].astype(str).str.strip()
                   if "Code_Analytique" in df_cc.columns else pd.Series("", index=df_cc.index))
        code_upper_cc = code_cc.str.upper()
        a_autre_famille_cc = pd.Series(False, index=df_cc.index)
        for c in autres_codes_cols:
            a_autre_famille_cc = a_autre_famille_cc | (df_cc[c].astype(str).str.strip() != "")

        def _classer_couverture(compte, code_up, autre_famille):
            if compte in ("CHARGES INDIRECTES REPARTIES", "PRODUITS INDIRECTS REPARTIS"):
                return "Indirect (réparti sur les titres)"
            if code_up in (label_ci_cc.upper(), label_pi_cc.upper()):
                return "Indirect (non encore réparti)"
            if code_up in ("", "NAN"):
                if autre_famille:
                    return "Hors périmètre édition (autre famille analytique)"
                return "⚠️ Orphelin (non affecté, aucune famille)"
            return "Titre (affecté)"

        couverture_cc = [
            _classer_couverture(c, k, a) for c, k, a in zip(compte_cc, code_upper_cc, a_autre_famille_cc)
        ]
        df_cc2 = df_cc.copy()
        df_cc2["_couverture"] = couverture_cc

        prefixes_charges_cc = tuple(params_cc.get("charges") or ["6"])
        mask_charges_cc = (compte_cc.str.startswith(prefixes_charges_cc)
                            | (compte_cc == "CHARGES INDIRECTES REPARTIES"))
        mask_produits_cc = (compte_cc.str.startswith("7")
                             | (compte_cc == "PRODUITS INDIRECTS REPARTIS"))

        for nom_scope, mask_scope, sens in [("Charges (comptes " + ",".join(prefixes_charges_cc) + ")", mask_charges_cc, 1),
                                             ("Produits (comptes 7xx)", mask_produits_cc, -1)]:
            sous = df_cc2[mask_scope]
            if sous.empty:
                continue
            g = sous.groupby("_couverture").apply(
                lambda x: pd.Series({
                    "Montant net (€)": sens * (x["Débit"].sum() - x["Crédit"].sum()),
                    "Nb lignes": len(x),
                }), include_groups=False
            )
            st.markdown(f"**{nom_scope}**")
            st.dataframe(g.style.format({"Montant net (€)": "{:,.2f}"}))
            nb_orphelin = sous[sous["_couverture"] == "⚠️ Orphelin (non affecté, aucune famille)"]
            if not nb_orphelin.empty:
                montant_orphelin = sens * (nb_orphelin["Débit"].sum() - nb_orphelin["Crédit"].sum())
                st.error(
                    f"❌ {len(nb_orphelin)} ligne(s) orpheline(s) détectée(s) ({fmt_fr(round(montant_orphelin, 2), 2)} €) "
                    f"— ni affectées à un titre, ni taguées « {label_ci_cc if sens == 1 else label_pi_cc} », ni "
                    f"rattachées à une autre famille analytique mappée. "
                    f"Ces montants sont invisibles dans le Tableau de bord et la Synthèse financière."
                )
                with st.expander(f"Voir le détail des lignes orphelines ({nom_scope})"):
                    st.dataframe(nb_orphelin[["Compte", "Libellé", "Date", "Débit", "Crédit"]].head(200))
            else:
                st.success(f"✅ Aucune ligne orpheline sur {nom_scope.lower()}.")
            nb_hors_perimetre = sous[sous["_couverture"] == "Hors périmètre édition (autre famille analytique)"]
            if not nb_hors_perimetre.empty:
                montant_hp = sens * (nb_hors_perimetre["Débit"].sum() - nb_hors_perimetre["Crédit"].sum())
                noms_autres = [n for i, n in enumerate(noms_familles_cc) if i > 0] or ["une autre famille"]
                st.info(
                    f"ℹ️ {len(nb_hors_perimetre)} ligne(s) ({fmt_fr(round(montant_hp, 2), 2)} €) sans code EDITION "
                    f"mais taguée(s) dans {', '.join(noms_autres)} — pas une anomalie, mais ces montants sont "
                    f"comptés dans le CA/charges globaux du Tableau de bord tout en étant hors périmètre "
                    f"« titres/ISBN » (donc absents de la fiche titre et de calculer_indicateurs_titres)."
                )
                with st.expander(f"Voir le détail ({nom_scope})"):
                    st.dataframe(nb_hors_perimetre[["Compte", "Libellé", "Date", "Débit", "Crédit"]].head(200))

        # --- Couverture par nature (détail des charges), en agrégé sur TOUS les titres ---
        # Complète la vérification titre par titre (fiche titre) par une vue globale : permet
        # de repérer immédiatement un compte manquant dans le mapping des natures (ex. un
        # sous-compte de commission oublié) sans avoir à inspecter chaque titre un par un.
        detail_charges_cc = params_cc.get("detail_charges")
        if detail_charges_cc:
            st.markdown("**Couverture par nature (détail des charges par nature), tous titres confondus**")
            df_titres_cc = df_cc2[df_cc2["_couverture"] == "Titre (affecté)"]
            df_charges_titres_cc = df_titres_cc[df_titres_cc["Compte"].astype(str).str.startswith(prefixes_charges_cc)]
            prefixes_stock_cc = tuple(params_cc.get("stock") or ["603"])
            df_charges_titres_cc = df_charges_titres_cc[
                ~df_charges_titres_cc["Compte"].astype(str).str.startswith(prefixes_stock_cc)
            ]
            total_charges_titres_cc = (df_charges_titres_cc["Débit"].sum() - df_charges_titres_cc["Crédit"].sum())
            comptes_couverts_cc = []
            lignes_nature = []
            mixte_cc = params_cc.get("contenu_fabrication_mixte") or {}
            mixte_comptes_cc = tuple(mixte_cc.get("comptes") or [])
            for nom_nat, prefixes_nat in detail_charges_cc.items():
                prefixes_eff = (list(prefixes_nat) + list(params_cc.get("provisions_reprises") or [])
                                 if nom_nat == "Provision pour retour" else list(prefixes_nat))
                if nom_nat in ("Contenu", "Fabrication") and mixte_comptes_cc:
                    prefixes_eff = prefixes_eff + list(mixte_comptes_cc)
                if not prefixes_eff:
                    continue
                m = df_charges_titres_cc["Compte"].astype(str).str.startswith(tuple(prefixes_eff))
                montant_nat = df_charges_titres_cc[m]["Débit"].sum() - df_charges_titres_cc[m]["Crédit"].sum()
                lignes_nature.append({"Nature": nom_nat, "Montant net (€)": montant_nat, "Nb lignes": int(m.sum())})
                comptes_couverts_cc.extend(prefixes_eff)
            reste_cc = (total_charges_titres_cc - sum(l["Montant net (€)"] for l in lignes_nature))
            lignes_nature.append({
                "Nature": "Non détaillé (« Autres charges directes »)",
                "Montant net (€)": reste_cc,
                "Nb lignes": int((~df_charges_titres_cc["Compte"].astype(str).str.startswith(tuple(comptes_couverts_cc))).sum()) if comptes_couverts_cc else len(df_charges_titres_cc),
            })
            df_nat_cc = pd.DataFrame(lignes_nature)
            st.dataframe(df_nat_cc.style.format({"Montant net (€)": "{:,.2f}"}))
            if total_charges_titres_cc and abs(reste_cc) / abs(total_charges_titres_cc) > 0.05:
                st.warning(
                    f"⚠️ {fmt_fr(round(reste_cc, 2), 2)} € ({abs(reste_cc)/abs(total_charges_titres_cc)*100:.1f}% "
                    f"des charges affectées aux titres) ne correspondent à aucune nature configurée. "
                    f"Vérifiez si un compte devrait être ajouté à l'une des natures ci-dessus (cf. ⚙️ "
                    f"Paramétrage analytique → Détail des charges par nature)."
                )
# =====================
# TABLEAU DE BORD ÉDITORIAL
# =====================
elif page == "📈 Tableau de bord éditorial":
    st.header("📈 Tableau de bord éditorial")
    if "df_pivot" not in st.session_state:
        st.warning("⚠️ Générez d'abord le socle via **⚙️ Paramétrage analytique**.")
        st.stop()

    df = st.session_state["df_pivot"].copy()
    params = st.session_state["param_comptes"]
    if st.session_state.get("repartition_active"):
        st.caption("ℹ️ Les charges/produits indirects ont été répartis sur les titres actifs.")

    # Filtres temporels
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    years = sorted(df["Date"].dt.year.dropna().unique().tolist())
    col_f1, col_f2 = st.columns([1, 3])
    with col_f1:
        annee = st.selectbox("Année", ["Toutes"] + [str(y) for y in years])
    if annee != "Toutes":
        df = df[df["Date"].dt.year == int(annee)]
    # Important : on NE supprime PAS les lignes sans date ici avant le calcul des indicateurs
    # clés. Les quote-parts de charges/produits indirects répartis (comptes "CHARGES INDIRECTES
    # REPARTIES"/"PRODUITS INDIRECTS REPARTIS", cf. ⚙️ Paramétrage analytique) sont enregistrées
    # avec une date vide (répartition globale sur la période, non rattachable à un mois précis).
    # Les supprimer ici les faisait disparaître des totaux du Tableau de bord (Charges totales,
    # Résultat net) alors qu'elles restaient prises en compte par la Synthèse financière et
    # l'Analyse par titre — d'où l'écart entre pages. df_date_ok (sous-ensemble avec date valide)
    # n'est utilisé que pour le graphique d'évolution mensuelle ci-dessous, qui a besoin d'un mois.
    df_date_ok = df.dropna(subset=["Date"]).copy()
    df_date_ok["Mois"] = df_date_ok["Date"].dt.to_period("M").astype(str)

    df_v = df[mask_ventes(df, params)]
    df_r = df[mask_retours(df, params)]
    df_rem = df[mask_remises(df, params)]
    df_c = df[mask_charges(df, params)]
    # CA distributeur (compte configurable, ex. 7011 = BLDD pour ce cas d'étude, adaptable à
    # tout autre distributeur) : base STRICTE de calcul du taux de retour/remise — distincte
    # du périmètre "ventes" large (params["ventes"], ex. 701 = 7010 + 7011 + ...) utilisé juste
    # en dessous pour les extournes, et du CA brut élargi (commissions/subventions/produits
    # divers). Le relevé du distributeur ne couvre que ce canal de vente précis : diviser par
    # un périmètre plus large (701 entier, ou le CA élargi) dilue artificiellement ces taux.
    prefixes_ca_distrib = tuple(params.get("ventes_distributeur") or params["ventes"])
    df_ca_distrib = df[df["Compte"].astype(str).str.startswith(prefixes_ca_distrib)]
    ca_brut_distrib = df_ca_distrib["Crédit"].sum()

    df_v_large = df[df["Compte"].astype(str).str.startswith(tuple(params["ventes"]))]
    # Extourne stricte sur le(s) compte(s) de ventes configuré(s) au sens large (ex. 701) :
    # facture annulée ou corrigée a posteriori — c'est le seul sens réel de « extourne sur ventes ».
    extournes_ventes = df_v_large["Débit"].sum()

    # CA brut/net inclut désormais tout produit (compte 7xx) qui n'est ni un retour ni une
    # remise — commissions, subventions, reprises, produits divers... (cf. mask_ventes) : ces
    # montants ne sont pas isolés dans une ligne séparée, ils sont directement dans le CA.
    ca_brut       = df_v["Crédit"].sum()
    # corrections_ventes = TOUS les débits du périmètre élargi "ventes" (mask_ventes), donc
    # extournes sur 701 + tout débit sur les autres comptes de produits (commissions,
    # subventions, produits divers, reprises...) repris dans ce même périmètre depuis que
    # celui-ci a été élargi (cf. mask_ventes). Il faut le déduire en totalité pour que le
    # résultat net se réconcilie avec le total comptable réel des comptes 6/7, mais on
    # distingue ci-dessous la part qui concerne réellement une extourne de vente du reste,
    # pour ne pas induire en erreur sur la nature de ce montant.
    corrections_ventes = df_v["Débit"].sum()
    autres_regul_produits = corrections_ventes - extournes_ventes
    total_retours = df_r["Débit"].sum() - df_r["Crédit"].sum()
    total_remises = df_rem["Débit"].sum() - df_rem["Crédit"].sum()
    ca_net        = ca_brut - total_retours - total_remises - corrections_ventes
    # Reprises sur provisions (ex. 781/7810 = reprise de provision pour retour) : exclues du CA
    # par mask_ventes ci-dessus (cf. mask_provisions_reprises), nettées ici directement contre
    # les charges — la reprise vient réduire la charge nette, en contrepartie de la dotation
    # initiale (ex. 6810), plutôt que de gonfler le chiffre d'affaires.
    df_prov_reprises = df[mask_provisions_reprises(df, params)]
    net_provisions_reprises = df_prov_reprises["Crédit"].sum() - df_prov_reprises["Débit"].sum()
    # Net débit-crédit : idem module Analyse par titre (comptes 603/713 à double sens).
    charges_tot   = (df_c["Débit"].sum() - df_c["Crédit"].sum()) - net_provisions_reprises
    resultat      = ca_net - charges_tot
    taux_retour   = (total_retours / ca_brut_distrib * 100) if ca_brut_distrib else 0
    taux_remise   = (total_remises / ca_brut_distrib * 100) if ca_brut_distrib else 0

    # KPIs
    st.subheader("Indicateurs clés")
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("CA brut", f"{fmt_fr(ca_brut, 0)} €")
    k2.metric("CA net", f"{fmt_fr(ca_net, 0)} €", delta=f"-{fmt_fr(total_retours+total_remises+corrections_ventes, 0)} €")
    if corrections_ventes:
        st.caption(
            f"ℹ️ Dont {fmt_fr(corrections_ventes, 0)} € de corrections déduites du CA net ci-dessus : "
            f"{fmt_fr(extournes_ventes, 0)} € d'extournes/corrections sur ventes proprement dites "
            f"(compte{'s' if len(params['ventes']) > 1 else ''} {', '.join(params['ventes'])}, factures "
            f"annulées ou corrigées a posteriori) et {fmt_fr(autres_regul_produits, 0)} € d'autres "
            "régularisations (débits) sur les autres comptes de produits inclus dans le CA élargi "
            "(commissions, subventions, produits divers, reprises...)."
        )
    k3.metric("Taux de retour", f"{taux_retour:.1f} %",
              delta_color="inverse", delta="⚠️ Élevé" if taux_retour > 25 else "✅ Normal")
    st.caption(f"ℹ️ Taux de retour et de remise calculés sur le CA distributeur (compte(s) "
               f"{', '.join(params.get('ventes_distributeur') or params['ventes'])} uniquement, hors "
               f"autres sous-comptes de ventes et hors commissions/subventions/produits divers) : "
               f"**{fmt_fr(ca_brut_distrib, 0)} €**. Taux de remise : **{taux_remise:.1f} %**.")
    k4.metric("Charges totales", f"{fmt_fr(charges_tot, 0)} €")
    if net_provisions_reprises:
        st.caption(f"ℹ️ Charges totales déjà nettes de {fmt_fr(net_provisions_reprises, 0)} € de reprises sur "
                   "provisions (ex. reprise de provision pour retour) — imputées à la dotation initiale plutôt "
                   "que comptées en CA.")
    k5.metric("Résultat net", f"{fmt_fr(resultat, 0)} €",
              delta_color="normal" if resultat >= 0 else "inverse")

    st.divider()
    col_g1, col_g2 = st.columns(2)

    # Évolution mensuelle CA
    with col_g1:
        st.subheader("Évolution mensuelle")
        # df_date_ok : sous-ensemble à date valide (cf. filtres temporels ci-dessus) — les
        # quote-parts indirectes réparties (sans date) sont exclues de ce graphique par mois,
        # mais restent dans les indicateurs clés (KPIs) calculés plus haut sur df_v/df_r/etc.
        df_v_date = df_date_ok[mask_ventes(df_date_ok, params)]
        df_r_date = df_date_ok[mask_retours(df_date_ok, params)]
        trend_v = df_v_date.groupby("Mois")["Crédit"].sum().reset_index().rename(columns={"Crédit": "CA brut"})
        trend_r = df_r_date.groupby("Mois")["Débit"].sum().reset_index().rename(columns={"Débit": "Retours"})
        trend = trend_v.merge(trend_r, on="Mois", how="left").fillna(0)
        trend["CA net"] = trend["CA brut"] - trend["Retours"]
        fig = go.Figure()
        fig.add_trace(go.Bar(x=trend["Mois"], y=trend["CA brut"], name="CA brut", marker_color="#3B82F6"))
        fig.add_trace(go.Bar(x=trend["Mois"], y=trend["Retours"], name="Retours", marker_color="#EF4444"))
        fig.add_trace(go.Scatter(x=trend["Mois"], y=trend["CA net"], name="CA net",
                                  mode="lines+markers", line=dict(color="#10B981", width=2)))
        fig.update_layout(barmode="overlay", legend=dict(orientation="h"), height=350,
                          xaxis_title="", yaxis_title="€", margin=dict(t=20))
        st.plotly_chart(fig, use_container_width=True)

    # Répartition par famille
    with col_g2:
        st.subheader("Répartition par famille")
        if df_v["Famille_Analytique"].str.strip().any():
            fam = df_v.groupby("Famille_Analytique")["Crédit"].sum().reset_index()
            fam = fam[fam["Famille_Analytique"].str.strip() != ""]
            if not fam.empty:
                fig2 = px.pie(fam, values="Crédit", names="Famille_Analytique",
                               hole=0.4, height=350)
                fig2.update_layout(margin=dict(t=20))
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.info("Aucune famille analytique renseignée.")
        else:
            st.info("Aucune famille analytique renseignée.")

    # Top 10 ISBN
    st.subheader("Top 10 titres par résultat net")
    # On ne classe que les vraies lignes ISBN (famille EDITION, hors code vide et hors
    # libellés indirects globaux) — sinon une ligne "CHARGES INDIRECTES" ou une ligne hors
    # comptes de charges/produits pouvait remonter en tête du classement.
    df_isbn_tb = filtrer_isbn_reels(df)
    top = df_isbn_tb.groupby("Code_Analytique", as_index=False).agg({"Crédit": "sum", "Débit": "sum"})
    top["Résultat"] = top["Crédit"] - top["Débit"]
    top10 = top.nlargest(10, "Résultat").copy()
    # Libellé affiché (anonymisé si le mode démonstration est actif) : distinct du
    # Code_Analytique réel utilisé pour tous les calculs, qui n'est jamais modifié.
    top10["Titre_affiche"] = top10["Code_Analytique"].apply(lambda c: label_affiche(c, df))
    fig3 = px.bar(top10, x="Titre_affiche", y="Résultat", text="Résultat",
                   color="Résultat", color_continuous_scale=["#EF4444", "#F59E0B", "#10B981"],
                   labels={"Titre_affiche": "ISBN / Titre", "Résultat": "Résultat net (€)"}, height=380)
    fig3.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
    fig3.update_layout(separators=", ")  # format FR : virgule decimale, espace milliers
    fig3.update_layout(showlegend=False, margin=dict(t=20))
    st.plotly_chart(fig3, use_container_width=True)

# =====================
# ANALYSE PAR TITRE
# =====================
elif page == "📖 Analyse par titre":
    st.header("📖 Analyse par titre (ISBN)")
    if "df_pivot" not in st.session_state:
        st.warning("⚠️ Générez d'abord le socle analytique.")
        st.stop()

    df = st.session_state["df_pivot"].copy()
    params = st.session_state["param_comptes"]
    if st.session_state.get("repartition_active"):
        st.caption("ℹ️ Les charges/produits indirects ont été répartis sur les titres actifs.")

    # Même filtre que le Tableau de bord : ne proposer que de vrais ISBN dans la liste.
    titres = sorted(filtrer_isbn_reels(df)["Code_Analytique"].astype(str).unique().tolist())
    if not titres:
        st.warning("Aucun ISBN/code analytique détecté dans les données.")
        st.stop()

    # ================================================
    # REPÈRES RAPIDES : significatifs / rentables / compliqués
    # ================================================
    st.subheader("🧭 Repères rapides")
    st.caption("Calculés sur l'ensemble des titres actifs, charges fixes imputées incluses "
               "si la répartition a été activée.")
    indicateurs = calculer_indicateurs_titres(df, params, titres)

    col_signif, col_top, col_bas = st.columns(3)

    def _liste_titres(container, sous_titre, df_tri, prefix_key, colonne_montant):
        container.markdown(f"**{sous_titre}**")
        if df_tri.empty:
            container.caption("Aucune donnée.")
            return
        for _, row in df_tri.iterrows():
            c1, c2 = container.columns([3, 1])
            c1.markdown(f"{row['Signal']} **{label_affiche(row['Code_Analytique'], df)}** — {fmt_fr(row[colonne_montant], 0)} €")
            if c2.button("Voir", key=f"{prefix_key}_{row['Code_Analytique']}", use_container_width=True):
                afficher_fiche_titre(row["Code_Analytique"], df, params)

    _liste_titres(col_signif, "📊 Les plus significatifs (CA brut)",
                  indicateurs.sort_values("Ventes HT", ascending=False).head(5),
                  "signif", "Ventes HT")
    _liste_titres(col_top, "🏆 Les plus rentables (résultat net)",
                  indicateurs.sort_values("Résultat net", ascending=False).head(5),
                  "rent", "Résultat net")
    _liste_titres(col_bas, "⚠️ Les plus compliqués (résultat net)",
                  indicateurs.sort_values("Résultat net", ascending=True).head(5),
                  "diff", "Résultat net")

    st.divider()

    # ================================================
    # INDICATEUR 1 — MARGE PAR TITRE (détail et alertes)
    # ================================================
    st.subheader("📐 Indicateur 1 — Marge par titre")
    st.caption("Marge brute = CA net − Charges directes · Marge nette = Marge brute − Quote-part charges "
               "indirectes. Seuils : 🔴 marge brute < 0 % · 🟠 entre 0 et 10 % · 🟢 cible > 15 %.")
    indic_marge = indicateurs.copy()
    indic_marge["Taux marge brute (%)"] = np.where(indic_marge["CA net"] != 0, indic_marge["Marge brute"] / indic_marge["CA net"] * 100, 0.0)
    indic_marge["Taux marge nette (%)"] = np.where(indic_marge["CA net"] != 0, indic_marge["Résultat net"] / indic_marge["CA net"] * 100, 0.0)
    marge_brute_totale = indic_marge["Marge brute"].sum()
    indic_marge["Contribution résultat global (%)"] = np.where(
        marge_brute_totale != 0, indic_marge["Marge brute"] / marge_brute_totale * 100, 0.0
    )
    droits_par_isbn_mrg = df[df["Compte"].astype(str) == "604300000"].groupby("Code_Analytique")["Débit"].sum()
    indic_marge["Droits d'auteurs période"] = indic_marge["Code_Analytique"].map(droits_par_isbn_mrg).fillna(0.0)
    indic_marge["Résultat analytique"] = indic_marge["Résultat net"] - indic_marge["Droits d'auteurs période"]

    def _alerte_marge_titre(t):
        if t < 0:   return "🔴 Déficitaire"
        if t < 10:  return "🟠 Fragile"
        if t >= 15: return "🟢 Cible atteinte"
        return "🟡 Intermédiaire"
    indic_marge["Alerte marge"] = indic_marge["Taux marge brute (%)"].apply(_alerte_marge_titre)

    mc1, mc2, mc3 = st.columns(3)
    mc1.metric("🔴 Titres déficitaires (marge brute < 0 %)", int((indic_marge["Taux marge brute (%)"] < 0).sum()))
    mc2.metric("🟠 Titres fragiles (0 – 10 %)", int(((indic_marge["Taux marge brute (%)"] >= 0) & (indic_marge["Taux marge brute (%)"] < 10)).sum()))
    mc3.metric("🟢 Titres cible (> 15 %)", int((indic_marge["Taux marge brute (%)"] >= 15).sum()))

    with st.expander("Voir le détail par titre (taux de marge, contribution, résultat analytique)"):
        cols_montant_mrg = ["CA net", "Marge brute", "Résultat net", "Droits d'auteurs période", "Résultat analytique"]
        cols_pct_mrg = ["Taux marge brute (%)", "Taux marge nette (%)", "Contribution résultat global (%)"]
        aff_mrg = indic_marge[["Code_Analytique"] + cols_montant_mrg + cols_pct_mrg + ["Alerte marge"]].sort_values("Marge brute", ascending=False)
        formats_mrg = {c: (lambda x: f"{fmt_fr(x, 0)} €") for c in cols_montant_mrg}
        formats_mrg.update({c: "{:.1f} %" for c in cols_pct_mrg})
        st.dataframe(aff_mrg.style.format(formats_mrg), use_container_width=True, hide_index=True)
        var_stock_mrg = df[df["Compte"].astype(str).str.startswith(("603", "713"))]
        variation_stock_mrg = (var_stock_mrg["Crédit"] - var_stock_mrg["Débit"]).sum()
        st.caption(f"ℹ️ Variation de stock globale (comptes 603/713, non ventilée par titre) : "
                   f"**{fmt_fr(variation_stock_mrg, 0)} €**.")

    st.divider()

    # ================================================
    # INDICATEUR — TAUX DE RETOUR PAR TITRE (repérage des gros retours)
    # ================================================
    st.subheader("🔁 Taux de retour par titre")
    st.caption("Taux de retour = Retours / CA distributeur (compte distributeur strict, cf. ⚙️ Paramétrage "
               "analytique). Triés du plus élevé au plus faible pour repérer rapidement les titres à "
               "gros retours.")
    indic_retour = indicateurs[["Code_Analytique", "Retours", "Remises", "CA distributeur",
                                 "Taux retour (%)", "Taux remise (%)", "Résultat net"]].copy()
    indic_retour["Titre"] = indic_retour["Code_Analytique"].apply(lambda c: label_affiche(c, df))

    def _alerte_taux_retour(t):
        if t >= 100: return "🔴 Extrême (≥ 100 %)"
        if t >= 35:  return "🔴 Élevé"
        if t >= 20:  return "🟠 À surveiller"
        return "🟢 Normal"
    indic_retour["Alerte"] = indic_retour["Taux retour (%)"].apply(_alerte_taux_retour)

    rc1, rc2, rc3 = st.columns(3)
    rc1.metric("🔴 Titres à taux de retour ≥ 35 %", int((indic_retour["Taux retour (%)"] >= 35).sum()))
    rc2.metric("dont ≥ 100 % (à vérifier en priorité)", int((indic_retour["Taux retour (%)"] >= 100).sum()))
    rc3.metric("🟠 Titres à surveiller (20 – 35 %)",
               int(((indic_retour["Taux retour (%)"] >= 20) & (indic_retour["Taux retour (%)"] < 35)).sum()))

    aff_retour = indic_retour[["Code_Analytique", "Titre", "CA distributeur", "Retours", "Remises",
                                "Taux retour (%)", "Taux remise (%)", "Résultat net", "Alerte"]] \
                     .sort_values("Taux retour (%)", ascending=False).reset_index(drop=True)

    # Rendu ligne par ligne (plutôt qu'un st.dataframe) pour pouvoir placer un bouton "Voir la
    # fiche" directement à côté de chaque titre — un st.dataframe seul ne permet pas d'y insérer
    # un élément interactif par ligne.
    largeurs_retour = [2.4, 1.0, 0.8, 0.8, 0.9, 0.9, 1.1, 1.3, 0.6]
    en_tetes_retour = ["Titre", "CA distributeur", "Retours", "Remises",
                        "Taux retour", "Taux remise", "Résultat net", "Alerte", ""]
    hc = st.columns(largeurs_retour)
    for c, libelle in zip(hc, en_tetes_retour):
        c.markdown(f"**{libelle}**")

    with st.container(height=480):
        for _, row in aff_retour.iterrows():
            rc = st.columns(largeurs_retour)
            rc[0].write(row["Titre"])
            rc[1].write(f"{fmt_fr(row['CA distributeur'], 0)} €")
            rc[2].write(f"{fmt_fr(row['Retours'], 0)} €")
            rc[3].write(f"{fmt_fr(row['Remises'], 0)} €")
            rc[4].write(f"{row['Taux retour (%)']:.1f} %")
            rc[5].write(f"{row['Taux remise (%)']:.1f} %")
            couleur_resultat = "#991b1b" if row["Résultat net"] < 0 else "#065f46"
            rc[6].markdown(f"<span style='color:{couleur_resultat}'>{fmt_fr(row['Résultat net'], 0)} €</span>",
                            unsafe_allow_html=True)
            rc[7].write(row["Alerte"])
            if rc[8].button("📖", key=f"fiche_retour_{row['Code_Analytique']}", help="Ouvrir la fiche de ce titre"):
                afficher_fiche_titre(row["Code_Analytique"], df, params)

    st.caption("ℹ️ Un taux très élevé (ex. > 100 %) vient parfois d'un CA distributeur très faible sur la "
               "période pour ce titre (dénominateur proche de 0) plutôt que d'un vrai afflux de retours — "
               "vérifiez la colonne **CA distributeur** avant de conclure à une anomalie.")
    st.caption("💡 Repère utile : un titre combinant **taux de retour élevé** et **résultat net faible/négatif** "
               "n'est pas forcément pénalisé par les retours eux-mêmes — vérifiez aussi ses charges de "
               "fabrication (fiche titre) : un tirage surdimensionné par rapport aux ventes réelles peut "
               "peser bien plus lourd sur le résultat que le coût des retours.")

    st.divider()

    st.markdown("Ou sélectionnez directement un titre puis ouvrez sa fiche détaillée — elle s'affiche "
                "dans une fenêtre dédiée avec un mini SIG (soldes intermédiaires de gestion, charges "
                "fixes imputées incluses) et son évolution mensuelle.")
    col_sel, col_btn = st.columns([3, 1])
    with col_sel:
        isbn_sel = st.selectbox("Titre (ISBN)", titres, label_visibility="collapsed",
                                 format_func=lambda c: label_affiche(c, df))
    with col_btn:
        ouvrir = st.button("📖 Ouvrir la fiche", type="primary", use_container_width=True)

    if ouvrir:
        afficher_fiche_titre(isbn_sel, df, params)
# =====================
# SIMULATEUR DE RENTABILITÉ
# =====================
elif page == "🎯 Simulateur de rentabilité":
    st.header("🎯 Simulateur de rentabilité")
    st.caption(
        "Simule l'impact d'hypothèses sur la marge d'un titre — taux de retour, charges variables "
        "(fabrication/commercialisation), clé de répartition des charges indirectes — sans modifier la "
        "comptabilité réelle. Objectif : objectiver une décision de réimpression ou d'arrêt de publication "
        "en chiffrant le seuil de rentabilité à atteindre."
    )
    if "df_pivot" not in st.session_state:
        st.warning("⚠️ Générez d'abord le socle analytique.")
        st.stop()

    df_sim = st.session_state["df_pivot"].copy()
    params_sim = st.session_state["param_comptes"]
    titres_sim = sorted(filtrer_isbn_reels(df_sim)["Code_Analytique"].astype(str).unique().tolist())
    if not titres_sim:
        st.warning("Aucun ISBN/code analytique détecté dans les données.")
        st.stop()

    isbn_sim = st.selectbox(
        "Titre à simuler", titres_sim, format_func=lambda c: label_affiche(c, df_sim), key="sim_isbn"
    )

    # ── Baseline (valeurs réelles constatées, mêmes formules que la fiche titre) ──
    df_t_sim = df_sim[df_sim["Code_Analytique"] == isbn_sim]
    df_v_sim   = df_t_sim[df_t_sim["Compte"].astype(str).str.startswith(tuple(params_sim["ventes"]))]
    df_v_distrib_sim = df_t_sim[df_t_sim["Compte"].astype(str).str.startswith(
        tuple(params_sim.get("ventes_distributeur") or params_sim["ventes"]))]
    df_r_sim   = df_t_sim[mask_retours(df_t_sim, params_sim)]
    df_rem_sim = df_t_sim[mask_remises(df_t_sim, params_sim)]
    prefixes_stock_sim = tuple(params_sim.get("stock") or ["603"])
    df_c_sim = df_t_sim[df_t_sim["Compte"].astype(str).str.startswith(tuple(params_sim["charges"]))
                          & (~df_t_sim["Compte"].astype(str).str.startswith(prefixes_stock_sim))]
    df_stock_sim = df_t_sim[df_t_sim["Compte"].astype(str).str.startswith(prefixes_stock_sim)]
    prefixes_prov_sim = tuple(params_sim.get("provisions_reprises") or [])
    df_prov_sim = (df_t_sim[df_t_sim["Compte"].astype(str).str.startswith(prefixes_prov_sim)]
                   if prefixes_prov_sim else df_t_sim.iloc[0:0])
    net_prov_sim = df_prov_sim["Crédit"].sum() - df_prov_sim["Débit"].sum()
    df_cfi_sim = df_t_sim[df_t_sim["Compte"].astype(str) == "CHARGES INDIRECTES REPARTIES"]

    ventes_ht_b      = df_v_sim["Crédit"].sum()
    ventes_distrib_b = df_v_distrib_sim["Crédit"].sum()
    retours_b        = df_r_sim["Débit"].sum() - df_r_sim["Crédit"].sum()
    remises_b        = df_rem_sim["Débit"].sum() - df_rem_sim["Crédit"].sum()
    charges_v_b      = (df_c_sim["Débit"].sum() - df_c_sim["Crédit"].sum()) - net_prov_sim
    variation_stock_b = df_stock_sim["Débit"].sum() - df_stock_sim["Crédit"].sum()
    charges_fixes_b  = df_cfi_sim["Débit"].sum()
    ca_net_b         = ventes_ht_b - retours_b - remises_b
    marge_brute_b    = ca_net_b - charges_v_b
    resultat_net_b   = marge_brute_b - variation_stock_b - charges_fixes_b
    taux_retour_b    = (retours_b / ventes_distrib_b * 100) if ventes_distrib_b else 0.0

    # Clé de répartition actuelle : nombre de titres actifs et quote-part de charges indirectes,
    # dérivés de la répartition déjà effectuée dans ⚙️ Paramétrage analytique (si activée).
    repartition_detail_sim = st.session_state.get("repartition_detail") or {}
    nb_titres_actuel = repartition_detail_sim.get("nb_titres_actifs", len(titres_sim))
    part_charge_actuelle = repartition_detail_sim.get("part_charge", 0.0)
    total_charges_indirectes_sim = part_charge_actuelle * nb_titres_actuel

    st.subheader("Situation actuelle (constatée)")
    bc1, bc2, bc3, bc4 = st.columns(4)
    bc1.metric("CA net", f"{fmt_fr(ca_net_b, 0)} €")
    bc2.metric("Marge brute", f"{fmt_fr(marge_brute_b, 0)} €")
    bc3.metric("Résultat net", f"{fmt_fr(resultat_net_b, 0)} €")
    bc4.metric("Taux de retour", f"{taux_retour_b:.1f} %")

    st.divider()
    st.subheader("Hypothèses de simulation")
    if not st.session_state.get("repartition_active"):
        st.info("ℹ️ La répartition des charges indirectes n'est pas activée (⚙️ Paramétrage analytique) : "
                 "la simulation de la clé de répartition n'aura donc aucun effet (charges fixes imputées "
                 "actuellement à 0 pour tous les titres).")

    col_h1, col_h2 = st.columns(2)
    with col_h1:
        taux_retour_hyp = st.slider(
            "Taux de retour hypothétique (%)", min_value=0.0, max_value=100.0,
            value=float(round(taux_retour_b, 1)), step=0.5,
            help="Remplace le taux de retour constaté (calculé sur le CA distributeur de ce titre) par "
                 "une hypothèse — ex. l'effet d'une renégociation avec le diffuseur ou d'un changement de "
                 "circuit de vente."
        )
        variation_charges_pct = st.slider(
            "Variation des charges variables (%)", min_value=-50, max_value=50, value=0, step=1,
            help="Simule une renégociation des charges directes (fabrication, commercialisation...) : "
                 "-20% par exemple pour une baisse du coût d'impression sur une réimpression."
        )
    with col_h2:
        nb_titres_hyp = st.slider(
            "Nombre de titres actifs hypothétique (clé de répartition)", min_value=1,
            max_value=max(int(nb_titres_actuel) * 2, int(nb_titres_actuel) + 5), value=int(nb_titres_actuel),
            step=1,
            help="Simule l'effet d'un catalogue plus restreint (ex. arrêt de titres non rentables) sur la "
                 "quote-part de charges indirectes supportée par CE titre — la même masse de charges "
                 "indirectes se répartit alors sur moins de titres, donc une quote-part plus lourde chacun."
        )
        st.caption(f"Actuellement : {int(nb_titres_actuel)} titres actifs, quote-part de "
                   f"{fmt_fr(part_charge_actuelle, 0)} €/titre.")

    # ── Recalcul selon les hypothèses ──
    retours_hyp       = ventes_distrib_b * taux_retour_hyp / 100
    charges_v_hyp      = charges_v_b * (1 + variation_charges_pct / 100)
    ca_net_hyp         = ventes_ht_b - retours_hyp - remises_b
    marge_brute_hyp    = ca_net_hyp - charges_v_hyp
    charges_fixes_hyp  = (total_charges_indirectes_sim / nb_titres_hyp) if (nb_titres_hyp and st.session_state.get("repartition_active")) else charges_fixes_b
    resultat_net_hyp   = marge_brute_hyp - variation_stock_b - charges_fixes_hyp

    st.divider()
    st.subheader("Résultat simulé")
    sc1, sc2, sc3, sc4 = st.columns(4)
    sc1.metric("CA net (simulé)", f"{fmt_fr(ca_net_hyp, 0)} €", delta=f"{fmt_fr(ca_net_hyp - ca_net_b, 0)} €")
    sc2.metric("Marge brute (simulée)", f"{fmt_fr(marge_brute_hyp, 0)} €", delta=f"{fmt_fr(marge_brute_hyp - marge_brute_b, 0)} €")
    sc3.metric("Résultat net (simulé)", f"{fmt_fr(resultat_net_hyp, 0)} €", delta=f"{fmt_fr(resultat_net_hyp - resultat_net_b, 0)} €")
    sc4.metric("Charges fixes imputées (simulées)", f"{fmt_fr(charges_fixes_hyp, 0)} €",
               delta=f"{fmt_fr(charges_fixes_hyp - charges_fixes_b, 0)} €", delta_color="inverse")

    # ── Seuil de rentabilité : taux de retour maximal avant résultat net simulé = 0,
    # les autres hypothèses (charges variables, clé de répartition) étant maintenues telles
    # que réglées ci-dessus. Résultat net linéaire en taux de retour → résolution directe.
    if ventes_distrib_b:
        seuil_taux_retour = (
            (ventes_ht_b - remises_b - charges_v_hyp - variation_stock_b - charges_fixes_hyp) / ventes_distrib_b * 100
        )
        seuil_taux_retour = max(0.0, seuil_taux_retour)
        marge_avant_seuil = seuil_taux_retour - taux_retour_hyp
        if marge_avant_seuil >= 0:
            st.success(
                f"✅ Seuil de rentabilité : ce titre reste bénéficiaire jusqu'à un taux de retour de "
                f"**{seuil_taux_retour:.1f} %** (soit {marge_avant_seuil:.1f} point(s) de marge par rapport "
                f"à l'hypothèse simulée de {taux_retour_hyp:.1f} %), toutes autres hypothèses inchangées."
            )
        else:
            st.error(
                f"❌ Avec les hypothèses actuelles, ce titre est déjà déficitaire : il faudrait un taux de "
                f"retour d'au plus **{seuil_taux_retour:.1f} %** pour repasser à l'équilibre (contre "
                f"{taux_retour_hyp:.1f} % simulé)."
            )
    else:
        st.info("Impossible de calculer un seuil de rentabilité : aucune vente sur le compte distributeur "
                "configuré pour ce titre.")

    with st.expander("Détail du calcul"):
        df_detail_sim = pd.DataFrame([
            {"Poste": "Ventes HT", "Actuel (€)": ventes_ht_b, "Simulé (€)": ventes_ht_b},
            {"Poste": "− Retours", "Actuel (€)": -retours_b, "Simulé (€)": -retours_hyp},
            {"Poste": "− Remises", "Actuel (€)": -remises_b, "Simulé (€)": -remises_b},
            {"Poste": "= CA net", "Actuel (€)": ca_net_b, "Simulé (€)": ca_net_hyp},
            {"Poste": "− Charges variables", "Actuel (€)": -charges_v_b, "Simulé (€)": -charges_v_hyp},
            {"Poste": "= Marge brute", "Actuel (€)": marge_brute_b, "Simulé (€)": marge_brute_hyp},
            {"Poste": "− Variation de stock", "Actuel (€)": -variation_stock_b, "Simulé (€)": -variation_stock_b},
            {"Poste": "− Charges fixes imputées", "Actuel (€)": -charges_fixes_b, "Simulé (€)": -charges_fixes_hyp},
            {"Poste": "= Résultat net", "Actuel (€)": resultat_net_b, "Simulé (€)": resultat_net_hyp},
        ])
        st.dataframe(
            df_detail_sim.style.format({"Actuel (€)": (lambda x: fmt_fr(x, 0)), "Simulé (€)": (lambda x: fmt_fr(x, 0))}),
            use_container_width=True, hide_index=True
        )
# =====================
# TRÉSORERIE PRÉVISIONNELLE
# =====================
elif page == "💰 Trésorerie prévisionnelle":
    st.header("💰 Trésorerie prévisionnelle")
    if "df_pivot" not in st.session_state:
        st.warning("⚠️ Générez d'abord le socle analytique.")
        st.stop()

    st.caption(
        "Tableau de flux de trésorerie (TFT) reconstitué à partir du grand livre — méthode directe, classement "
        "par nature de compte (encaissements clients, paiements fournisseurs, dettes sociales/fiscales, "
        "investissement, financement), sur le modèle du tableau de flux fourni. Les montants sont calculés à "
        "partir des dates d'enregistrement comptable et peuvent différer légèrement d'un relevé bancaire réel "
        "en cas de décalage d'encaissement/décaissement."
    )

    df_tr = st.session_state["df_pivot"].copy()
    df_tr["Compte"] = df_tr["Compte"].astype(str).str.strip()
    df_tr["Date"]   = pd.to_datetime(df_tr["Date"], errors="coerce")
    df_tr["Débit"]  = pd.to_numeric(df_tr["Débit"], errors="coerce").fillna(0)
    df_tr["Crédit"] = pd.to_numeric(df_tr["Crédit"], errors="coerce").fillna(0)
    if "Journal" in df_tr.columns:
        df_tr["Journal"] = df_tr["Journal"].astype(str).str.strip()
    else:
        df_tr["Journal"] = ""
    df_tr = df_tr.dropna(subset=["Date"])
    if df_tr.empty:
        st.warning("Aucune écriture datée dans le socle analytique.")
        st.stop()

    # ---- Mapping par défaut des lignes du TFT (plan comptable général) ----
    DEFAULT_MAPPING = [
        ("Exploitation",   "Encaissements clients",                                  "411",             "credit"),
        ("Exploitation",   "Paiements fournisseurs",                                 "401,408",         "debit_neg"),
        ("Exploitation",   "Paiements dettes sociales - Salariés",                   "421,428",         "debit_neg"),
        ("Exploitation",   "Paiements dettes sociales - Organismes",                 "431,437,438",     "debit_neg"),
        ("Exploitation",   "Paiements dettes fiscales - TVA",                        "445510000",       "debit_neg"),
        ("Exploitation",   "Paiements dettes fiscales - IS",                         "444",             "debit_neg"),
        ("Exploitation",   "Encaissements aides et subventions d'exploitation",      "740",             "credit"),
        ("Exploitation",   "Autres débiteurs et créditeurs",                         "419,446,448",     "net"),
        ("Investissement", "Acquisitions d'immobilisations incorp. et corp.",        "201,205,215,218", "debit_neg"),
        ("Investissement", "Acquisitions/Cessions d'immobilisations financières",    "271,275,277",     "net"),
        ("Investissement", "Produits de cession des immobilisations",               "675,775",         "credit"),
        ("Financement",    "Variation des comptes courants d'associés",             "455",             "net"),
        ("Financement",    "Augmentation/(Diminution) du capital social",           "101",             "net"),
        ("Financement",    "Subventions d'investissement",                          "13",              "net"),
        ("Financement",    "Billets de trésorerie",                                 "4673,4674",       "net"),
        ("Financement",    "Encaissement d'emprunt",                                "164,168",         "credit"),
        ("Financement",    "Décaissement d'emprunt",                                "164,168",         "debit_neg"),
    ]
    SECTION_ORDER = ["Exploitation", "Investissement", "Financement"]

    with st.expander("⚙️ Configuration des comptes du TFT (avancé)"):
        st.caption("Préfixes de comptes (séparés par des virgules) utilisés pour classer chaque ligne du tableau. "
                   "À adapter si votre plan comptable diffère. Note : les comptes 438/604300000 etc. liés aux "
                   "droits d'auteurs sont également suivis dans le module « Droits d'auteurs ».")
        mapping = []
        for section, label, prefixes_def, mode in DEFAULT_MAPPING:
            p = st.text_input(f"{section} — {label}", value=prefixes_def, key=f"tft_pfx_{label}")
            mapping.append((section, label, [x.strip() for x in p.split(",") if x.strip()], mode))
        comptes_banque = st.text_input("Comptes de trésorerie (banque/caisse)", value="512,530,580", key="tft_banque")
        comptes_banque = tuple(x.strip() for x in comptes_banque.split(",") if x.strip())
        journaux_exclus_txt = st.text_input(
            "Journaux exclus des flux (report à nouveau)", value="AN", key="tft_journaux_exclus",
            help="Les écritures de report à nouveau (reprise des soldes d'ouverture en début d'exercice, "
                 "souvent codées « AN ») ne sont pas de vrais flux de la période : elles sont exclues du "
                 "classement, mais utilisées pour suggérer la trésorerie à l'ouverture (solde bancaire repris)."
        )
        journaux_exclus = tuple(x.strip() for x in journaux_exclus_txt.split(",") if x.strip())

    def montants_par_mois(sub_df, mode):
        """Agrégation vectorisée par mois — gère proprement le cas d'un sous-ensemble vide."""
        if sub_df.empty:
            return {}
        deb = sub_df.groupby("Mois")["Débit"].sum()
        cre = sub_df.groupby("Mois")["Crédit"].sum()
        if mode == "credit":
            res = cre
        elif mode == "debit_neg":
            res = -deb
        elif mode == "net":
            res = cre.sub(deb, fill_value=0)
        else:
            res = cre * 0
        return res.to_dict()

    mask_an = df_tr["Journal"].isin(journaux_exclus) if journaux_exclus else pd.Series(False, index=df_tr.index)
    df_tr_flux = df_tr[~mask_an].copy()

    col1, col2 = st.columns(2)
    with col1:
        date_debut = st.date_input("Date de départ (réalisé)", df_tr["Date"].min())
        banque_an = df_tr[mask_an & df_tr["Compte"].str.startswith(comptes_banque)]
        if not banque_an.empty:
            solde_suggere = float(banque_an["Débit"].sum() - banque_an["Crédit"].sum())
        else:
            banque_avant = df_tr_flux[df_tr_flux["Compte"].str.startswith(comptes_banque) & (df_tr_flux["Date"] < pd.to_datetime(date_debut))]
            solde_suggere = float(banque_avant["Débit"].sum() - banque_avant["Crédit"].sum())
        tresorerie_ouverture = st.number_input(
            "Trésorerie à l'ouverture (€)", value=solde_suggere, step=100.0,
            help="Solde bancaire à la date de départ. Suggéré depuis les écritures de report à nouveau sur les "
                 "comptes de trésorerie si elles existent, sinon depuis les écritures antérieures à cette date "
                 "(0 si aucune des deux — à saisir manuellement dans ce cas)."
        )
    with col2:
        horizon = st.slider("Horizon de projection (mois)", 0, 24, 6)
        st.markdown("**Scénarios de projection (au-delà du réalisé)**")
        croissance_opt  = st.number_input("Croissance encaissements — optimiste (%/mois)", value=4.0, step=0.5) / 100
        croissance_cent = st.number_input("Croissance encaissements — central (%/mois)", value=2.0, step=0.5) / 100
        croissance_pess = st.number_input("Croissance encaissements — pessimiste (%/mois)", value=0.0, step=0.5) / 100
        evolution_charges = st.number_input("Évolution charges décaissées (%/mois)", value=1.0, step=0.5) / 100

    df_period = df_tr_flux[df_tr_flux["Date"] >= pd.to_datetime(date_debut)].copy()
    if df_period.empty:
        st.warning("Aucune écriture après la date de départ.")
        st.stop()
    df_period["Mois"] = df_period["Date"].dt.to_period("M")
    mois_realises = sorted(df_period["Mois"].unique())

    lignes = {}
    for section, label, prefixes, mode in mapping:
        if not prefixes:
            continue
        sub = df_period[df_period["Compte"].str.startswith(tuple(prefixes))]
        lignes[(section, label)] = montants_par_mois(sub, mode)

    table = pd.DataFrame(0.0, index=pd.MultiIndex.from_tuples(lignes.keys(), names=["Section", "Ligne"]), columns=mois_realises)
    for key, vals in lignes.items():
        for m, v in vals.items():
            table.loc[key, m] = v

    sous_totaux = table.groupby(level="Section", sort=False).sum()
    flux_net_total = table.sum()
    treso_real = tresorerie_ouverture + flux_net_total.cumsum()

    # Mis en session pour réutilisation par le Référentiel des indicateurs de pilotage
    # (indicateur 2 — Trésorerie prévisionnelle), sans avoir à tout recalculer.
    st.session_state["treso_table"] = table
    st.session_state["treso_real"] = treso_real
    st.session_state["treso_ouverture"] = tresorerie_ouverture
    st.session_state["treso_mois_realises"] = mois_realises

    _MOIS_FR = ["", "Janv.", "Févr.", "Mars", "Avr.", "Mai", "Juin", "Juil.", "Août", "Sept.", "Oct.", "Nov.", "Déc."]
    def mois_label(p):
        return f"{_MOIS_FR[p.month]} {p.year}"

    # ---- Indicateurs ----
    m1, m2, m3 = st.columns(3)
    m1.metric("Trésorerie à l'ouverture", f"{fmt_fr(tresorerie_ouverture, 0)} €")
    m2.metric("Trésorerie à la clôture (réalisé)", f"{fmt_fr(treso_real.iloc[-1], 0)} €",
              delta=f"{fmt_fr(treso_real.iloc[-1] - tresorerie_ouverture, 0)} €")
    m3.metric("Flux net généré (période réalisée)", f"{fmt_fr(flux_net_total.sum(), 0)} €")

    # ---- Tableau détaillé (réalisé) ----
    display_rows, display_index = [], []
    for section in SECTION_ORDER:
        if section not in table.index.get_level_values("Section"):
            continue
        for key in table.index:
            if key[0] == section:
                display_rows.append(table.loc[key])
                display_index.append(key[1])
        display_rows.append(sous_totaux.loc[section])
        display_index.append(f"▶ Flux net — {section}")
    display_rows.append(treso_real)
    display_index.append("● Trésorerie à la clôture (cumulée)")

    df_affiche = pd.DataFrame(display_rows, index=display_index)
    df_affiche.columns = [mois_label(m) for m in df_affiche.columns]

    def style_lignes(row):
        if row.name.startswith("▶") or row.name.startswith("●"):
            return ["font-weight: bold; background-color: rgba(59,130,246,0.12)"] * len(row)
        return [""] * len(row)

    st.dataframe(df_affiche.style.apply(style_lignes, axis=1).format((lambda x: f"{fmt_fr(x, 0)} €")), use_container_width=True)

    # ---- Projection (scénarios) ----
    def base_ligne(key):
        vals = table.loc[key]
        return vals.iloc[-3:].mean() if len(vals) >= 3 else vals.mean()

    def construire_projection(taux_encaissement, taux_charges):
        futurs = [mois_realises[-1] + i for i in range(1, horizon + 1)]
        proj = pd.DataFrame(0.0, index=table.index, columns=futurs)
        for key in table.index:
            section, label = key
            if section == "Exploitation" and "encaissement" in label.lower():
                taux = taux_encaissement
            elif section == "Exploitation":
                taux = taux_charges
            else:
                # Pas de nouvelle opération d'investissement/financement supposée par défaut
                continue
            v = base_ligne(key)
            for m in futurs:
                v *= (1 + taux)
                proj.loc[key, m] = v
        return proj

    if horizon > 0:
        scenarios = {
            "Optimiste":  construire_projection(croissance_opt, evolution_charges),
            "Central":    construire_projection(croissance_cent, evolution_charges),
            "Pessimiste": construire_projection(croissance_pess, evolution_charges),
        }
        st.session_state["treso_scenarios"] = scenarios
        courbes = {}
        for nom, proj in scenarios.items():
            flux_proj = proj.sum()
            treso_proj = treso_real.iloc[-1] + flux_proj.cumsum()
            courbes[nom] = pd.concat([treso_real, treso_proj])

        fig = go.Figure()
        fig.add_trace(go.Scatter(x=[mois_label(m) for m in treso_real.index], y=treso_real.values,
                                  name="Réalisé", line=dict(color="#111827", width=2.5)))
        couleurs = {"Optimiste": "#10B981", "Central": "#3B82F6", "Pessimiste": "#EF4444"}
        for nom, serie in courbes.items():
            proj_part = serie.iloc[len(treso_real) - 1:]
            fig.add_trace(go.Scatter(x=[mois_label(m) for m in proj_part.index], y=proj_part.values,
                                      name=nom, line=dict(color=couleurs[nom], width=2, dash="dot")))
        fig.add_hline(y=0, line_dash="dash", line_color="gray", annotation_text="Seuil zéro")
        fig.update_layout(title="Trésorerie — réalisé et projection à 3 scénarios",
                          xaxis_title="Mois", yaxis_title="Trésorerie cumulée (€)",
                          legend=dict(orientation="h"), height=420)
        st.plotly_chart(fig, use_container_width=True)

        treso_cent_complete = courbes["Central"]
        if treso_cent_complete.min() < 0:
            mois_neg = treso_cent_complete[treso_cent_complete < 0].index[0]
            st.error(f"⚠️ Alerte : la trésorerie passe en négatif en **{mois_label(mois_neg)}** dans le scénario central !")
    else:
        scenarios = {}
        st.info("Augmentez l'horizon de projection pour visualiser les scénarios optimiste / central / pessimiste.")

    # ---- Export Excel ----
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_affiche.to_excel(writer, sheet_name="Réalisé")
        for nom, proj in scenarios.items():
            rows, idx = [], []
            for section in SECTION_ORDER:
                for key in proj.index:
                    if key[0] == section:
                        rows.append(proj.loc[key]); idx.append(key[1])
                rows.append(proj[proj.index.get_level_values("Section") == section].sum()); idx.append(f"▶ Flux net — {section}")
            flux_proj = proj.sum()
            rows.append(tresorerie_ouverture + pd.concat([flux_net_total, flux_proj]).cumsum().iloc[len(flux_net_total):])
            idx.append("● Trésorerie à la clôture (cumulée)")
            df_export = pd.DataFrame(rows, index=idx)
            df_export.columns = [mois_label(m) for m in df_export.columns]
            df_export.to_excel(writer, sheet_name=f"Prévisionnel_{nom}"[:31])
        wb = writer.book
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                label_cell = row[0]
                if isinstance(label_cell.value, str) and (label_cell.value.startswith("▶") or label_cell.value.startswith("●")):
                    for cell in row:
                        cell.font = cell.font.copy(bold=True)
    buffer.seek(0)
    st.download_button("📥 Exporter le tableau de flux de trésorerie (Excel)", buffer,
                        file_name="Tableau_Flux_Tresorerie.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ================================================
    # INDICATEUR 2 — SOUS-INDICATEURS DE PILOTAGE
    # ================================================
    st.divider()
    with st.expander("📐 Sous-indicateurs de pilotage (BFR éditorial, DSO BLDD, seuil de CA, cumul 6 mois)"):
        seuil_securite = st.number_input(
            "Seuil de sécurité défini avec le dirigeant (€)", value=0.0, step=500.0,
            help="Seuil orange : la trésorerie prévisionnelle du mois à venir passe sous ce montant.", key="ind2_seuil_secu"
        )
        solde_prev_prochain_mois = None
        if horizon > 0 and scenarios:
            flux_central_i2 = scenarios["Central"].sum()
            if len(flux_central_i2) > 0:
                solde_prev_prochain_mois = treso_real.iloc[-1] + flux_central_i2.iloc[0]
        if solde_prev_prochain_mois is not None:
            if solde_prev_prochain_mois < 0:
                alerte_i2 = "🔴 Tension critique"
            elif solde_prev_prochain_mois < seuil_securite:
                alerte_i2 = "🟠 Sous le seuil de sécurité"
            else:
                alerte_i2 = "🟢 Normal"
            st.metric("Solde prévisionnel — mois à venir", f"{fmt_fr(treso_real.iloc[-1] + flux_central_i2.iloc[0], 0)} €", delta=alerte_i2)

        col_bfr1, col_bfr2 = st.columns(2)
        cpt_fabrication_i2 = col_bfr1.text_input("Comptes fabrication (décaissements)", value="604,605", key="ind2_cpt_fab")
        cpt_bldd_i2 = col_bfr2.text_input("Compte client BLDD (encaissements)", value="411100011", key="ind2_cpt_bldd")
        prefixes_fab_i2 = tuple(x.strip() for x in cpt_fabrication_i2.split(",") if x.strip())
        decaissements_fab_i2 = df_tr[df_tr["Compte"].str.startswith(prefixes_fab_i2)]["Débit"].sum() if prefixes_fab_i2 else 0.0
        df_bldd_i2 = df_tr[df_tr["Compte"] == cpt_bldd_i2]
        encaissements_bldd_i2 = df_bldd_i2["Crédit"].sum()
        bfr_editorial = decaissements_fab_i2 - encaissements_bldd_i2
        st.metric("BFR éditorial (décaissements fabrication − encaissements BLDD)", f"{fmt_fr(bfr_editorial, 0)} €")

        if not df_bldd_i2.empty and df_bldd_i2["Débit"].sum() > 0:
            solde_client_bldd = df_bldd_i2["Débit"].sum() - df_bldd_i2["Crédit"].sum()
            ca_facture_bldd = df_bldd_i2["Débit"].sum()
            nb_jours_periode_i2 = (df_tr["Date"].dropna().max() - df_tr["Date"].dropna().min()).days or 365
            dso_bldd = solde_client_bldd / ca_facture_bldd * nb_jours_periode_i2
            st.metric("Délai moyen d'encaissement BLDD (DSO estimé)", f"{fmt_fr(dso_bldd, 0)} jours")
            st.caption("Estimé par : solde du compte client BLDD à date / CA facturé BLDD sur la période × nombre "
                       "de jours de la période (proxy DSO standard).")
        else:
            st.caption("ℹ️ Aucune écriture trouvée sur le compte client BLDD indiqué — DSO non calculable.")

        label_ci_i2 = st.session_state.get("labels_indirect", {}).get("charges", "CHARGES INDIRECTES")
        masque_ci_i2 = df_tr["Compte"].astype(str).str.strip() == label_ci_i2
        charges_fixes_totales_i2 = (df_tr[masque_ci_i2]["Débit"] - df_tr[masque_ci_i2]["Crédit"]).sum()
        if charges_fixes_totales_i2 == 0 and "df_pivot" in st.session_state:
            df_repartie = st.session_state["df_pivot"]
            masque_cfi_i2 = df_repartie["Compte"].astype(str) == "CHARGES INDIRECTES REPARTIES"
            charges_fixes_totales_i2 = df_repartie[masque_cfi_i2]["Débit"].sum()
        nb_mois_periode_i2 = max(1, round((df_tr["Date"].dropna().max() - df_tr["Date"].dropna().min()).days / 30.44))
        charges_fixes_mensuelles_i2 = charges_fixes_totales_i2 / nb_mois_periode_i2
        df_v_i2 = df_tr[df_tr["Compte"].str.startswith(("701",))]
        df_rem_i2 = df_tr[df_tr["Compte"].str.startswith(("7091",))]
        taux_remise_moyen_i2 = (df_rem_i2["Débit"].sum() / df_v_i2["Crédit"].sum()) if df_v_i2["Crédit"].sum() else 0.0
        seuil_mensuel_ca = (charges_fixes_mensuelles_i2 / (1 - taux_remise_moyen_i2)
                            if taux_remise_moyen_i2 < 1 else charges_fixes_mensuelles_i2)
        st.metric("Seuil mensuel de CA (charges fixes / (1 − taux de remise moyen))", f"{fmt_fr(seuil_mensuel_ca, 0)} €")
        st.caption("⚠️ Les comptes ventes (701) et remises (7091) utilisés ici sont ceux du secteur livre standard "
                   "— ajustez-les via ⚙️ Paramétrage analytique si les vôtres diffèrent.")

        if horizon > 0 and scenarios:
            st.markdown("**Cumul de trésorerie sur 6 mois** _(scénario central)_")
            cumul_6m = treso_real.iloc[-1] + scenarios["Central"].sum().iloc[:6].cumsum()
            st.line_chart(cumul_6m)
# =====================
# DROITS D'AUTEURS
# =====================
elif page == "✍️ Droits d'auteurs":
    st.header("📚 ROYALTIES EDITION — Droits d'auteurs & URSSAF")

    # ──────────────────────────────────────────────
    # TAUX URSSAF (précompte diffuseur — en vigueur 2024)
    # Le diffuseur/éditeur prélève ces cotisations sur les droits bruts avant
    # de les verser à l'auteur, puis les reverse à l'URSSAF.
    # ──────────────────────────────────────────────
    TAUX_CSG_CRDS        = 0.097   # CSG (9,2%) + CRDS (0,5%) = 9,7%
    TAUX_FORMATION_PROF  = 0.01    # Contribution formation professionnelle = 1,0%
    # Cotisation retraite complémentaire RAAP (variable selon revenus) : on laisse paramétrable
    # Attention : ces taux s'appliquent sur 98,25% des droits bruts (abattement de 1,75%)
    ASSIETTE_COEFF       = 0.9825  # Abattement forfaitaire frais professionnels

    st.info("""
    **Comment fonctionne ce module ?**

    Les droits d'auteurs **déjà comptabilisés** sont lus directement dans votre comptabilité analytique
    (onglet Réel, qui fait foi). Un simulateur simple reste disponible pour estimer un montant
    *avant* la provision de clôture, à titre indicatif.

    1. **📋 Référentiel** : correspondance ISBN ↔ auteur (détectée automatiquement si vous utilisez une
       famille analytique AUTEUR ; à défaut, à compléter manuellement) + contrats (taux, paliers)
    2. **🧮 Simulateur** : estimation des droits bruts (taux forfaitaire ou paliers) puis du précompte
       URSSAF et de la contribution diffuseur — à titre indicatif, avant provision
    3. **📒 Réel (comptabilisé)** : montants réellement provisionnés, lus directement dans le grand livre
    4. **📄 Relevés par auteur** : relevé de compte par auteur (priorité au réel), exportable en Excel
    """)

    # ──────────────────────────────────────────────
    # INITIALISATION DU RÉFÉRENTIEL EN SESSION
    # Structure : liste de dicts {auteur, isbn, titre, taux_base, paliers, part_auteur}
    # paliers : liste de {seuil, taux} — ex. [{seuil:0, taux:10}, {seuil:10000, taux:12}]
    # Ce référentiel sert (a) de repli si aucune famille analytique AUTEUR n'est mappée dans le
    # grand livre (correspondance ISBN → auteur, part de chacun en cas de co-auteurs), et (b) de
    # base de taux (forfaitaire ou paliers) pour le 🧮 Simulateur.
    # ──────────────────────────────────────────────
    if "royalties_referentiel" not in st.session_state:
        st.session_state["royalties_referentiel"] = []

    onglet1, onglet2, onglet3, onglet4 = st.tabs([
        "📋 Référentiel",
        "🧮 Simulateur",
        "📒 Réel (comptabilisé)",
        "📄 Relevés par auteur"
    ])

    # ══════════════════════════════════════════════
    # ONGLET 1 — RÉFÉRENTIEL
    # ══════════════════════════════════════════════
    with onglet1:
        mapping_auto_preview = resoudre_mapping_auteurs(st.session_state.get("df_pivot")) if "df_pivot" in st.session_state else {}
        if mapping_auto_preview:
            st.success(f"✅ {len(mapping_auto_preview)} correspondance(s) ISBN → auteur détectée(s) "
                       f"automatiquement depuis une famille analytique AUTEUR de votre export — pas besoin "
                       f"de les ressaisir ci-dessous pour l'onglet **📒 Réel (comptabilisé)**.")
            st.dataframe(
                pd.DataFrame([{"ISBN": k, "Auteur (détecté)": v} for k, v in mapping_auto_preview.items()]),
                use_container_width=True
            )
            st.caption("Le référentiel ci-dessous reste utile en cas de correspondance ISBN → auteur "
                       "manquante ou pour préciser la répartition entre co-auteurs.")

            # ── Statut fiscal des auteurs détectés automatiquement ──
            # Le régime (BNC vs option pour les traitements et salaires) est un choix
            # contractuel de l'auteur, invisible depuis la seule comptabilité analytique :
            # il conditionne pourtant si le diffuseur doit précompter les cotisations
            # sociales (CSG/CRDS, formation professionnelle, RAAP) en plus de la
            # contribution diffuseur (due dans tous les cas). Demandé explicitement ici
            # plutôt que déduit, pour ne jamais se tromper de régime dans la déclaration.
            st.markdown("**Statut fiscal de chaque auteur** _(conditionne le précompte URSSAF)_")
            if "statut_fiscal_auto" not in st.session_state:
                st.session_state["statut_fiscal_auto"] = {}
            # Coercition défensive en str + filtrage des valeurs vides/NaN résiduelles : même si
            # resoudre_mapping_auteurs() est censé les exclure en amont, on ne prend pas le risque
            # qu'une valeur non-str (float NaN notamment) fasse échouer sorted()/set() (comparaison
            # float/str impossible) et casse toute la page.
            auteurs_detectes = sorted({
                str(v).strip() for v in mapping_auto_preview.values()
                if pd.notna(v) and str(v).strip() not in ("", "nan", "None")
            })
            for auteur_det in auteurs_detectes:
                valeur_defaut = st.session_state["statut_fiscal_auto"].get(auteur_det, "Traitements et salaires (option assimilé)")
                st.session_state["statut_fiscal_auto"][auteur_det] = st.selectbox(
                    f"Statut fiscal — {auteur_det}",
                    ["BNC (droits d'auteur)", "Traitements et salaires (option assimilé)"],
                    index=0 if valeur_defaut.startswith("BNC") else 1,
                    key=f"statut_fiscal_auto_{auteur_det}",
                    help="BNC : l'auteur gère lui-même ses cotisations, seule la contribution diffuseur est due "
                         "par l'éditeur. Traitements et salaires : le diffuseur précompte en plus CSG/CRDS, "
                         "formation professionnelle et RAAP, à reverser à l'URSSAF."
                )
        else:
            st.info("💡 Si votre logiciel comptable permet de tagger une famille analytique supplémentaire "
                    "nommée **AUTEUR** (en plus d'EDITION/COMMUNICATION/Types de dépenses) — mappable dans "
                    "⚙️ Paramétrage analytique — la correspondance ISBN → auteur sera détectée automatiquement "
                    "ici. En l'absence de cette famille, complétez le référentiel manuel ci-dessous.")

        st.subheader("Saisie des contrats auteurs")

        st.markdown("**Ajouter un contrat**")
        col_a, col_b = st.columns(2)
        with col_a:
            inp_auteur  = st.text_input("Nom de l'auteur")
            inp_isbn    = st.text_input("ISBN (code analytique exact du SOCLE)")
            inp_titre   = st.text_input("Titre du livre (pour l'affichage)")
        with col_b:
            inp_assiette = st.selectbox(
                "Assiette de calcul",
                ["CA net HT (après retours et remises)", "CA brut HT", "Prix public HT"]
            )
            inp_part = st.number_input(
                "Part de cet auteur (%)", min_value=0.0, max_value=100.0, value=100.0, step=1.0,
                help="Si plusieurs auteurs partagent les droits, indiquez la part de chacun. La somme doit faire 100%."
            )
            inp_statut = st.selectbox(
                "Statut fiscal de l'auteur",
                ["Traitements et salaires (option assimilé)", "BNC (droits d'auteur)"],
                index=0,
                help="Traitements et salaires : le diffuseur précompte CSG/CRDS, formation professionnelle et "
                     "RAAP en plus de la contribution diffuseur, à reverser à l'URSSAF (préréglé par défaut pour "
                     "ce client). BNC : l'auteur gère lui-même ses cotisations, seule la contribution diffuseur "
                     "est due par l'éditeur."
            )

        st.markdown("**Taux des droits**")
        mode_taux = st.radio(
            "Mode de calcul",
            ["Taux forfaitaire (un seul taux, quel que soit le CA)", "Paliers progressifs (le taux change selon le CA)"],
            horizontal=True,
            key="mode_taux_contrat"
        )

        if mode_taux.startswith("Taux forfaitaire"):
            taux_forfait = st.number_input("Taux forfaitaire (%)", value=10.0, step=0.5, key="taux_forfait_contrat")
            paliers = [{"seuil": 0, "taux": taux_forfait}]
        else:
            st.caption("Saisissez les paliers de CA à partir desquels le taux change.")
            nb_paliers = st.number_input("Nombre de paliers", min_value=2, max_value=5, value=2, step=1)
            paliers = []
            for i in range(int(nb_paliers)):
                pc1, pc2 = st.columns(2)
                with pc1:
                    seuil = st.number_input(
                        f"Palier {i+1} — CA à partir de (€)", value=0 if i == 0 else i * 5000,
                        key=f"seuil_{i}"
                    )
                with pc2:
                    taux = st.number_input(
                        f"Palier {i+1} — Taux (%)", value=10.0 if i == 0 else 12.0,
                        key=f"taux_{i}"
                    )
                paliers.append({"seuil": seuil, "taux": taux})

        if st.button("➕ Ajouter ce contrat"):
            if inp_auteur and inp_isbn:
                st.session_state["royalties_referentiel"].append({
                    "auteur":    inp_auteur,
                    "isbn":      inp_isbn.strip(),
                    "titre":     inp_titre or inp_isbn,
                    "assiette":  inp_assiette,
                    "statut_fiscal": inp_statut,
                    "part":      inp_part,
                    "paliers":   paliers
                })
                st.success(f"✅ Contrat ajouté : {inp_auteur} / {inp_titre or inp_isbn}")
            else:
                st.warning("Veuillez renseigner au minimum l'auteur et l'ISBN.")

        # ── Import CSV du référentiel ──
        st.markdown("---")
        st.markdown("**Import du référentiel depuis un fichier CSV**")
        st.caption("Le fichier doit avoir les colonnes : auteur, isbn, titre, assiette, part, statut_fiscal "
                   "(BNC ou TS), seuil_1, taux_1, seuil_2, taux_2, ...")
        fichier_ref = st.file_uploader("Importer le référentiel (CSV)", type=["csv"], key="ref_csv")
        if fichier_ref:
            try:
                df_ref_import = pd.read_csv(fichier_ref)
                df_ref_import.columns = df_ref_import.columns.str.strip().str.lower()
                for _, row in df_ref_import.iterrows():
                    paliers_import = []
                    i = 1
                    while f"seuil_{i}" in row and f"taux_{i}" in row:
                        if pd.notna(row[f"seuil_{i}"]) and pd.notna(row[f"taux_{i}"]):
                            paliers_import.append({"seuil": float(row[f"seuil_{i}"]), "taux": float(row[f"taux_{i}"])})
                        i += 1
                    if not paliers_import:
                        paliers_import = [{"seuil": 0, "taux": 10.0}]
                    statut_brut = str(row.get("statut_fiscal", "TS")).strip().upper()
                    statut_import = ("BNC (droits d'auteur)" if statut_brut in ("BNC",)
                                       else "Traitements et salaires (option assimilé)")
                    st.session_state["royalties_referentiel"].append({
                        "auteur":   str(row.get("auteur", "")),
                        "isbn":     str(row.get("isbn", "")).strip(),
                        "titre":    str(row.get("titre", row.get("isbn", ""))),
                        "assiette": str(row.get("assiette", "CA net HT (après retours et remises)")),
                        "statut_fiscal": statut_import,
                        "part":     float(row.get("part", 100)),
                        "paliers":  paliers_import
                    })
                st.success(f"✅ {len(df_ref_import)} contrats importés.")
            except Exception as e:
                st.error(f"Erreur lors de l'import : {e}")

        # ── Affichage et gestion du référentiel ──
        st.markdown("---")
        st.subheader("Contrats enregistrés")
        ref = st.session_state["royalties_referentiel"]
        if ref:
            rows_display = []
            for idx, c in enumerate(ref):
                paliers_str = " | ".join(
                    f">{fmt_fr(p['seuil'], 0)}€ → {p['taux']}%" for p in c["paliers"]
                )
                rows_display.append({
                    "#": idx,
                    "Auteur": c["auteur"],
                    "ISBN": c["isbn"],
                    "Titre": c["titre"],
                    "Part (%)": c["part"],
                    "Assiette": c["assiette"],
                    "Statut fiscal": c.get("statut_fiscal", "Traitements et salaires (option assimilé)"),
                    "Paliers": paliers_str
                })
            df_ref_display = pd.DataFrame(rows_display)
            st.dataframe(df_ref_display, use_container_width=True)

            # Suppression
            idx_suppr = st.number_input("Supprimer le contrat n°", min_value=0, max_value=len(ref)-1, step=1)
            if st.button("🗑️ Supprimer ce contrat"):
                st.session_state["royalties_referentiel"].pop(int(idx_suppr))
                st.rerun()

            # Export CSV du référentiel
            buffer_ref = BytesIO()
            df_ref_display.to_csv(buffer_ref, index=False)
            buffer_ref.seek(0)
            st.download_button(
                "📥 Exporter le référentiel (CSV)", buffer_ref,
                file_name="referentiel_contrats.csv", mime="text/csv"
            )
        else:
            st.info("Aucun contrat enregistré. Commencez par en ajouter un ci-dessus.")

    # ══════════════════════════════════════════════
    # ONGLET 2 — SIMULATEUR (DROITS D'AUTEUR + URSSAF)
    # Estimation simple, à titre indicatif avant provision de clôture. Le taux peut être
    # forfaitaire (un seul taux) ou progressif par paliers de CA — cf. mode choisi dans le
    # référentiel de contrats (onglet précédent).
    # ══════════════════════════════════════════════
    with onglet2:
        st.subheader("Simulateur — droits d'auteur & précompte URSSAF")
        st.caption("⚠️ Estimation indicative (taux/paliers saisis dans le référentiel), avant provision de "
                   "clôture. Les montants réellement dus et déclarés sont dans l'onglet **📒 Réel (comptabilisé)**.")

        if "df_pivot" not in st.session_state:
            st.warning("⚠️ Générer d'abord le SOCLE EDITION.")
        elif not st.session_state["royalties_referentiel"]:
            st.warning("⚠️ Saisir au moins un contrat dans l'onglet Référentiel.")
        else:
            df_pivot_sim = st.session_state["df_pivot"].copy()
            params_sim   = st.session_state["param_comptes"]

            # ── Étape 1 : droits bruts (CA par ISBN × taux forfaitaire ou paliers) ──
            st.markdown("**Étape 1 — Droits bruts**")

            def ca_isbn_sim(df, prefix_list):
                if not prefix_list:
                    return pd.Series(dtype=float)
                mask = df["Compte"].astype(str).str.startswith(tuple(prefix_list))
                return df[mask].groupby("Code_Analytique")["Crédit"].sum() \
                     - df[mask].groupby("Code_Analytique")["Débit"].sum()

            ca_brut_sim    = ca_isbn_sim(df_pivot_sim, params_sim["ventes"])
            ca_retours_sim = ca_isbn_sim(df_pivot_sim, params_sim["retours"])
            ca_remises_sim = ca_isbn_sim(df_pivot_sim, params_sim["remises"])

            def calcul_droits_paliers_sim(base, paliers):
                """Calcule les droits dus selon des paliers progressifs (ou un taux forfaitaire
                si un seul palier avec seuil=0 est fourni)."""
                paliers_sorted = sorted(paliers, key=lambda p: p["seuil"])
                droits = 0.0
                for i, palier in enumerate(paliers_sorted):
                    seuil_bas = palier["seuil"]
                    seuil_haut = paliers_sorted[i+1]["seuil"] if i+1 < len(paliers_sorted) else float("inf")
                    if base <= seuil_bas:
                        break
                    tranche = min(base, seuil_haut) - seuil_bas
                    droits += tranche * palier["taux"] / 100
                return droits

            resultats_sim = []
            for contrat in st.session_state["royalties_referentiel"]:
                isbn = contrat["isbn"]
                ca_b  = float(ca_brut_sim.get(isbn, 0))
                ca_r  = float(ca_retours_sim.get(isbn, 0))
                ca_re = float(ca_remises_sim.get(isbn, 0))
                ca_n  = ca_b - abs(ca_r) - abs(ca_re)

                base = ca_b if contrat["assiette"] == "CA brut HT" else max(ca_n, 0)

                mode_calc = ("Taux forfaitaire" if len(contrat["paliers"]) == 1
                             and contrat["paliers"][0]["seuil"] == 0 else "Paliers progressifs")
                droits_bruts_total = calcul_droits_paliers_sim(base, contrat["paliers"])
                droits_bruts_part  = droits_bruts_total * contrat["part"] / 100

                resultats_sim.append({
                    "Auteur":          contrat["auteur"],
                    "ISBN":            isbn,
                    "Titre":           contrat["titre"],
                    "Statut fiscal":   contrat.get("statut_fiscal", "Traitements et salaires (option assimilé)"),
                    "Mode de calcul":  mode_calc,
                    "Part auteur (%)": contrat["part"],
                    "Assiette":        contrat["assiette"],
                    "CA brut (€)":     round(ca_b, 2),
                    "Retours (€)":     round(abs(ca_r), 2),
                    "Remises (€)":     round(abs(ca_re), 2),
                    "Base calcul (€)": round(base, 2),
                    "Droits bruts (€)": round(droits_bruts_part, 2),
                })

            df_sim = pd.DataFrame(resultats_sim)

            if df_sim.empty:
                st.info("Aucune correspondance trouvée entre les ISBN du référentiel et ceux du SOCLE.")
            else:
                st.dataframe(
                    df_sim.style.format({c: (lambda x: fmt_fr(x, 2)) for c in
                                          ["CA brut (€)", "Retours (€)", "Remises (€)",
                                           "Base calcul (€)", "Droits bruts (€)"]}),
                    use_container_width=True
                )
                st.metric("💰 Total droits d'auteurs bruts estimés", f"{fmt_fr(df_sim['Droits bruts (€)'].sum(), 2)} €")

                # ── Étape 2 : précompte URSSAF + contribution diffuseur ──
                st.divider()
                st.markdown("**Étape 2 — Précompte URSSAF et contribution diffuseur**")
                st.caption(
                    "Contribution diffuseur due dans tous les cas. Précompte (CSG/CRDS, formation "
                    "professionnelle, RAAP) dû uniquement pour les auteurs en **traitements et salaires** "
                    "(statut fixé dans le référentiel) ; un auteur en **BNC** gère lui-même ses cotisations."
                )

                col_u1, col_u2, col_u3, col_u4 = st.columns(4)
                taux_csg_crds_sim = col_u1.number_input(
                    "CSG + CRDS (%)", value=9.70, step=0.01, key="sim_csg_crds",
                    help="CSG 9,2% + CRDS 0,5% — appliqué uniquement aux auteurs en traitements et salaires"
                )
                taux_fp_sim = col_u2.number_input(
                    "Formation professionnelle (%)", value=1.00, step=0.01, key="sim_fp",
                    help="Appliqué uniquement aux auteurs en traitements et salaires"
                )
                taux_raap_sim = col_u3.number_input(
                    "Retraite complémentaire RAAP (%)", value=0.0, step=0.01, key="sim_raap",
                    help="Variable selon revenus de l'auteur. Appliqué uniquement aux auteurs en traitements et salaires."
                )
                taux_diffuseur_sim = col_u4.number_input(
                    "Contribution diffuseur (%)", value=1.10, step=0.05, key="sim_diffuseur",
                    help="Due sur les droits bruts, quel que soit le statut fiscal de l'auteur"
                )

                est_ts_sim = df_sim["Statut fiscal"].str.startswith("Traitements")
                df_sim["Assiette URSSAF (€)"]  = df_sim["Droits bruts (€)"] * ASSIETTE_COEFF
                df_sim["Précompte URSSAF (€)"] = np.where(
                    est_ts_sim,
                    df_sim["Assiette URSSAF (€)"] * taux_csg_crds_sim / 100
                    + df_sim["Droits bruts (€)"] * taux_fp_sim / 100
                    + df_sim["Assiette URSSAF (€)"] * taux_raap_sim / 100,
                    0.0
                )
                df_sim["Contribution diffuseur (€)"] = df_sim["Droits bruts (€)"] * taux_diffuseur_sim / 100
                df_sim["Total dû à l'URSSAF (€)"]    = df_sim["Précompte URSSAF (€)"] + df_sim["Contribution diffuseur (€)"]
                df_sim["Net à payer auteur (€)"]     = df_sim["Droits bruts (€)"] - df_sim["Précompte URSSAF (€)"]

                cols_sim_urssaf = ["Auteur", "Titre", "Statut fiscal", "Droits bruts (€)",
                                    "Précompte URSSAF (€)", "Contribution diffuseur (€)",
                                    "Total dû à l'URSSAF (€)", "Net à payer auteur (€)"]
                st.dataframe(
                    df_sim[cols_sim_urssaf].style.format({c: (lambda x: fmt_fr(x, 2)) for c in cols_sim_urssaf if "(€)" in c}),
                    use_container_width=True
                )

                total_precompte_sim = df_sim["Précompte URSSAF (€)"].sum()
                total_diffuseur_sim = df_sim["Contribution diffuseur (€)"].sum()
                total_urssaf_sim    = df_sim["Total dû à l'URSSAF (€)"].sum()
                col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                col_m1.metric("Droits bruts totaux", f"{fmt_fr(df_sim['Droits bruts (€)'].sum(), 2)} €")
                col_m2.metric("Précompte URSSAF (TS uniquement)", f"{fmt_fr(total_precompte_sim, 2)} €")
                col_m3.metric("Contribution diffuseur (tous)", f"{fmt_fr(total_diffuseur_sim, 2)} €")
                col_m4.metric("Total estimé dû à l'URSSAF", f"{fmt_fr(total_urssaf_sim, 2)} €")

    # ══════════════════════════════════════════════
    # ONGLET 3 — RÉEL (COMPTABILISÉ)
    # ══════════════════════════════════════════════
    with onglet3:
        st.subheader("Montants réellement comptabilisés — titre par titre, puis auteur par auteur")
        st.caption("Ce module ne recalcule rien : il lit directement, pour chaque ISBN, les écritures déjà "
                   "provisionnées dans votre grand livre analytique (droits bruts, contribution diffuseur, "
                   "précompte URSSAF, net dû à l'auteur), pour ne jamais diverger de la comptabilité réelle.")

        if "df_pivot" not in st.session_state:
            st.warning("⚠️ Générer d'abord le SOCLE EDITION.")
        else:
            df_pivot_reel = st.session_state["df_pivot"].copy()
            df_pivot_reel["Date"] = pd.to_datetime(df_pivot_reel["Date"], errors="coerce")

            st.markdown("**Période de déclaration**")
            dates_valides = df_pivot_reel["Date"].dropna()
            date_min_defaut = dates_valides.min() if not dates_valides.empty else pd.Timestamp("2024-01-01")
            date_max_defaut = dates_valides.max() if not dates_valides.empty else pd.Timestamp.today()
            col_p1, col_p2 = st.columns(2)
            periode_debut = col_p1.date_input("Du", value=date_min_defaut.date(), key="reel_periode_debut")
            periode_fin   = col_p2.date_input("Au", value=date_max_defaut.date(), key="reel_periode_fin")
            st.caption("Les montants ci-dessous ne portent que sur les écritures dont la date est comprise "
                       "dans cette période (ex. un trimestre ou un mois de déclaration URSSAF).")

            mask_periode = (
                (df_pivot_reel["Date"] >= pd.to_datetime(periode_debut))
                & (df_pivot_reel["Date"] <= pd.to_datetime(periode_fin))
            )
            df_pivot_reel = df_pivot_reel[mask_periode]

            st.markdown("**Comptes à lire dans le grand livre** _(réutilisables d'un exercice à l'autre)_")
            col_c1, col_c2, col_c3, col_c4 = st.columns(4)
            compte_droits_bruts = col_c1.text_input("Droits bruts (charge)", value="604300000", key="cpt_droits_bruts")
            compte_diffuseur    = col_c2.text_input("Contribution diffuseur", value="645106", key="cpt_diffuseur")
            compte_urssaf       = col_c3.text_input("URSSAF à payer", value="438106", key="cpt_urssaf")
            compte_net_du       = col_c4.text_input("Droits d'auteurs à payer (net)", value="408106", key="cpt_net_du")
            st.caption("⚠️ Le compte de droits bruts peut aussi contenir d'autres prestations (achats non liés "
                       "aux droits d'auteurs) selon votre plan de comptes — vérifiez le détail en cas d'écart "
                       "inattendu avec vos attentes.")

            def _par_isbn_net(compte, sens):
                """Solde NET par ISBN (Débit − Crédit ou Crédit − Débit selon `sens`), et non un
                seul sens sommé isolément : une écriture de reclassement/régularisation (ex. une
                correction OD qui débite ce compte pour en réduire le solde) doit bien réduire le
                montant retenu, plutôt que d'être ignorée parce qu'elle porte sur le "mauvais" sens."""
                m = df_pivot_reel["Compte"].astype(str).str.strip() == str(compte).strip()
                if not m.any():
                    return pd.Series(dtype=float)
                g = df_pivot_reel[m].groupby("Code_Analytique")
                if sens == "debit":
                    return g["Débit"].sum() - g["Crédit"].sum()
                return g["Crédit"].sum() - g["Débit"].sum()

            droits_bruts_s = _par_isbn_net(compte_droits_bruts, "debit")
            diffuseur_s    = _par_isbn_net(compte_diffuseur, "debit")
            urssaf_s       = _par_isbn_net(compte_urssaf, "credit")
            net_du_s       = _par_isbn_net(compte_net_du, "credit")

            isbns_reels = sorted(set(droits_bruts_s.index) | set(diffuseur_s.index)
                                | set(urssaf_s.index) | set(net_du_s.index))
            isbns_reels = [i for i in isbns_reels
                           if str(i).strip() not in ("", "CHARGES INDIRECTES", "PRODUITS INDIRECTS")]

            if not isbns_reels:
                st.info("Aucun montant trouvé sur ces comptes pour l'instant. Vérifiez les numéros de comptes "
                        "ci-dessus, ou provisionnez d'abord les droits d'auteurs dans votre grand livre.")
            else:
                mapping_auto = resoudre_mapping_auteurs(df_pivot_reel)
                referentiel_par_isbn = {}
                for c in st.session_state["royalties_referentiel"]:
                    referentiel_par_isbn.setdefault(c["isbn"], []).append((c["auteur"], c["part"], c["titre"]))

                source_mapping = ("détection automatique (famille analytique AUTEUR)" if mapping_auto
                                   else "référentiel manuel (onglet Référentiel)")
                st.caption(f"Correspondance ISBN → auteur utilisée : {source_mapping}.")

                lignes, isbn_sans_auteur = [], []
                for isbn in isbns_reels:
                    droits_bruts = float(droits_bruts_s.get(isbn, 0))
                    diffuseur    = float(diffuseur_s.get(isbn, 0))
                    urssaf_total = float(urssaf_s.get(isbn, 0))
                    net_du       = float(net_du_s.get(isbn, 0))

                    if isbn in referentiel_par_isbn:
                        parts = referentiel_par_isbn[isbn]
                    elif isbn in mapping_auto:
                        parts = [(mapping_auto[isbn], 100.0, isbn)]
                    else:
                        parts = [("(auteur non identifié)", 100.0, isbn)]
                        isbn_sans_auteur.append(isbn)

                    for auteur, part, titre in parts:
                        coeff = part / 100.0
                        lignes.append({
                            "ISBN": isbn,
                            "Titre": titre if titre and titre != isbn else isbn,
                            "Auteur": auteur,
                            "Statut fiscal": obtenir_statut_fiscal(auteur),
                            "Part (%)": part,
                            "Droits bruts (€)": round(droits_bruts * coeff, 2),
                            "Contribution diffuseur (€)": round(diffuseur * coeff, 2),
                            "Précompte URSSAF (€)": round(urssaf_total * coeff, 2),
                            "Net du a l'auteur (€)": round(net_du * coeff, 2),
                        })

                df_reel = pd.DataFrame(lignes)

                st.session_state["df_royalties_reel"] = df_reel

                if isbn_sans_auteur:
                    st.warning(f"⚠️ {len(isbn_sans_auteur)} ISBN avec des droits comptabilisés mais sans auteur "
                               f"identifié : {', '.join(isbn_sans_auteur)}. Ajoutez-les dans l'onglet "
                               f"Référentiel pour qu'ils apparaissent nommément dans les relevés.")

                cols_montant = ["Droits bruts (€)", "Contribution diffuseur (€)",
                                "Précompte URSSAF (€)", "Net du a l'auteur (€)"]
                st.dataframe(
                    df_reel.style.format({c: (lambda x: fmt_fr(x, 2)) for c in cols_montant}),
                    use_container_width=True
                )

                total_droits_bruts = df_reel["Droits bruts (€)"].sum()
                total_diffuseur    = df_reel["Contribution diffuseur (€)"].sum()
                total_urssaf       = df_reel["Précompte URSSAF (€)"].sum()
                total_net_du       = df_reel["Net du a l'auteur (€)"].sum()
                col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                col_m1.metric("Droits bruts comptabilisés", f"{fmt_fr(total_droits_bruts, 2)} €")
                col_m2.metric("Précompte URSSAF (TS)", f"{fmt_fr(total_urssaf, 2)} €")
                col_m3.metric("Contribution diffuseur (tous)", f"{fmt_fr(total_diffuseur, 2)} €")
                col_m4.metric("Net dû aux auteurs", f"{fmt_fr(total_net_du, 2)} €")

                df_par_auteur_reel = df_reel.groupby("Auteur", as_index=False)[cols_montant].sum()
                fig_reel = px.bar(
                    df_par_auteur_reel.sort_values("Net du a l'auteur (€)", ascending=False),
                    x="Auteur", y="Net du a l'auteur (€)", text_auto=".0f",
                    title="Net dû par auteur (comptabilisé)"
                )
                fig_reel.update_traces(textposition="outside")
                st.plotly_chart(fig_reel, use_container_width=True)

                # ================================================
                # RÉCAPITULATIF — droits dus par auteur + déclaration URSSAF de la période
                # ================================================
                st.divider()
                st.subheader(f"📋 Récapitulatif de la période du {periode_debut.strftime('%d/%m/%Y')} "
                             f"au {periode_fin.strftime('%d/%m/%Y')}")

                st.markdown("**Droits dus par auteur** _(net après précompte, tel que comptabilisé)_")
                st.dataframe(
                    df_par_auteur_reel.sort_values("Net du a l'auteur (€)", ascending=False)
                                      .style.format({c: (lambda x: fmt_fr(x, 2)) for c in cols_montant}),
                    use_container_width=True
                )

                st.markdown(f"""
                <div style='padding:14px 18px; border-radius:12px; background:#eef2ff;
                            margin-top:8px; margin-bottom:8px'>
                    <div style='font-weight:600; font-size:15px; margin-bottom:6px'>
                        🏛️ Déclaration URSSAF à effectuer sur cette période
                    </div>
                    <div>Précompte URSSAF à reverser (auteurs en traitements et salaires uniquement) :
                         <b>{fmt_fr(total_urssaf, 2)} €</b></div>
                    <div>Contribution diffuseur à reverser (tous auteurs) :
                         <b>{fmt_fr(total_diffuseur, 2)} €</b></div>
                    <div>Total à reverser à l'URSSAF : <b>{fmt_fr(round(total_urssaf + total_diffuseur, 2), 2)} €</b></div>
                    <div>Assis sur des droits bruts comptabilisés de : <b>{fmt_fr(total_droits_bruts, 2)} €</b></div>
                </div>
                """, unsafe_allow_html=True)

                buffer_reel = BytesIO()
                with pd.ExcelWriter(buffer_reel, engine="openpyxl") as writer:
                    df_reel.to_excel(writer, index=False, sheet_name="Detail_par_titre")
                    df_par_auteur_reel.to_excel(writer, index=False, sheet_name="Recap_par_auteur")
                    pd.DataFrame([{
                        "Période début": periode_debut, "Période fin": periode_fin,
                        "Droits bruts (€)": total_droits_bruts,
                        "Précompte URSSAF à déclarer, TS uniquement (€)": total_urssaf,
                        "Contribution diffuseur à déclarer, tous auteurs (€)": total_diffuseur,
                        "Total à reverser à l'URSSAF (€)": round(total_urssaf + total_diffuseur, 2),
                        "Net dû aux auteurs (€)": total_net_du,
                    }]).to_excel(writer, index=False, sheet_name="Declaration_URSSAF")
                buffer_reel.seek(0)
                st.download_button(
                    "📥 Exporter le récapitulatif de la période (Excel)",
                    buffer_reel,
                    file_name=f"Droits_auteurs_{periode_debut}_{periode_fin}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="export_reel_periode"
                )

                # ================================================
                # INDICATEUR 4 — SOUS-INDICATEURS DE PILOTAGE (à-valoir, couverture, prévisionnel)
                # ================================================
                st.divider()
                with st.expander("📐 Sous-indicateurs de pilotage (à-valoir, couverture, prévisionnel 6 mois)"):
                    st.caption("Calculés sur l'historique complet du grand livre (et non la seule période de "
                               "déclaration ci-dessus), car l'amortissement d'un à-valoir est un suivi cumulatif.")
                    df_pivot_complet = st.session_state["df_pivot"].copy()
                    df_pivot_complet["Date"] = pd.to_datetime(df_pivot_complet["Date"], errors="coerce")

                    col_i4a, col_i4b = st.columns(2)
                    cpt_avance4 = col_i4a.text_input("Compte avances / à-valoirs", value="409600", key="ind4_cpt_avance")
                    taux_contractuel4 = col_i4b.number_input("Taux contractuel par défaut (%)", value=10.0, step=0.5, key="ind4_taux") / 100

                    avance_debit_isbn = df_pivot_complet[df_pivot_complet["Compte"].astype(str) == cpt_avance4].groupby("Code_Analytique")["Débit"].sum()
                    avance_credit_isbn = df_pivot_complet[df_pivot_complet["Compte"].astype(str) == cpt_avance4].groupby("Code_Analytique")["Crédit"].sum()
                    premiere_date_avance = df_pivot_complet[
                        (df_pivot_complet["Compte"].astype(str) == cpt_avance4) & (df_pivot_complet["Débit"] > 0)
                    ].groupby("Code_Analytique")["Date"].min()
                    droits_bruts_complet = df_pivot_complet[df_pivot_complet["Compte"].astype(str) == compte_droits_bruts].groupby("Code_Analytique")["Débit"].sum()

                    isbns_avalent = sorted(set(avance_debit_isbn.index) | set(droits_bruts_complet.index))
                    isbns_avalent = [i for i in isbns_avalent if str(i).strip() not in ("", "CHARGES INDIRECTES", "PRODUITS INDIRECTS")]

                    if isbns_avalent:
                        indic4 = pd.DataFrame({"Code_Analytique": isbns_avalent}).set_index("Code_Analytique")
                        indic4["Droits comptabilisés (réel, historique)"] = droits_bruts_complet.reindex(indic4.index, fill_value=0.0)
                        indic4["À-valoir initial (cumul débit)"] = avance_debit_isbn.reindex(indic4.index, fill_value=0.0)
                        indic4["Droits cumulés versés (cumul crédit)"] = avance_credit_isbn.reindex(indic4.index, fill_value=0.0)
                        indic4["À-valoir restant à amortir"] = indic4["À-valoir initial (cumul débit)"] - indic4["Droits cumulés versés (cumul crédit)"]
                        indic4["Taux de couverture (%)"] = np.where(
                            indic4["À-valoir initial (cumul débit)"] > 0,
                            indic4["Droits cumulés versés (cumul crédit)"] / indic4["À-valoir initial (cumul débit)"] * 100, 100.0
                        )
                        mois_ecoules = premiere_date_avance.reindex(indic4.index).apply(
                            lambda d: (pd.Timestamp.today() - d).days / 30.44 if pd.notna(d) else np.nan
                        )

                        def _alerte_avalent(isbn):
                            restant = indic4.loc[isbn, "À-valoir restant à amortir"]
                            couverture = indic4.loc[isbn, "Taux de couverture (%)"]
                            ecoule = mois_ecoules.get(isbn, np.nan)
                            if pd.notna(ecoule) and restant > 0 and ecoule > 12:
                                return "🔴 À-valoir non amorti > 12 mois"
                            if pd.notna(ecoule) and couverture < 50 and ecoule > 6:
                                return "🟠 Couverture < 50 % après 6 mois"
                            return "🟢 Normal"
                        indic4["Alerte"] = [_alerte_avalent(i) for i in indic4.index]

                        df_pivot_complet["Mois"] = df_pivot_complet["Date"].dt.to_period("M")
                        v_i4 = df_pivot_complet[df_pivot_complet["Compte"].astype(str).str.startswith("701")].groupby("Mois")["Crédit"].sum()
                        r_i4 = df_pivot_complet[df_pivot_complet["Compte"].astype(str).str.startswith("709")].groupby("Mois")["Débit"].sum()
                        ca_net_i4 = v_i4.sub(r_i4, fill_value=0).sort_index()
                        ca_net_total4 = ca_net_i4.sum()

                        ia1, ia2, ia3 = st.columns(3)
                        ia1.metric("Droits bruts comptabilisés (historique)", f"{fmt_fr(indic4['Droits comptabilisés (réel, historique)'].sum(), 0)} €")
                        ia2.metric("À-valoir restant à amortir (total)", f"{fmt_fr(indic4['À-valoir restant à amortir'].sum(), 0)} €")
                        taux_moyen_droits_ca4 = (indic4["Droits comptabilisés (réel, historique)"].sum() / ca_net_total4 * 100) if ca_net_total4 else 0.0
                        ia3.metric("Taux moyen droits d'auteurs / CA net", f"{taux_moyen_droits_ca4:.1f} %" if ca_net_total4 else "N/A")

                        formats_i4 = {c: (lambda x: f"{fmt_fr(x, 0)} €") for c in ["Droits comptabilisés (réel, historique)", "À-valoir initial (cumul débit)",
                                                                "Droits cumulés versés (cumul crédit)", "À-valoir restant à amortir"]}
                        formats_i4["Taux de couverture (%)"] = "{:.1f} %"
                        st.dataframe(
                            indic4.reset_index().rename(columns={"index": "Code_Analytique"}).style.format(formats_i4),
                            use_container_width=True, hide_index=True
                        )

                        base_i4 = ca_net_i4.iloc[-3:].mean() if len(ca_net_i4) >= 3 else (ca_net_i4.mean() if len(ca_net_i4) else 0.0)
                        droits_prev_6mois = base_i4 * 6 * taux_contractuel4
                        st.metric("Droits prévisionnels sur 6 mois (estimation, taux contractuel appliqué au CA net moyen)",
                                  f"{fmt_fr(droits_prev_6mois, 0)} €")

                        echeance_semestrielle = pd.Timestamp.today() >= pd.Timestamp("2027-12-20")
                        st.caption(
                            "**Échéances :** déclaration trimestrielle URSSAF + reddition " +
                            ("**semestrielle** (accord interprofessionnel CPE-LAP-SNE du 20/12/2022, applicable depuis le 20/12/2027)"
                             if echeance_semestrielle else
                             "**annuelle** (passera à semestrielle à compter du 20/12/2027, conformément à l'accord "
                             "interprofessionnel CPE-LAP-SNE du 20/12/2022)")
                        )
                    else:
                        st.info("Aucune écriture trouvée sur le compte d'avances/à-valoirs indiqué.")

    # ══════════════════════════════════════════════
    # ONGLET 4 — RELEVÉS PAR AUTEUR
    # ══════════════════════════════════════════════
    with onglet4:
        st.subheader("Relevés de droits par auteur")

        if "df_royalties_reel" not in st.session_state or st.session_state["df_royalties_reel"].empty:
            st.warning("⚠️ Effectuez d'abord la lecture dans l'onglet 'Réel (comptabilisé)'.")
        else:
            st.caption("✅ Source : montants réellement comptabilisés (onglet 📒 Réel).")
            df_releves = st.session_state["df_royalties_reel"].copy()
            auteurs    = sorted(df_releves["Auteur"].unique().tolist())

            auteur_sel = st.selectbox("Sélectionnez un auteur", ["Tous"] + auteurs)

            if auteur_sel != "Tous":
                df_auteur = df_releves[df_releves["Auteur"] == auteur_sel]
            else:
                df_auteur = df_releves

            cols_releve_base = [
                "Auteur", "Statut fiscal", "Titre", "ISBN", "Part (%)",
                "Droits bruts (€)", "Contribution diffuseur (€)", "Précompte URSSAF (€)",
                "Net du a l'auteur (€)"
            ]
            cols_dispo = [c for c in cols_releve_base if c in df_auteur.columns]

            st.dataframe(
                df_auteur[cols_dispo].style.format({
                    c: (lambda x: fmt_fr(x, 2)) for c in cols_dispo if "(€)" in c
                }),
                use_container_width=True
            )

            # Résumé par auteur (si "Tous")
            if auteur_sel == "Tous":
                st.subheader("Synthèse par auteur")
                cols_sum = {c: "sum" for c in cols_dispo if "(€)" in c}
                df_synth = df_releves.groupby("Auteur", as_index=False).agg(cols_sum)
                st.dataframe(
                    df_synth.style.format({c: (lambda x: fmt_fr(x, 2)) for c in cols_sum}),
                    use_container_width=True
                )

            # ── Export Excel multi-feuilles ──
            buffer_xl = BytesIO()
            with pd.ExcelWriter(buffer_xl, engine="openpyxl") as writer:
                # Feuille 1 : détail complet
                df_auteur[cols_dispo].to_excel(
                    writer, index=False,
                    sheet_name=f"Détail_{auteur_sel[:20]}"
                )
                # Feuille 2 : synthèse par auteur
                if auteur_sel == "Tous" and "df_synth" in dir():
                    df_synth.to_excel(writer, index=False, sheet_name="Synthèse auteurs")
                # Feuille 3 : référentiel contrats
                if st.session_state["royalties_referentiel"]:
                    rows_ref = []
                    for c in st.session_state["royalties_referentiel"]:
                        row = {
                            "auteur": c["auteur"], "isbn": c["isbn"],
                            "titre": c["titre"], "part": c["part"],
                            "assiette": c["assiette"]
                        }
                        for i, p in enumerate(c["paliers"], 1):
                            row[f"seuil_{i}"] = p["seuil"]
                            row[f"taux_{i}"]  = p["taux"]
                        rows_ref.append(row)
                    pd.DataFrame(rows_ref).to_excel(writer, index=False, sheet_name="Référentiel contrats")

            buffer_xl.seek(0)
            st.download_button(
                label=f"📥 Télécharger le relevé — {auteur_sel}",
                data=buffer_xl,
                file_name=f"Royalties_{auteur_sel.replace(' ','_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# =====================
# RETOURS & REMISES
# =====================
elif page == "📦 Retours & Remises":
    st.header("📦 Retours & Remises")
    if "df_pivot" not in st.session_state:
        st.warning("⚠️ Générez d'abord le socle analytique.")
        st.stop()

    df = st.session_state["df_pivot"].copy()
    param = st.session_state.get("param_comptes", {})
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Mois"] = df["Date"].dt.strftime("%Y-%m")

    seuil_alerte = st.sidebar.number_input("Seuil alerte taux de retour (%)", value=25, step=5)

    def filtre(df_src, prefix_list, exclude_prefix_list=None):
        if not prefix_list: return pd.DataFrame()
        mask = df_src["Compte"].astype(str).str.startswith(tuple(prefix_list))
        if exclude_prefix_list:
            mask = mask & (~df_src["Compte"].astype(str).str.startswith(tuple(exclude_prefix_list)))
        f = df_src[mask].copy()
        if not f.empty: f["Montant_net"] = f["Débit"] - f["Crédit"]
        return f

    # Exclusion des comptes remises du filtre retours (cf. mask_retours) : évite un double
    # comptage quand le compte remises (ex. 7091) est un sous-compte du compte retours (709).
    df_ret = filtre(df, param.get("retours", []), exclude_prefix_list=param.get("remises"))
    df_rem = filtre(df, param.get("remises", []))
    df_v   = filtre(df, param.get("ventes", []))
    # CA distributeur (compte configurable, ex. 7011 = BLDD pour ce cas d'étude) : base STRICTE
    # des taux de retour/remise, distincte du CA brut large affiché ci-dessous (params["ventes"],
    # ex. 701 = 7010 + 7011 + ...) — cf. Tableau de bord éditorial et Synthèse financière, mêmes
    # définitions.
    df_v_distrib_rr = filtre(df, param.get("ventes_distributeur") or param.get("ventes", []))

    total_retours = abs(df_ret["Montant_net"].sum()) if not df_ret.empty else 0
    total_remises = abs(df_rem["Montant_net"].sum()) if not df_rem.empty else 0
    total_ventes  = df_v["Crédit"].sum() if not df_v.empty else 0
    total_ventes_distrib = df_v_distrib_rr["Crédit"].sum() if not df_v_distrib_rr.empty else 0
    taux_ret = (total_retours / total_ventes_distrib * 100) if total_ventes_distrib else 0
    taux_rem = (total_remises / total_ventes_distrib * 100) if total_ventes_distrib else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("CA brut", f"{fmt_fr(total_ventes, 0)} €")
    c2.metric("Total retours", f"{fmt_fr(total_retours, 0)} €")
    c3.metric("Taux de retour", f"{taux_ret:.1f} %",
              delta="⚠️ Dépasse le seuil !" if taux_ret > seuil_alerte else "✅ Normal",
              delta_color="inverse" if taux_ret > seuil_alerte else "normal")
    c4.metric("Taux de remise", f"{taux_rem:.1f} %")

    if taux_ret > seuil_alerte:
        st.error(f"🚨 Alerte : le taux de retour ({taux_ret:.1f}%) dépasse votre seuil de {seuil_alerte}% !")

    # Comparaison retours réels vs provision 681
    df_prov = df[df["Compte"].astype(str).str.startswith("681")]
    provision = (df_prov["Débit"] - df_prov["Crédit"]).sum() if not df_prov.empty else 0
    ecart = total_retours - provision
    st.info(f"📋 Provision retours comptabilisée (681) : **{fmt_fr(provision, 0)} €** — "
            f"Écart avec retours réels : **{fmt_fr(ecart, 0)} €** "
            f"({'sous-provision' if ecart > 0 else 'sur-provision'})")

    col_g1, col_g2 = st.columns(2)
    if not df_ret.empty:
        with col_g1:
            trend_ret = df_ret.groupby("Mois")["Montant_net"].sum().abs().reset_index()
            fig1 = px.bar(trend_ret, x="Mois", y="Montant_net", title="Retours mensuels (€)",
                           text="Montant_net", labels={"Montant_net": "€"}, color_discrete_sequence=["#EF4444"])
            fig1.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
            fig1.update_layout(separators=", ")  # format FR : virgule decimale, espace milliers
            fig1.update_layout(height=320, margin=dict(t=40))
            st.plotly_chart(fig1, use_container_width=True)
    if not df_rem.empty:
        with col_g2:
            trend_rem = df_rem.groupby("Mois")["Montant_net"].sum().abs().reset_index()
            fig2 = px.bar(trend_rem, x="Mois", y="Montant_net", title="Remises mensuelles (€)",
                           text="Montant_net", labels={"Montant_net": "€"}, color_discrete_sequence=["#F59E0B"])
            fig2.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
            fig2.update_layout(separators=", ")  # format FR : virgule decimale, espace milliers
            fig2.update_layout(height=320, margin=dict(t=40))
            st.plotly_chart(fig2, use_container_width=True)

    if not df_ret.empty:
        # Même filtre "vraie ligne ISBN" que les autres modules, pour ne pas laisser une
        # ligne "CHARGES INDIRECTES" ou un code vide polluer le classement des retours par titre.
        df_ret_isbn = filtrer_isbn_reels(df_ret)
        if not df_ret_isbn.empty:
            ret_isbn = df_ret_isbn.groupby("Code_Analytique")["Montant_net"].sum().abs().reset_index()
            ret_isbn = ret_isbn.sort_values("Montant_net", ascending=False)
            st.subheader("Retours par titre")
            st.dataframe(ret_isbn.style.format({"Montant_net": (lambda x: f"{fmt_fr(x, 0)} €")}), hide_index=True)

    # ================================================
    # INDICATEUR 3 — SOUS-INDICATEURS DE PILOTAGE
    # ================================================
    st.divider()
    with st.expander("📐 Sous-indicateurs de pilotage (provision, évolution, collection, méthode Parly)"):
        TAUX_TVA_LIVRE = 0.055  # taux réduit applicable au livre en France
        provision_retours = 0.10 * total_ventes * (1 + TAUX_TVA_LIVRE)
        ventes_nettes_globales = total_ventes - total_retours - provision_retours
        if taux_ret > 30:
            alerte_ref3 = "🔴 Mise en place excessive"
        elif taux_ret >= 20:
            alerte_ref3 = "🟠 À surveiller"
        else:
            alerte_ref3 = "🟢 Cible sectorielle atteinte"
        r1, r2, r3 = st.columns(3)
        r1.metric("Taux de retour (seuils référentiel 20 %/30 %)", f"{taux_ret:.1f} %", delta=alerte_ref3)
        r2.metric("Provision pour retours futurs (10 % ventes brutes TTC)", f"{fmt_fr(provision_retours, 0)} €")
        r3.metric("Ventes nettes (après retours + provision)", f"{fmt_fr(ventes_nettes_globales, 0)} €")

        idx_collection = next((i for i, n in enumerate(st.session_state.get("noms_familles_actives", []))
                                if "collection" in n.lower()), None)
        if idx_collection is not None:
            codes_cols_ret = st.session_state.get("codes_cols", [])
            col_collection = codes_cols_ret[idx_collection] if idx_collection < len(codes_cols_ret) else None
            if col_collection and col_collection in df.columns and not df_v.empty:
                st.markdown("**Taux de retour par collection**")
                v_coll = df_v.groupby(col_collection)["Crédit"].sum()
                r_coll = df_ret.groupby(col_collection)["Montant_net"].sum().abs() if not df_ret.empty else pd.Series(dtype=float)
                taux_coll = (r_coll / v_coll * 100).fillna(0).rename("Taux de retour (%)")
                st.dataframe(taux_coll.to_frame().style.format("{:.1f} %"), use_container_width=True)
        else:
            st.caption("ℹ️ Aucune famille analytique « Collection » mappée — mappable dans ⚙️ Paramétrage "
                       "analytique si votre export le permet.")

        st.markdown("**Évolution du taux de retour — M vs M-1 vs M-12**")
        v_m3 = df_v.groupby("Mois")["Crédit"].sum() if not df_v.empty else pd.Series(dtype=float)
        r_m3 = df_ret.groupby("Mois")["Montant_net"].sum().abs() if not df_ret.empty else pd.Series(dtype=float)
        taux_m3 = (r_m3 / v_m3 * 100).dropna().sort_index()
        if len(taux_m3) >= 2:
            mois_index_dt = pd.PeriodIndex(taux_m3.index, freq="M")
            taux_m3.index = mois_index_dt
            dernier_m3 = mois_index_dt[-1]
            m_moins_1_3 = dernier_m3 - 1
            m_moins_12_3 = dernier_m3 - 12
            rc1, rc2, rc3 = st.columns(3)
            rc1.metric(f"M ({dernier_m3})", f"{taux_m3.iloc[-1]:.1f} %")
            rc2.metric(f"M-1 ({m_moins_1_3})", f"{taux_m3.get(m_moins_1_3):.1f} %" if m_moins_1_3 in taux_m3.index else "N/A")
            rc3.metric(f"M-12 ({m_moins_12_3})", f"{taux_m3.get(m_moins_12_3):.1f} %" if m_moins_12_3 in taux_m3.index else "N/A")
            st.line_chart(taux_m3)
        else:
            st.caption("Historique insuffisant pour comparer M / M-1 / M-12.")

        st.markdown("**Référentiel complémentaire (optionnel)** — pour le coût financier des retours et le "
                    "ratio ventes nettes / tirage (méthode Parly). Laissez à 0 si non disponible.")
        titres_ret = sorted(filtrer_isbn_reels(df)["Code_Analytique"].astype(str).unique().tolist())
        if "referentiel_titres_manuel" not in st.session_state or set(st.session_state["referentiel_titres_manuel"]["ISBN"]) != set(titres_ret):
            st.session_state["referentiel_titres_manuel"] = pd.DataFrame({
                "ISBN": titres_ret, "PPHT (€)": [0.0] * len(titres_ret),
                "Coût fabrication unitaire (€)": [0.0] * len(titres_ret), "Tirage initial (ex.)": [0] * len(titres_ret),
            })
        ref_titres_ret = st.data_editor(st.session_state["referentiel_titres_manuel"], use_container_width=True,
                                        num_rows="fixed", key="editor_ref_titres_retours", hide_index=True)
        st.session_state["referentiel_titres_manuel"] = ref_titres_ret

        if titres_ret and not df_ret_isbn.empty:
            params_ret = st.session_state.get("param_comptes", {})
            indic_ret = calculer_indicateurs_titres(df, params_ret, titres_ret)
            ref_idx_ret = ref_titres_ret.set_index("ISBN")
            indic3_ret = indic_ret[["Code_Analytique", "Ventes HT", "Retours"]].set_index("Code_Analytique").join(ref_idx_ret, how="left").fillna(0)
            indic3_ret["Coût financier des retours (€)"] = np.where(
                indic3_ret["PPHT (€)"] > 0,
                indic3_ret["Retours"] * (1 - indic3_ret["Coût fabrication unitaire (€)"] / indic3_ret["PPHT (€)"]),
                0.0
            )
            indic3_ret["Nb exemplaires vendus nets (est.)"] = np.where(
                indic3_ret["PPHT (€)"] > 0, (indic3_ret["Ventes HT"] - indic3_ret["Retours"]) / indic3_ret["PPHT (€)"], 0.0
            )
            indic3_ret["Ratio ventes nettes / tirage (Parly)"] = np.where(
                indic3_ret["Tirage initial (ex.)"] > 0,
                indic3_ret["Nb exemplaires vendus nets (est.)"] / indic3_ret["Tirage initial (ex.)"], 0.0
            )
            indic3_ret["Alerte Parly"] = np.where(
                (indic3_ret["Tirage initial (ex.)"] > 0) & (indic3_ret["Ratio ventes nettes / tirage (Parly)"] < 0.65),
                "🔴 Sous le seuil critique (0,65)", "—"
            )
            st.markdown("**Coût financier des retours et ratio Parly par titre**")
            aff3_ret = indic3_ret[["Retours", "Coût financier des retours (€)", "Nb exemplaires vendus nets (est.)",
                                   "Tirage initial (ex.)", "Ratio ventes nettes / tirage (Parly)", "Alerte Parly"]]
            st.dataframe(aff3_ret.style.format({"Retours": (lambda x: f"{fmt_fr(x, 0)} €"), "Coût financier des retours (€)": (lambda x: f"{fmt_fr(x, 0)} €"),
                                                "Nb exemplaires vendus nets (est.)": (lambda x: fmt_fr(x, 0)),
                                                "Ratio ventes nettes / tirage (Parly)": "{:.2f}"}),
                        use_container_width=True)
            st.caption("Le nombre d'exemplaires retournés/vendus est estimé en divisant les montants € par le "
                      "PPHT saisi ci-dessus (la comptabilité ne porte que des montants, pas des quantités).")
# =====================
# SYNTHÈSE FINANCIÈRE
# =====================
elif page == "📊 Synthèse financière":
    st.header("📊 Synthèse financière globale")
    if "df_pivot" not in st.session_state:
        st.warning("⚠️ Générez d'abord le socle analytique.")
        st.stop()

    df = st.session_state["df_pivot"].copy()
    params = st.session_state["param_comptes"]

    def filtre_m(df_src, prefix_list, exclude_prefix_list=None):
        if not prefix_list: return pd.DataFrame()
        mask = df_src["Compte"].astype(str).str.startswith(tuple(prefix_list))
        if exclude_prefix_list:
            mask = mask & (~df_src["Compte"].astype(str).str.startswith(tuple(exclude_prefix_list)))
        f = df_src[mask].copy()
        if not f.empty: f["Montant_net"] = f["Débit"] - f["Crédit"]
        return f

    def filtre_mask(df_src, mask):
        f = df_src[mask].copy()
        if not f.empty: f["Montant_net"] = f["Débit"] - f["Crédit"]
        return f

    # mask_ventes/mask_charges incluent aussi les quote-parts de produits/charges indirects
    # répartis (comptes non numériques "PRODUITS INDIRECTS REPARTIS"/"CHARGES INDIRECTES
    # REPARTIES"), sans quoi activer la répartition ferait disparaître ces montants des
    # totaux globaux ci-dessous (ils ne matchent plus le préfixe numérique "701"/"6").
    # CA brut/net inclut désormais tout produit (compte 7xx) qui n'est ni un retour ni une
    # remise — commissions, subventions, reprises, produits divers... (cf. mask_ventes) : ces
    # montants ne sont pas isolés dans une ligne séparée, ils sont directement dans le CA.
    df_v   = filtre_mask(df, mask_ventes(df, params))
    # Exclusion des comptes remises du filtre retours (cf. mask_retours) : évite un double
    # comptage quand le compte remises (ex. 7091) est un sous-compte du compte retours (709).
    df_r   = filtre_m(df, params["retours"], exclude_prefix_list=params.get("remises"))
    df_rem = filtre_m(df, params["remises"])
    df_c   = filtre_mask(df, mask_charges(df, params))
    # CA distributeur (compte configurable, ex. 7011 = BLDD pour ce cas d'étude) : base STRICTE
    # de calcul du taux de retour/remise — distincte du périmètre "ventes" large (params["ventes"],
    # ex. 701 = 7010 + 7011 + ...) utilisé juste en dessous pour les extournes, et du CA brut
    # élargi (commissions/subventions/produits divers). Le relevé du distributeur ne couvre que
    # ce canal de vente précis : diviser par un périmètre plus large dilue artificiellement ces taux.
    prefixes_ca_distrib = tuple(params.get("ventes_distributeur") or params["ventes"])
    df_ca_distrib = df[df["Compte"].astype(str).str.startswith(prefixes_ca_distrib)]
    ca_brut_distrib = df_ca_distrib["Crédit"].sum() if not df_ca_distrib.empty else 0

    df_v_large = df[df["Compte"].astype(str).str.startswith(tuple(params["ventes"]))]
    # Extourne stricte sur le(s) compte(s) de ventes configuré(s) au sens large (ex. 701) :
    # facture annulée ou corrigée a posteriori — c'est le seul sens réel de « extourne sur ventes ».
    extournes_ventes = df_v_large["Débit"].sum() if not df_v_large.empty else 0

    ca_brut       = df_v["Crédit"].sum() if not df_v.empty else 0
    # corrections_ventes = TOUS les débits du périmètre élargi "ventes" (mask_ventes) : extournes
    # sur 701 + tout débit sur les autres comptes de produits inclus dans ce périmètre depuis
    # son élargissement (commissions, subventions, produits divers, reprises...). Sans déduire
    # ce total, le résultat net ne se réconcilie pas avec le total comptable réel des comptes
    # 6/7 — mais on isole ci-dessous la part qui est une vraie extourne de vente du reste, pour
    # ne pas induire en erreur sur la nature du montant (cf. Tableau de bord éditorial).
    corrections_ventes = df_v["Débit"].sum() if not df_v.empty else 0
    autres_regul_produits = corrections_ventes - extournes_ventes
    total_retours = abs(df_r["Montant_net"].sum())  if not df_r.empty else 0
    total_remises = abs(df_rem["Montant_net"].sum()) if not df_rem.empty else 0
    ca_net        = ca_brut - total_retours - total_remises - corrections_ventes
    # Reprises sur provisions (ex. 781/7810) : exclues du CA par mask_ventes (cf.
    # mask_provisions_reprises), nettées ici contre les charges — cf. Tableau de bord éditorial.
    df_prov_reprises = df[mask_provisions_reprises(df, params)]
    net_provisions_reprises = df_prov_reprises["Crédit"].sum() - df_prov_reprises["Débit"].sum()
    # Net débit-crédit (Montant_net déjà calculé par filtre_m ci-dessus).
    charges_tot     = (df_c["Montant_net"].sum() if not df_c.empty else 0) - net_provisions_reprises
    resultat_net  = ca_net - charges_tot
    marge_pct     = (resultat_net / ca_brut * 100) if ca_brut else 0

    soldes = [ca_brut, -total_retours, -total_remises, -corrections_ventes, ca_net, -charges_tot, resultat_net]
    libelles = ["CA brut", "Retours", "Remises", "Corrections produits", "CA net", "Charges", "Résultat net"]
    df_summary = pd.DataFrame({"Poste": libelles, "Montant (€)": soldes})
    if corrections_ventes:
        st.caption(
            f"ℹ️ Corrections produits ({fmt_fr(corrections_ventes, 0)} €) = "
            f"{fmt_fr(extournes_ventes, 0)} € d'extournes/corrections sur ventes proprement dites "
            f"(compte{'s' if len(params['ventes']) > 1 else ''} {', '.join(params['ventes'])}) + "
            f"{fmt_fr(autres_regul_produits, 0)} € d'autres régularisations (débits) sur les autres "
            "comptes de produits inclus dans le CA élargi (commissions, subventions, produits divers)."
        )
    if net_provisions_reprises:
        st.caption(f"ℹ️ Charges déjà nettes de {fmt_fr(net_provisions_reprises, 0)} € de reprises sur "
                   "provisions (ex. reprise de provision pour retour) — imputées à la dotation initiale "
                   "plutôt que comptées en CA.")

    # ================================================
    # 🚦 INDICATEURS CLÉS DE PILOTAGE (référentiel de la mission)
    # ================================================
    st.subheader("🚦 Indicateurs clés de pilotage")
    st.caption("Vue d'ensemble des 4 indicateurs du référentiel — le détail et les sous-indicateurs sont "
               "disponibles dans leur module respectif (📖 Analyse par titre, 💰 Trésorerie prévisionnelle, "
               "📦 Retours & Remises, ✍️ Droits d'auteurs).")

    ic1, ic2, ic3, ic4 = st.columns(4)

    # --- Indicateur 1 — Marge par titre ---
    titres_synth = sorted(filtrer_isbn_reels(df)["Code_Analytique"].astype(str).unique().tolist())
    if titres_synth:
        indic1_synth = calculer_indicateurs_titres(df, params, titres_synth)
        taux_mb_synth = np.where(indic1_synth["CA net"] != 0, indic1_synth["Marge brute"] / indic1_synth["CA net"] * 100, 0.0)
        nb_deficit = int((taux_mb_synth < 0).sum())
        with ic1:
            st.markdown("**1️⃣ Marge par titre**")
            st.metric("Titres déficitaires", nb_deficit,
                      delta="🔴 À traiter" if nb_deficit else "🟢 Aucun", delta_color="off")
    else:
        with ic1:
            st.markdown("**1️⃣ Marge par titre**")
            st.caption("Aucun titre actif.")

    # --- Indicateur 2 — Trésorerie prévisionnelle ---
    with ic2:
        st.markdown("**2️⃣ Trésorerie prévisionnelle**")
        if "treso_real" in st.session_state:
            treso_real_synth = st.session_state["treso_real"]
            scenarios_synth = st.session_state.get("treso_scenarios", {})
            if scenarios_synth and len(scenarios_synth["Central"].sum()) > 0:
                solde_m1 = treso_real_synth.iloc[-1] + scenarios_synth["Central"].sum().iloc[0]
                alerte_synth2 = "🔴 Négatif" if solde_m1 < 0 else "🟢 Positif"
                st.metric("Solde prévisionnel (M+1)", f"{fmt_fr(solde_m1, 0)} €", delta=alerte_synth2, delta_color="off")
            else:
                st.metric("Trésorerie de clôture (réalisé)", f"{fmt_fr(treso_real_synth.iloc[-1], 0)} €")
        else:
            st.caption("Consultez 💰 Trésorerie prévisionnelle.")

    # --- Indicateur 3 — Taux de retour ---
    taux_retour_synth = (total_retours / ca_brut_distrib * 100) if ca_brut_distrib else 0.0
    taux_remise_synth = (total_remises / ca_brut_distrib * 100) if ca_brut_distrib else 0.0
    if taux_retour_synth > 30:
        alerte_synth3 = "🔴 Excessif"
    elif taux_retour_synth >= 20:
        alerte_synth3 = "🟠 À surveiller"
    else:
        alerte_synth3 = "🟢 Normal"
    with ic3:
        st.markdown("**3️⃣ Taux de retour**")
        st.metric("Taux de retour global", f"{taux_retour_synth:.1f} %", delta=alerte_synth3, delta_color="off")
        st.caption(f"Taux de remise : {taux_remise_synth:.1f} % (base CA distributeur, {fmt_fr(ca_brut_distrib, 0)} €).")

    # --- Indicateur 4 — Droits d'auteurs ---
    with ic4:
        st.markdown("**4️⃣ Droits d'auteurs**")
        avance_debit_synth = df[df["Compte"].astype(str) == "409600"]["Débit"].sum()
        avance_credit_synth = df[df["Compte"].astype(str) == "409600"]["Crédit"].sum()
        avalent_restant_synth = avance_debit_synth - avance_credit_synth
        if avance_debit_synth > 0:
            st.metric("À-valoir restant à amortir", f"{fmt_fr(avalent_restant_synth, 0)} €")
        else:
            st.caption("Aucune écriture sur le compte 409600.")

    st.divider()

    col1, col2 = st.columns([1.2, 1])
    with col1:
        st.subheader("Compte de résultat synthétique")
        st.dataframe(df_summary.style.format({"Montant (€)": (lambda x: fmt_fr(x, 0))}), hide_index=True, height=280)
        st.metric("Taux de marge nette", f"{marge_pct:.1f} %")
    with col2:
        st.subheader("Waterfall")
        colors = ["#3B82F6", "#EF4444", "#EF4444", "#EF4444", "#10B981", "#EF4444",
                  "#10B981" if resultat_net >= 0 else "#EF4444"]
        fig = go.Figure(go.Waterfall(
            name="Résultat",
            orientation="v",
            measure=["absolute", "relative", "relative", "relative", "total", "relative", "total"],
            x=libelles, y=soldes,
            connector={"line": {"color": "gray", "width": 0.5}},
            decreasing={"marker": {"color": "#EF4444"}},
            increasing={"marker": {"color": "#10B981"}},
            totals={"marker": {"color": "#3B82F6"}}
        ))
        fig.update_layout(height=350, margin=dict(t=20), yaxis_title="€")
        st.plotly_chart(fig, use_container_width=True)

    # Export PDF synthèse (Excel ici, PDF nécessiterait reportlab)
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_summary.to_excel(writer, index=False, sheet_name="Synthese")
    buffer.seek(0)
    st.download_button("📥 Exporter la synthèse (Excel)", buffer,
                        file_name="Synthese_Financiere.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.divider()
    st.markdown("##### 📋 Référentiel des 4 indicateurs de pilotage")
    st.caption("Définition, formule et seuil d'alerte de chaque indicateur — le détail chiffré est disponible "
               "dans le module correspondant.")
    recap = pd.DataFrame([
        {"#": 1, "Indicateur": "Marge par titre", "Formule simplifiée": "CA net − Charges directes − Variation stock",
         "Seuil alerte": "Marge brute < 0 % → alerte rouge", "Source principale": "Export analytique (701/709/6xx)",
         "Module détaillé": "📖 Analyse par titre"},
        {"#": 2, "Indicateur": "Trésorerie prévisionnelle", "Formule simplifiée": "Encaissements prévisionnels − Décaissements",
         "Seuil alerte": "Solde prévisionnel < 0 → alerte rouge", "Source principale": "BLDD + Compta générale + Programme éditorial",
         "Module détaillé": "💰 Trésorerie prévisionnelle"},
        {"#": 3, "Indicateur": "Taux de retour", "Formule simplifiée": "Retours / Ventes brutes",
         "Seuil alerte": "Taux > 30 % → alerte rouge", "Source principale": "Relevé BLDD mensuel",
         "Module détaillé": "📦 Retours & Remises"},
        {"#": 4, "Indicateur": "Droits d'auteurs", "Formule simplifiée": "Ventes nettes PPHT × 10 %",
         "Seuil alerte": "À-valoir non amorti > 12 mois", "Source principale": "BLDD + Contrats auteurs + Compte 409600",
         "Module détaillé": "✍️ Droits d'auteurs"},
    ])
    st.dataframe(recap, use_container_width=True, hide_index=True)

# =====================
# ASSISTANT IA — 3 NIVEAUX
# =====================
elif page == "🤖 Assistant IA":
    st.header("🤖 Assistant IA — Expert éditorial")
    client = get_client()
    if client is None:
        st.error("⚠️ Clé API Anthropic non configurée. Ajoutez `ANTHROPIC_API_KEY` dans vos secrets Streamlit Cloud.")
        st.code("""
# Dans Streamlit Cloud > Settings > Secrets :
ANTHROPIC_API_KEY = "sk-ant-..."
        """)
        st.stop()

    tab1, tab2, tab3 = st.tabs([
        "💬 Niveau 1 — Questions sur les modules",
        "🔍 Niveau 2 — Analyse de vos données",
        "🤖 Niveau 3 — Agent conversationnel"
    ])

    # ─── NIVEAU 1 : Assistant contextuel ───
    with tab1:
        st.markdown("**Posez une question sur l'utilisation de l'outil ou sur la comptabilité éditoriale.**")
        questions_types = [
            "Comment paramétrer mes comptes de retours ?",
            "Que signifie le taux de retour et quel seuil est normal en édition ?",
            "Comment fonctionne la provision pour retours (compte 681) ?",
            "Comment lire le tableau de bord éditorial ?",
            "Qu'est-ce qu'une reddition de comptes d'auteur ?",
            "Comment calculer la marge brute éditeur ?"
        ]
        st.markdown("**Questions fréquentes :**")
        cols = st.columns(3)
        for i, q in enumerate(questions_types):
            if cols[i % 3].button(q, key=f"q1_{i}", use_container_width=True):
                with st.spinner("Réflexion en cours..."):
                    resp = client.messages.create(
                        model="claude-opus-4-5",
                        max_tokens=600,
                        system=SYSTEM_PROMPT,
                        messages=[{"role": "user", "content": q}]
                    )
                    st.info(f"**Question :** {q}")
                    st.success(resp.content[0].text)
        st.divider()
        question_libre = st.text_input("Ou posez votre propre question :", placeholder="Ex: Comment traiter un avoir client en édition ?")
        if st.button("Envoyer", key="btn_n1") and question_libre:
            with st.spinner("Réflexion en cours..."):
                resp = client.messages.create(
                    model="claude-opus-4-5",
                    max_tokens=600,
                    system=SYSTEM_PROMPT,
                    messages=[{"role": "user", "content": question_libre}]
                )
                st.success(resp.content[0].text)

    # ─── NIVEAU 2 : Analyste sur les données ───
    with tab2:
        if "df_pivot" not in st.session_state:
            st.warning("⚠️ Chargez d'abord des données pour activer l'analyse IA.")
        else:
            st.markdown("**L'IA analyse vos données et génère un commentaire de gestion personnalisé.**")
            if st.button("🔍 Analyser mes données maintenant", type="primary"):
                summary = build_data_summary()
                prompt = f"""Voici les données analytiques d'une maison d'édition indépendante :
{summary}

En tant qu'expert-comptable spécialisé en édition, fournis :
1. Un commentaire de gestion synthétique (3-4 phrases)
2. Les 2-3 points d'attention prioritaires avec leur niveau de risque
3. Des recommandations concrètes et actionnables pour améliorer la rentabilité
4. Une évaluation du taux de retour par rapport aux standards du secteur (entre 20% et 35% en édition française)

Sois précis, factuel et orienté action."""
                with st.spinner("Analyse en cours..."):
                    resp = client.messages.create(
                        model="claude-opus-4-5",
                        max_tokens=900,
                        system=SYSTEM_PROMPT,
                        messages=[{"role": "user", "content": prompt}]
                    )
                    st.markdown("### Analyse IA de vos données")
                    st.markdown(resp.content[0].text)
            st.divider()
            st.markdown("**Demande d'analyse spécifique :**")
            analyse_custom = st.text_area("Que souhaitez-vous analyser en particulier ?",
                                           placeholder="Ex: Quels sont les signaux d'alerte sur ma trésorerie prévisionnelle ?")
            if st.button("Analyser", key="btn_n2") and analyse_custom:
                summary = build_data_summary()
                with st.spinner("Analyse en cours..."):
                    resp = client.messages.create(
                        model="claude-opus-4-5",
                        max_tokens=800,
                        system=SYSTEM_PROMPT,
                        messages=[{
                            "role": "user",
                            "content": f"Données de l'éditeur :\n{summary}\n\nQuestion : {analyse_custom}"
                        }]
                    )
                    st.markdown(resp.content[0].text)

    # ─── NIVEAU 3 : Agent conversationnel multi-tours ───
    with tab3:
        st.markdown("**Dialogue libre et multi-tours avec l'agent. Il mémorise le contexte de la conversation.**")
        # Contexte données injecté au départ
        data_context = ""
        if "df_pivot" in st.session_state:
            data_context = f"\n\nDONNÉES DISPONIBLES :\n{build_data_summary()}"
        system_with_data = SYSTEM_PROMPT + data_context

        # Affichage historique
        for msg in st.session_state["messages_agent"]:
            with st.chat_message(msg["role"], avatar="🧑" if msg["role"] == "user" else "🤖"):
                st.markdown(msg["content"])

        # Input utilisateur
        if prompt_user := st.chat_input("Posez votre question à l'agent..."):
            st.session_state["messages_agent"].append({"role": "user", "content": prompt_user})
            with st.chat_message("user", avatar="🧑"):
                st.markdown(prompt_user)
            with st.chat_message("assistant", avatar="🤖"):
                with st.spinner("L'agent réfléchit..."):
                    resp = client.messages.create(
                        model="claude-opus-4-5",
                        max_tokens=1000,
                        system=system_with_data,
                        messages=st.session_state["messages_agent"]
                    )
                    answer = resp.content[0].text
                    st.markdown(answer)
                    st.session_state["messages_agent"].append({"role": "assistant", "content": answer})

        if st.button("🗑️ Effacer la conversation", key="clear_chat"):
            st.session_state["messages_agent"] = []
            st.rerun()


# =====================
# PAGE : RAPPORT DE PILOTAGE
# =====================
# À ajouter dans la liste des pages :
#   "📄 Rapport de pilotage"
# Et dans le bloc elif :
elif page == "📄 Rapport de pilotage":
    st.header("📄 Rapport de pilotage analytique")
    st.markdown("*Génération automatique du rapport de restitution au dirigeant - Format Word professionnel*")

    if "df_pivot" not in st.session_state:
        st.warning("⚠️ Aucune donnée chargée. Veuillez d'abord importer votre Grand Livre analytique dans la page **📂 Import des données**.")
        st.stop()

    import datetime as _dt
    df     = st.session_state["df_pivot"].copy()
    params = st.session_state.get("param_comptes", {})

    st.subheader("⚙️ Paramètres du rapport")
    col1, col2 = st.columns(2)
    with col1:
        nom_editeur   = st.text_input("Nom de la maison d'édition", value="Éditions de l'Argonaute")
        nom_cabinet   = st.text_input("Nom du cabinet", value="CAB ÉDITION")
        exercice      = st.text_input("Exercice", value="2025/2026")
    with col2:
        date_rapport  = st.date_input("Date du rapport", value=_dt.date.today())
        ref_mission   = st.text_input("Référence lettre de mission", value="LM-2025-001")
        periode_debut = st.text_input("Début de période", value="01/04/2025")
        periode_fin   = st.text_input("Fin de période", value="31/03/2026")

    mode_anon = st.checkbox("🕶️ Anonymiser les titres", value=st.session_state.get("mode_anonyme", True))
    nb_top    = st.slider("Nombre de titres Top / Flop", 3, 10, 5)
    st.divider()

    titres_actifs = st.session_state.get("titres_actifs", [])
    if not titres_actifs:
        titres_actifs = df[df["Code_Analytique"].astype(str).str.match(r"^97[89]\d{10}")]["Code_Analytique"].unique().tolist()
    if not titres_actifs:
        st.error("❌ Aucun EAN valide trouvé. Vérifiez le paramétrage analytique.")
        st.stop()

    res = calculer_indicateurs_titres(df, params, titres_actifs)
    mapping = obtenir_mapping_anonymisation(df)
    res["Titre"] = res["Code_Analytique"].apply(lambda c: mapping.get(c, c) if mode_anon else label_affiche(c, df))

    # Calcul aligné sur le tableau de bord — utiliser les mêmes mask_ pour cohérence
    df_v_   = df[mask_ventes(df, params)]
    df_r_   = df[mask_retours(df, params)]
    df_rem_ = df[mask_remises(df, params)]
    df_c_   = df[mask_charges(df, params)]
    net_prov_ = 0
    if params.get("provisions_reprises"):
        _mp = df["Compte"].astype(str).str.startswith(tuple(params["provisions_reprises"]))
        net_prov_ = df[_mp]["Crédit"].sum() - df[_mp]["Débit"].sum()
    ca_total  = (df_v_["Crédit"].sum() - df_v_["Débit"].sum()
                 - (df_r_["Débit"].sum() - df_r_["Crédit"].sum())
                 - (df_rem_["Débit"].sum() - df_rem_["Crédit"].sum()))
    charges_tot_ = (df_c_["Débit"].sum() - df_c_["Crédit"].sum()) - net_prov_
    mn_total  = ca_total - charges_tot_
    # Charges indirectes : chercher dans Code_Analytique (non réparties)
    # ET dans Compte (réparties via le module paramétrage)
    valeurs_ci_code = ["CHARGES INDIRECTES", "FRAIS GENERAUX"]
    valeurs_ci_code += [x.upper() for x in params.get("libelles_indirects", [])]
    mask_ci_code = df["Code_Analytique"].astype(str).str.upper().isin(valeurs_ci_code)
    mask_ci_cpt  = df["Compte"].astype(str) == COMPTE_CHARGES_INDIRECTES_REPARTIES
    ci_total     = df[mask_ci_code | mask_ci_cpt]["Débit"].sum() - df[mask_ci_code | mask_ci_cpt]["Crédit"].sum()
    # Fallback Famille_Analytique si toujours 0
    if ci_total == 0 and "Famille_Analytique" in df.columns:
        mask_ci2 = df["Famille_Analytique"].astype(str).str.upper().str.contains("INDIRECT|STRUCTURE|GENERAUX", na=False)
        ci_total = df[mask_ci2]["Débit"].sum() - df[mask_ci2]["Crédit"].sum()
    nb_eans    = len(titres_actifs)
    ci_par_ean = ci_total / nb_eans if nb_eans > 0 else 0
    mb_total   = res["Marge brute"].sum()
    taux_mb    = (mb_total / ca_total * 100) if ca_total > 0 else 0
    n_pos      = (res["Marge brute"] > 0).sum()
    n_neg      = (res["Marge brute"] <= 0).sum()
    res_sorted = res.sort_values("Marge brute", ascending=False).reset_index(drop=True)

    st.subheader("📊 Aperçu")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("CA net", f"{fmt_fr(ca_total)} €")
    m2.metric("Marge brute", f"{fmt_fr(mb_total)} €", f"{fmt_fr(taux_mb,1)} %")
    m3.metric("Résultat net", f"{fmt_fr(mn_total)} €")
    m4.metric("Titres déficitaires", f"{n_neg} / {nb_eans}")

    col_t, col_f = st.columns(2)
    with col_t:
        st.markdown(f"**🟢 Top {nb_top}**")
        st.dataframe(res_sorted.head(nb_top)[["Titre","CA net","Marge brute"]].style.format({"CA net":"{:,.0f}","Marge brute":"{:+,.0f}"}), hide_index=True, use_container_width=True)
    with col_f:
        st.markdown(f"**🔴 Flop {nb_top}**")
        st.dataframe(res_sorted.tail(nb_top).iloc[::-1][["Titre","CA net","Marge brute"]].style.format({"CA net":"{:,.0f}","Marge brute":"{:+,.0f}"}), hide_index=True, use_container_width=True)

    st.divider()
    st.subheader("📄 Générer le rapport Word")

    if st.button("🖨️ Générer le rapport de pilotage", type="primary", use_container_width=True):
        with st.spinner("Génération en cours..."):
            try:
                from docx import Document as _Doc
                from docx.shared import Pt as _Pt, Cm as _Cm, RGBColor as _RGB
                from docx.enum.text import WD_ALIGN_PARAGRAPH as _AP
                from docx.enum.table import WD_ALIGN_VERTICAL as _AV
                from docx.oxml.ns import qn as _qn
                from docx.oxml import OxmlElement as _OE
                from io import BytesIO as _BIO

                _MR = _RGB(0x1F,0x4E,0x79)
                _TR = _RGB(0xC0,0x52,0x2A)
                _WH = _RGB(0xFF,0xFF,0xFF)
                _GC = "E2EFDA"; _RC = "FFDDD9"; _YC = "FFF2CC"
                _BC = "D6E4F0"; _LC = "F5E6DF"

                def _shd(cell, h):
                    tc = cell._tc; p = tc.get_or_add_tcPr()
                    for o in p.findall(_qn("w:shd")): p.remove(o)
                    s = _OE("w:shd"); s.set(_qn("w:val"),"clear"); s.set(_qn("w:color"),"auto"); s.set(_qn("w:fill"),h); p.append(s)

                def _brd(cell, c="CCCCCC", z="4"):
                    tc = cell._tc; p = tc.get_or_add_tcPr()
                    for o in p.findall(_qn("w:tcBorders")): p.remove(o)
                    b = _OE("w:tcBorders")
                    for e in ["top","left","bottom","right"]:
                        x = _OE(f"w:{e}"); x.set(_qn("w:val"),"single"); x.set(_qn("w:sz"),z); x.set(_qn("w:color"),c); b.append(x)
                    p.append(b)

                def _wc(cell, text, bold=False, italic=False, color=None, size=9, center=False, fill=None):
                    if fill: _shd(cell, fill)
                    _brd(cell)
                    cell.vertical_alignment = _AV.CENTER
                    p = cell.paragraphs[0]; p.clear()
                    p.alignment = _AP.CENTER if center else _AP.LEFT
                    p.paragraph_format.space_before = _Pt(2); p.paragraph_format.space_after = _Pt(2)
                    r = p.add_run(str(text)); r.bold=bold; r.italic=italic
                    r.font.size=_Pt(size); r.font.name="Arial"
                    r.font.color.rgb = color if color else _RGB(0,0,0)

                def _hr(tbl, cols, fill="1F4E79"):
                    row = tbl.rows[0]
                    for i,(t,w) in enumerate(cols):
                        c = row.cells[i]; _wc(c,t,bold=True,color=_WH,size=9,center=True,fill=fill)
                        if w: c.width=_Cm(w)

                def _fn(n, sign=False):
                    return ("+" if (sign and n>=0) else "") + f"{n:,.0f} EUR".replace(",", " ")

                def _fp(n):
                    return f"{'+'if n>=0 else ''}{n:.1f}%"

                def _sep(d, col="C0522A"):
                    p=d.add_paragraph(); p.paragraph_format.space_before=_Pt(2); p.paragraph_format.space_after=_Pt(6)
                    pr=p._p.get_or_add_pPr(); pb=_OE("w:pBdr"); bt=_OE("w:bottom")
                    bt.set(_qn("w:val"),"single"); bt.set(_qn("w:sz"),"6"); bt.set(_qn("w:color"),col)
                    pb.append(bt); pr.append(pb)

                def _P(d, text, bold=False, italic=False, size=10, color=None, align="left", sb=6, sa=6, line=None):
                    p=d.add_paragraph(); p.paragraph_format.space_before=_Pt(sb); p.paragraph_format.space_after=_Pt(sa)
                    if line: p.paragraph_format.line_spacing=_Pt(line)
                    p.alignment={"left":_AP.LEFT,"center":_AP.CENTER,"right":_AP.RIGHT,"justify":_AP.JUSTIFY}.get(align,_AP.LEFT)
                    r=p.add_run(text); r.bold=bold; r.italic=italic; r.font.size=_Pt(size); r.font.name="Arial"
                    if color: r.font.color.rgb=color
                    return p

                def _Pm(d, runs, align="left", sb=6, sa=6, keep_next=True):
                    p=d.add_paragraph(); p.paragraph_format.space_before=_Pt(sb); p.paragraph_format.space_after=_Pt(sa)
                    p.paragraph_format.keep_with_next=keep_next
                    p.alignment={"left":_AP.LEFT,"center":_AP.CENTER,"right":_AP.RIGHT}.get(align,_AP.LEFT)
                    for text,bold,italic,size,color in runs:
                        r=p.add_run(text); r.bold=bold; r.italic=italic; r.font.size=_Pt(size); r.font.name="Arial"
                        if color: r.font.color.rgb=color
                    return p

                top_list  = [res_sorted.iloc[i] for i in range(min(nb_top, len(res_sorted)))]
                flop_list = [res_sorted.iloc[-(i+1)] for i in range(min(nb_top, len(res_sorted)))]

                d = _Doc()
                for s in d.sections:
                    s.top_margin=_Cm(2); s.bottom_margin=_Cm(2); s.left_margin=_Cm(2.5); s.right_margin=_Cm(2)

                # En-tete
                _P(d, nom_cabinet, bold=True, size=14, color=_MR, sb=0, sa=2)
                _P(d, "Expert-comptable - Commissaire aux comptes", size=10, color=_RGB(0x55,0x55,0x55), sb=0, sa=2)
                _P(d, "Membre de l'Ordre des Experts-Comptables", size=9, italic=True, color=_RGB(0x88,0x88,0x88), sb=0, sa=4)
                pd=d.add_paragraph(); pd.alignment=_AP.RIGHT; pd.paragraph_format.space_before=_Pt(0); pd.paragraph_format.space_after=_Pt(4)
                _mois_fr = {1:"janvier",2:"février",3:"mars",4:"avril",5:"mai",6:"juin",7:"juillet",8:"août",9:"septembre",10:"octobre",11:"novembre",12:"décembre"}
                _date_fr = f"{date_rapport.day} {_mois_fr[date_rapport.month]} {date_rapport.year}"
                rd=pd.add_run(f"Lille, le {_date_fr}"); rd.font.size=_Pt(10); rd.font.name="Arial"
                _sep(d)

                _P(d, "Monsieur le Directeur", bold=True, size=11, sb=4, sa=2)
                _P(d, nom_editeur, size=11, sb=0, sa=2)
                _P(d, "Siege social : Lille (59)", size=10, color=_RGB(0x55,0x55,0x55), sb=0, sa=10)

                to=d.add_table(rows=1, cols=1); to.style="Table Grid"
                co=to.rows[0].cells[0]; _shd(co,_BC); _brd(co)
                po=co.paragraphs[0]; po.paragraph_format.space_before=_Pt(4); po.paragraph_format.space_after=_Pt(4)
                r1o=po.add_run("Objet : "); r1o.bold=True; r1o.font.size=_Pt(10); r1o.font.name="Arial"; r1o.font.color.rgb=_MR
                r2o=po.add_run(f"Rapport de pilotage analytique - Exercice {exercice} - Mission n° {ref_mission}"); r2o.font.size=_Pt(10); r2o.font.name="Arial"

                _P(d, "", sb=10, sa=0)
                _P(d, "Monsieur,", bold=True, size=11, sb=0, sa=6)
                _P(d, f"Conformement aux diligences prevues dans notre lettre de mission n° {ref_mission} et a l'organisation trimestrielle de restitution convenue avec vous, nous avons l'honneur de vous adresser le present rapport de pilotage analytique relatif a l'exercice clos le {periode_fin}.", size=10, align="justify", sb=0, sa=6, line=14)
                _P(d, f"Ce rapport a ete etabli a partir des donnees de votre comptabilite analytique pour la periode du {periode_debut} au {periode_fin}, traitees par notre outil Python VISION EDITION. Il porte sur l'ensemble des {nb_eans} ISBN actifs de votre catalogue.", size=10, align="justify", sb=0, sa=6, line=14)
                _P(d, f"Nous attirons votre attention sur le point central suivant : si votre maison d'edition degage une marge brute globale de {_fp(taux_mb)}, la prise en compte des charges indirectes de structure conduit a un resultat analytique net de {_fn(mn_total, True)}. Cette situation appelle une analyse approfondie que nous vous soumettons ci-apres.", size=10, align="justify", sb=0, sa=10, line=14)

                # I
                _Pm(d,[("I.   ",True,False,13,_TR),("SYNTHESE DE LA PERFORMANCE ANALYTIQUE",True,False,13,_MR)],sb=8,sa=4)
                _sep(d)
                _P(d,"L'analyse du grand livre analytique de l'exercice fait ressortir les resultats suivants :",size=10,sb=0,sa=6)
                ts=d.add_table(rows=6,cols=3); ts.style="Table Grid"
                _hr(ts,[("Indicateur",7),("Montant",3.5),("Commentaire",6)])
                for i,(lbl,val,cmt,fl) in enumerate([
                    ("CA net editeur",_fn(ca_total),f"{nb_eans} ISBN actifs",_BC),
                    ("Charges directes",_fn(ca_total-mb_total),"Fabrication, droits, personnel","FFFFFF"),
                    ("Marge brute sur ventes",_fn(mb_total,True),f"Taux : {_fp(taux_mb)}",_GC),
                    ("Charges indirectes de structure",_fn(ci_total),f"Quote-part/ISBN : {ci_par_ean:.2f} EUR",_YC),
                    ("Resultat analytique net",_fn(mn_total,True),"Apres affectation charges de structure",_RC if mn_total<0 else _GC),
                ]):
                    rw=ts.rows[i+1]
                    _wc(rw.cells[0],lbl,bold=True,color=_MR,size=9,fill=fl)
                    _wc(rw.cells[1],val,bold=True,center=True,size=10,color=_RGB(0xC0,0,0) if (i==4 and mn_total<0) else _MR,fill=fl)
                    _wc(rw.cells[2],cmt,italic=True,size=8,color=_RGB(0x55,0x55,0x55),fill=fl)
                _P(d,f"L'exercice {exercice} presente une marge brute de {_fn(mb_total,True)} ({_fp(taux_mb)}). Toutefois, les charges indirectes ({_fn(ci_total)}), reparties sur {nb_eans} ISBN actifs a raison de {ci_par_ean:.2f} EUR/titre, conduisent a un resultat analytique net de {_fn(mn_total,True)}.",size=10,align="justify",sb=6,sa=10,line=14)

                # II
                _Pm(d,[("II.   ",True,False,13,_TR),("ANALYSE DE LA RENTABILITE PAR TITRE",True,False,13,_MR)],sb=8,sa=4)
                _sep(d)
                _P(d,f"Sur les {nb_eans} ISBN actifs au catalogue, {n_pos} degagent une marge brute positive et {n_neg} affichent une marge brute negative ({n_neg/max(nb_eans,1)*100:.1f} % du catalogue).",size=10,align="justify",sb=0,sa=8,line=14)

                _P(d,f"A. Les {nb_top} titres les plus contributifs",bold=True,size=10,color=_MR,sb=4,sa=4)
                tt=d.add_table(rows=nb_top+1,cols=5); tt.style="Table Grid"
                _hr(tt,[("Rg",0.8),("Titre",5.5),("CA net (EUR)",2.2),("Marge brute (EUR)",2.8),("Statut",1.2)])
                for i,rd_ in enumerate(top_list):
                    bg=_LC if i%2==0 else "FFFFFF"; rw=tt.rows[i+1]
                    _wc(rw.cells[0],str(i+1),bold=True,center=True,size=9,color=_TR,fill=bg)
                    _wc(rw.cells[1],str(rd_["Titre"])[:44],bold=True,size=9,color=_MR,fill=bg)
                    _wc(rw.cells[2],f"{rd_['CA net']:,.0f} EUR".replace(",", " "),center=True,size=9,fill=bg)
                    _wc(rw.cells[3],f"{rd_['Marge brute']:+,.0f} EUR".replace(",", " "),bold=True,center=True,size=10,color=_RGB(0x21,0x73,0x46),fill=_GC)
                    _wc(rw.cells[4],"OK",bold=True,center=True,size=9,color=_RGB(0x21,0x73,0x46),fill=_GC)
                if top_list:
                    _P(d,f"Nous attirons votre attention sur '{top_list[0]['Titre']}' qui, avec une marge brute de {top_list[0]['Marge brute']:+,.0f} EUR, constitue le titre le plus contributif de votre catalogue.",size=10,align="justify",sb=6,sa=10,line=14)

                _P(d,f"B. Les {nb_top} titres les plus deficitaires - Action corrective requise",bold=True,size=10,color=_RGB(0xC0,0,0),sb=4,sa=4)
                tf=d.add_table(rows=nb_top+1,cols=5); tf.style="Table Grid"
                _hr(tf,[("Rg",0.8),("Titre",5.5),("CA net (EUR)",2.2),("Marge brute (EUR)",2.8),("Statut",1.2)],fill="C00000")
                for i,rd_ in enumerate(flop_list):
                    bg="FFF5F5" if i%2==0 else "FFFFFF"; rw=tf.rows[i+1]
                    _wc(rw.cells[0],str(i+1),bold=True,center=True,size=9,color=_RGB(0xC0,0,0),fill=bg)
                    _wc(rw.cells[1],str(rd_["Titre"])[:44],bold=True,size=9,color=_RGB(0xC0,0,0),fill=bg)
                    _wc(rw.cells[2],f"{rd_['CA net']:,.0f} EUR".replace(",", " "),center=True,size=9,fill=bg)
                    _wc(rw.cells[3],f"{rd_['Marge brute']:+,.0f} EUR".replace(",", " "),bold=True,center=True,size=10,color=_RGB(0xC0,0,0),fill=_RC)
                    _wc(rw.cells[4],"Alerte",bold=True,center=True,size=9,color=_RGB(0xC0,0,0),fill=_RC)
                if flop_list:
                    _P(d,f"Nous appelons votre attention sur '{flop_list[0]['Titre']}' qui accuse une perte de {abs(flop_list[0]['Marge brute']):,.0f} EUR. L'ensemble de ces titres deficitaires appelle les mesures correctives detaillees dans la section IV.",size=10,align="justify",sb=6,sa=10,line=14)

                # III
                _Pm(d,[("III.   ",True,False,13,_TR),("CHARGES INDIRECTES DE STRUCTURE",True,False,13,_MR)],sb=8,sa=4)
                _sep(d)
                _P(d,f"Les charges indirectes de structure s'elevent a {_fn(ci_total)}, soit {ci_total/max(ca_total,1)*100:.1f} % du chiffre d'affaires net. Avec {nb_eans} ISBN actifs, chaque titre doit degager une marge brute minimale de {ci_par_ean:.0f} EUR pour couvrir sa quote-part de structure.",size=10,align="justify",sb=0,sa=10,line=14)

                # IV
                _Pm(d,[("IV.   ",True,False,13,_TR),("RECOMMANDATIONS ET PLAN D'ACTION",True,False,13,_MR)],sb=8,sa=4)
                _sep(d)
                _P(d,"Au vu de l'analyse ci-avant, nous vous soumettons les recommandations suivantes :",size=10,sb=0,sa=6)
                tr=d.add_table(rows=5,cols=4); tr.style="Table Grid"
                _hr(tr,[("Priorite",1.5),("Recommandation",5.5),("Action",4),("Delai",2)])
                for i,(prio,rec,action,delai,bg) in enumerate([
                    ("1 - URGENT",f"Analyse des {n_neg} titres deficitaires","Decision maintien ou arret commercial","Immediat",_RC),
                    ("2 - URGENT","Arret reimpression des titres deficitaires","Gel des commandes impression","Immediat","FFFFFF"),
                    ("3 - IMPORTANT","Reexamen des charges indirectes de structure","Audit des postes de charges","T2 2026",_RC),
                    ("4 - IMPORTANT",f"Reimpression prioritaire des {nb_top} titres porteurs","Lancement procedure reimpression","T2 2026","FFFFFF"),
                ]):
                    rw=tr.rows[i+1]; iu="URGENT" in prio
                    _wc(rw.cells[0],prio,bold=True,center=True,size=8,color=_RGB(0xC0,0,0) if iu else _RGB(0xA6,0x7C,0),fill=_RC if iu else _YC)
                    _wc(rw.cells[1],rec,bold=True,size=9,fill=bg)
                    _wc(rw.cells[2],action,size=9,fill=bg)
                    _wc(rw.cells[3],delai,center=True,bold=True,size=9,color=_RGB(0xC0,0,0) if iu else _RGB(0xA6,0x7C,0),fill=bg)

                _P(d,"",sb=10,sa=0)

                # V
                _Pm(d,[("V.   ",True,False,13,_TR),("CONCLUSION ET PROCHAINES ETAPES",True,False,13,_MR)],sb=8,sa=4)
                _sep(d)
                _P(d,"Le present rapport met en evidence la necessite d'engager rapidement une reflexion sur la structure de votre catalogue et sur le niveau de vos charges indirectes. Nous vous proposons de nous reunir dans les prochaines semaines afin de discuter des recommandations formulees ci-dessus.",size=10,align="justify",sb=0,sa=6,line=14)
                _P(d,"Nous demeurons a votre entiere disposition et vous prions d'agreer, Monsieur, l'expression de nos salutations distinguees.",size=10,align="justify",sb=0,sa=16,line=14)

                # Signatures
                ts2=d.add_table(rows=1,cols=2); ts2.style="Table Grid"
                for cs,ti,fi,ci_ in [(ts2.rows[0].cells[0],f"L'Expert-Comptable\n{nom_cabinet}",_BC,_MR),(ts2.rows[0].cells[1],f"Le Dirigeant\n{nom_editeur}",_LC,_TR)]:
                    _shd(cs,fi); _brd(cs)
                    ps=cs.paragraphs[0]; ps.paragraph_format.space_before=_Pt(6); ps.paragraph_format.space_after=_Pt(6)
                    for j,li in enumerate([ti,"\n\nSignature : ______________________","\nDate : ___________________________"]):
                        rs=ps.add_run(li); rs.font.size=_Pt(9); rs.font.name="Arial"; rs.bold=(j==0); rs.font.color.rgb=ci_

                # Note
                _P(d,"",sb=8,sa=0)
                pn=d.add_paragraph(); pn.paragraph_format.space_before=_Pt(4); pn.paragraph_format.space_after=_Pt(4)
                rn1=pn.add_run("Note methodologique : "); rn1.bold=True; rn1.font.size=_Pt(8); rn1.font.name="Arial"; rn1.font.color.rgb=_RGB(0x55,0x55,0x55)
                an="Titres anonymises. " if mode_anon else ""
                rn2=pn.add_run(f"Rapport genere automatiquement par Python VISION EDITION ({periode_debut} au {periode_fin}, {nb_eans} ISBN actifs). {an}Charges indirectes ({_fn(ci_total)}) reparties forfaitairement ({ci_par_ean:.2f} EUR/titre). Source : comptabilite analytique {nom_editeur} - exercice {exercice} - donnees provisoires.")
                rn2.italic=True; rn2.font.size=_Pt(8); rn2.font.name="Arial"; rn2.font.color.rgb=_RGB(0x77,0x77,0x77)

                buf=_BIO(); d.save(buf); buf.seek(0)
                sfx="_anon" if mode_anon else ""
                fn=f"Rapport_Pilotage_{nom_editeur[:20].replace(' ','_')}_{exercice.replace('/','_')}{sfx}.docx"
                st.success("✅ Rapport généré avec succès !")
                st.download_button(label="⬇️ Télécharger le rapport Word", data=buf, file_name=fn,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True, type="primary")

            except ImportError:
                st.error("❌ python-docx non installé. Ajoutez `python-docx` dans requirements.txt.")
            except Exception as e:
                st.error(f"❌ Erreur : {str(e)}")
                st.exception(e)

# =====================
# FOOTER
# =====================
st.divider()
st.markdown(
    "<div style='text-align:center; color:#888; font-size:12px'>"
    "© 2025 Nicolas CUISSET — Mémoire d'expertise comptable — Maisons d'édition indépendantes"
    "</div>",
    unsafe_allow_html=True
)
