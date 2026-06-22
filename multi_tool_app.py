# -*- coding: utf-8 -*-
"""
multi_tool_app.py — Couche UI uniquement.
Toute la logique métier est dans services/.
La persistence est dans database/.

Pour brancher une base de données :
    1. Implémenter database/repository.IResultRepository
    2. Remplacer la ligne :  _repo = SessionRepository()
       par votre implémentation (ex: PostgresRepository, SupabaseRepository…)
    3. C'est tout — aucune autre modification nécessaire.
"""

from __future__ import annotations
import io, os, tempfile, zipfile
from datetime import datetime
from typing import List, Tuple

import pandas as pd
import streamlit as st

# ── Services (logique métier pure) ────────────────────────────────────────────
from services import transforms, io_service, m2_service
from services import classification_service, multiconnexion_service
from services import pc_service, cpn_service
from services.transforms import sanitize_code, sanitize_numeric, to_xlsx

# ── Repository (persistence) ──────────────────────────────────────────────────
from database.session_repository import SessionRepository

# ════════════════════════════════════════════════════════════════════════════════
# POINT D'ÉCHANGE : remplacez cette ligne pour changer de backend de persistence
# Ex: _repo = PostgresRepository(os.environ["DATABASE_URL"])
# ════════════════════════════════════════════════════════════════════════════════
_repo = SessionRepository()

# ── Imports optionnels ────────────────────────────────────────────────────────
try:
    import win32com.client as win32  # type: ignore
    IS_OUTLOOK = True
except ImportError:
    IS_OUTLOOK = False

try:
    import psutil  # type: ignore
    def _ram() -> str:
        return f"{psutil.Process().memory_info().rss / 1_048_576:,.0f} Mo"
except ModuleNotFoundError:
    def _ram() -> str:
        return "n/a"

# ── Config ────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Rexel Toolbox", page_icon="🛠", layout="wide")


def _today() -> str:
    return datetime.today().strftime("%y%m%d")


# ══ DESIGN SYSTEM CSS ═════════════════════════════════════════════════════════
def _inject_css() -> None:
    st.markdown("""
<style>
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 1.5rem !important; padding-bottom: 2rem !important; }
.stApp, [data-testid="stAppViewContainer"], [data-testid="stMainBlockContainer"] {
    background-color: #ffffff !important;
}
[data-testid="stMain"] { background-color: #ffffff !important; }

[data-testid="stSidebar"] {
    background: hsl(229, 84%, 39%) !important;
    border-right: none !important;
}
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] div { color: rgba(255,255,255,0.92) !important; }
[data-testid="stSidebar"] .stRadio > div > label {
    border-radius: 8px !important;
    padding: 9px 14px !important;
    margin: 2px 6px !important;
    cursor: pointer;
    transition: background 0.15s ease;
    font-size: 0.875rem !important;
    font-weight: 500 !important;
}
[data-testid="stSidebar"] .stRadio > div > label:hover {
    background: hsl(229, 84%, 30%) !important;
}
[data-testid="stSidebar"] .stRadio > div > label[data-checked="true"] {
    background: hsl(229, 84%, 30%) !important;
    font-weight: 700 !important;
}
[data-testid="stSidebar"] .stRadio > div > label > div:first-child { display: none !important; }
[data-testid="stSidebar"] .stButton > button {
    background: rgba(255,255,255,0.15) !important;
    color: white !important;
    border: 1px solid rgba(255,255,255,0.35) !important;
    border-radius: 8px !important;
    width: calc(100% - 16px) !important;
    margin: 4px 8px !important;
    font-size: 0.8rem !important;
    font-weight: 500 !important;
    box-shadow: none !important;
}
[data-testid="stSidebar"] .stButton > button:hover {
    background: rgba(255,255,255,0.25) !important;
    transform: none !important;
}

.stButton > button {
    background: hsl(229, 84%, 39%) !important;
    color: white !important; border: none !important;
    border-radius: 8px !important; padding: 0.55rem 1.4rem !important;
    font-weight: 600 !important; font-size: 0.875rem !important;
    transition: background 0.15s ease, transform 0.1s ease !important;
    box-shadow: 0 1px 4px rgba(26,58,199,0.25) !important;
}
.stButton > button:hover { background: hsl(229,84%,30%) !important; transform: translateY(-1px) !important; }
.stButton > button:disabled { background: hsl(220,13%,75%) !important; transform: none !important; }

.stDownloadButton > button {
    background: white !important; color: hsl(229,84%,39%) !important;
    border: 1.5px solid hsl(229,84%,39%) !important; border-radius: 8px !important;
    font-weight: 600 !important; transition: all 0.15s ease !important;
}
.stDownloadButton > button:hover { background: hsl(229,84%,96%) !important; transform: translateY(-1px) !important; }

[data-testid="stFileUploaderDropzone"] {
    border: 2px dashed hsl(229,84%,72%) !important;
    border-radius: 10px !important; background: hsl(229,84%,97%) !important;
}
[data-testid="stFileUploaderDropzone"]:hover { border-color: hsl(229,84%,39%) !important; }

.stTabs [data-baseweb="tab-list"] { gap: 4px; border-bottom: 2px solid hsl(220,13%,91%); background: transparent; }
.stTabs [data-baseweb="tab"] {
    border-radius: 8px 8px 0 0 !important; padding: 8px 20px !important;
    font-weight: 500 !important; color: hsl(215,16%,47%) !important;
    border: none !important; background: transparent !important;
}
.stTabs [aria-selected="true"] {
    color: hsl(229,84%,39%) !important;
    border-bottom: 2px solid hsl(229,84%,39%) !important; font-weight: 700 !important;
}

.stTextInput input, [data-testid="stNumberInput"] input {
    border-radius: 8px !important; border-color: hsl(220,13%,85%) !important;
}
.stTextInput input:focus, [data-testid="stNumberInput"] input:focus {
    border-color: hsl(229,84%,39%) !important; box-shadow: 0 0 0 3px hsl(229,84%,92%) !important;
}

.streamlit-expanderHeader { border-radius: 8px !important; background: hsl(229,84%,97%) !important; font-weight: 600 !important; }
[data-testid="stDataFrame"] { border-radius: 10px !important; overflow: hidden; border: 1px solid hsl(220,13%,91%) !important; }
[data-testid="stAlert"] { border-radius: 10px !important; }

.rexel-page-header {
    background: linear-gradient(135deg, hsl(229,84%,36%), hsl(229,84%,50%));
    color: white; padding: 22px 28px; border-radius: 14px; margin-bottom: 24px;
}
.rexel-page-header h1 { font-size:1.55rem; font-weight:700; margin:0 0 4px 0; color:white !important; }
.rexel-page-header p  { font-size:0.87rem; opacity:0.85; margin:0; color:white !important; }

.rexel-metric {
    background: white; border-radius: 12px; padding: 18px 20px;
    border: 1px solid hsl(220,13%,91%); box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    display: flex; align-items: center; justify-content: space-between; gap: 12px; height: 100%;
}
.rexel-metric-label { font-size:.72rem; font-weight:600; color:hsl(215,16%,47%); text-transform:uppercase; letter-spacing:.05em; }
.rexel-metric-title { font-size:.95rem; font-weight:700; color:hsl(215,28%,17%); margin:2px 0; }
.rexel-metric-desc  { font-size:.75rem; color:hsl(215,16%,57%); }
.rexel-metric-icon  { width:42px; height:42px; border-radius:10px; background:hsl(229,84%,95%);
                       display:flex; align-items:center; justify-content:center; font-size:1.25rem; flex-shrink:0; }

.rexel-tool-card {
    background: white; border-radius: 12px; padding: 20px;
    border: 1px solid hsl(220,13%,91%); box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    transition: box-shadow .2s ease, transform .1s ease; margin-bottom: 4px;
}
.rexel-tool-card:hover { box-shadow: 0 4px 14px rgba(0,0,0,0.09); transform: translateY(-2px); }
.rexel-tool-icon { width:38px; height:38px; border-radius:9px; background:hsl(229,84%,95%);
                    display:inline-flex; align-items:center; justify-content:center; font-size:1.1rem; }
.rexel-tool-title { font-size:1rem; font-weight:700; color:hsl(215,28%,17%); margin:0; }
.rexel-tool-desc  { font-size:.82rem; color:hsl(215,16%,47%); margin:8px 0 14px 0; line-height:1.5; }
.badge-active { display:inline-block; padding:2px 8px; border-radius:20px; font-size:.68rem;
                font-weight:700; letter-spacing:.04em; background:#dcfce7; color:#166534; }

.sidebar-logo { display:flex; align-items:center; gap:10px;
                padding:14px 18px 18px 18px; border-bottom:1px solid rgba(255,255,255,.18); margin-bottom:10px; }
.sidebar-logo-icon { width:34px; height:34px; background:white; border-radius:8px;
                      display:flex; align-items:center; justify-content:center;
                      font-weight:800; font-size:.95rem; color:hsl(229,84%,39%); flex-shrink:0; }
.sidebar-logo-title { margin:0; font-size:.9rem; font-weight:700; color:white !important; line-height:1.2; }
.sidebar-logo-sub   { margin:0; font-size:.68rem; color:rgba(255,255,255,.72) !important; }
</style>
""", unsafe_allow_html=True)


# ══ CACHE WRAPPER (lecture fichiers) ═════════════════════════════════════════
@st.cache_data(show_spinner=False)
def _cached_read(file_bytes: bytes, filename: str) -> pd.DataFrame:
    return io_service.read_file(file_bytes, filename)


def read_any(upload) -> pd.DataFrame:
    return _cached_read(upload.getvalue(), upload.name)


# ══ UI HELPERS ════════════════════════════════════════════════════════════════

def page_header(title: str, subtitle: str) -> None:
    st.markdown(f"""
    <div class="rexel-page-header">
        <h1>{title}</h1>
        <p>{subtitle}</p>
    </div>""", unsafe_allow_html=True)


def _render_downloads(tool_id: str) -> None:
    store = _repo.list_files(tool_id)
    if not store:
        return
    items = list(store.values())
    cols = st.columns(min(len(items), 3))
    for i, info in enumerate(items):
        with cols[i % len(cols)]:
            st.download_button(
                info["label"], info["data"],
                file_name=info["filename"], mime=info["mime"],
                key=f"dl__{tool_id}__{info['filename']}",
                use_container_width=True,
            )
    if len(items) > 1:
        folder = st.text_input("📁 Nom du dossier ZIP", value=tool_id.upper(),
                               key=f"{tool_id}__zip_name")
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for info in items:
                zf.writestr(os.path.join(folder, info["filename"]), info["data"])
        buf.seek(0)
        st.download_button(
            "📦 Télécharger tous les fichiers (.zip)", buf.getvalue(),
            file_name=f"{folder}.zip", mime="application/zip",
            key=f"zip__{tool_id}", use_container_width=True,
        )


def _render_df(tool_id: str, key: str = "result", rows: int = 5) -> None:
    df = _repo.load_dataframe(tool_id, key)
    if df is not None:
        with st.expander(f"📊 Aperçu ({len(df):,} lignes)", expanded=False):
            st.dataframe(df.head(rows), use_container_width=True)


def _preview_file(upload) -> None:
    try:
        df = read_any(upload)
    except Exception as e:
        st.error(f"{upload.name} – lecture impossible : {e}")
        return
    with st.expander(f"📄 Aperçu – {upload.name}", expanded=False):
        st.dataframe(df.head(), use_container_width=True)
        meta = pd.DataFrame({"N°": range(1, len(df.columns)+1), "Colonne": df.columns})
        st.table(meta)


def _uploader_state(prefix: str, lots: dict) -> None:
    for key in lots:
        st.session_state.setdefault(f"{prefix}_{key}_files", [])
        st.session_state.setdefault(f"{prefix}_{key}_names", [])

    cols = st.columns(len(lots))
    for (key, (title, lab_ref, lab_val)), col in zip(lots.items(), cols):
        with col:
            st.markdown(f"**{title}**")
            uploads = st.file_uploader(
                "Déposer votre fichier", type=("csv", "xlsx", "xls"),
                accept_multiple_files=True, key=f"{prefix}_{key}_up",
                label_visibility="collapsed",
            )
            if uploads:
                new = [u for u in uploads
                       if u.name not in st.session_state[f"{prefix}_{key}_names"]]
                if new:
                    st.session_state[f"{prefix}_{key}_files"] += new
                    st.session_state[f"{prefix}_{key}_names"] += [u.name for u in new]
                    st.success(f"✓ {len(new)} fichier(s) ajouté(s)")
                    for up in new:
                        _preview_file(up)
            st.number_input(lab_ref, 1, 50, 1, key=f"{prefix}_{key}_ref",
                            help="Index de la colonne référence produit")
            st.number_input(lab_val, 1, 50, 2, key=f"{prefix}_{key}_val",
                            help="Index de la colonne code M2")
            n = len(st.session_state[f"{prefix}_{key}_files"])
            st.caption(f"📎 {n} fichier(s) · RAM {_ram()}")


def reset_page() -> None:
    _repo.clear()
    st.rerun()


# ══ PAGES ═════════════════════════════════════════════════════════════════════

def page_dashboard() -> None:
    page_header("🛠 Rexel Multi-Outils B2B",
                "Tableau de bord — sélectionnez un outil dans le menu latéral")

    tools_meta = [
        ("🔄", "Mise à jour M2",      "Mappe les codes Mach_2 entre plan N‑1 et N",     "Mise à jour Mach_2"),
        ("🧩", "Classification Code", "Génère DFRXHYBRCMR et AFRXHYBRCMR",              "Classification Code"),
        ("📦", "Multiconnexion",       "Fichiers PF1–PF6 pour création de comptes B2B",  "Multiconnexion"),
        ("🗂️", "Personal Catalogue",  "Produit DFRXHYBRPCP, AFRXHYBRCMP et les ACK",    "Personal Catalogue"),
        ("📑", "CPN",                  "Produit cartésien Référence × Comptes → DFRX",   "CPN"),
    ]

    cols = st.columns(5)
    for col, (icon, name, desc, _) in zip(cols, tools_meta):
        with col:
            st.markdown(f"""
            <div class="rexel-metric">
                <div>
                    <div class="rexel-metric-label">Outil</div>
                    <div class="rexel-metric-title">{name}</div>
                    <div class="rexel-metric-desc">{desc}</div>
                </div>
                <div class="rexel-metric-icon">{icon}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
    st.markdown("### Outils disponibles")

    rows_of_3 = [tools_meta[i:i+3] for i in range(0, len(tools_meta), 3)]
    for row in rows_of_3:
        cols = st.columns(3)
        for col, (icon, name, desc, nav_key) in zip(cols, row):
            with col:
                st.markdown(f"""
                <div class="rexel-tool-card">
                    <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px">
                        <div class="rexel-tool-icon">{icon}</div>
                        <div>
                            <div class="rexel-tool-title">{name}</div>
                            <span class="badge-active">Actif</span>
                        </div>
                    </div>
                    <div class="rexel-tool-desc">{desc}</div>
                </div>""", unsafe_allow_html=True)
                if st.button("Ouvrir", key=f"dash_{nav_key}", use_container_width=True):
                    st.session_state["nav_main"] = nav_key
                    st.rerun()


def page_update_m2() -> None:
    page_header("🔄 Mise à jour des codes Mach_2",
                "Mappez les anciens codes M2 vers les nouveaux via les plans d'offre N‑1 et N")

    tab_pc, tab_cli = st.tabs(["📂 Personal Catalogue", "🤝 Appairage Client"])

    with tab_pc:
        LOTS_PC = {
            "old": ("Ancien plan d'offre", "Colonne Référence produit", "Colonne Ancien M2"),
            "new": ("Nouveau plan d'offre", "Colonne Référence produit", "Colonne Nouveau M2"),
        }
        _uploader_state("pc", LOTS_PC)

        if st.button("🚀 Générer le fichier", key="pc_generate"):
            if not all(st.session_state[f"pc_{k}_files"] for k in LOTS_PC):
                st.warning("Chargez à la fois les fichiers N‑1 **et** N.")
                st.stop()
            with st.spinner("Calcul en cours…"):
                old_df = pd.concat([read_any(f) for f in st.session_state["pc_old_files"]],
                                   ignore_index=True).drop_duplicates()
                new_df = pd.concat([read_any(f) for f in st.session_state["pc_new_files"]],
                                   ignore_index=True).drop_duplicates()
                maj_df = m2_service.build_m2_update(
                    old_df, new_df,
                    st.session_state["pc_old_ref"] - 1, st.session_state["pc_old_val"] - 1,
                    st.session_state["pc_new_ref"] - 1, st.session_state["pc_new_val"] - 1,
                )
            _repo.save_dataframe("pc_maj", "result", maj_df)
            _repo.save_file("pc_maj", f"M2_MisAJour_{_today()}.csv",
                            maj_df.to_csv(index=False, sep=";").encode(),
                            "text/csv", "⬇️ M2_MisAJour.csv")
            st.success(f"✓ {len(maj_df):,} codes traités")

        _render_downloads("pc_maj")
        _render_df("pc_maj")

    with tab_cli:
        LOTS_CL = {
            "old": ("Ancien plan d'offre",  "Colonne Référence produit", "Colonne Ancien M2"),
            "new": ("Nouveau plan d'offre",  "Colonne Référence produit", "Colonne Nouveau M2"),
            "map": ("Appairage Client",      "Colonne Ancien M2",         "Colonne Code famille client"),
        }
        _uploader_state("cl", LOTS_CL)

        if not st.session_state.get("cl_cols") and st.session_state.get("cl_new_files"):
            cols_new = []
            for f in st.session_state["cl_new_files"]:
                cols_new += read_any(f).columns.tolist()
            st.session_state["cl_cols"] = sorted(set(cols_new))

        extra_cols = st.multiselect(
            "Colonnes additionnelles pour `a_remplir.csv`",
            options=st.session_state.get("cl_cols", []),
        )

        if st.button("🚀 Générer les fichiers d'appairage", key="cl_generate"):
            if not all(st.session_state[f"cl_{k}_files"] for k in LOTS_CL):
                st.warning("Chargez les **3** jeux de données.")
                st.stop()
            if (st.session_state["cl_old_ref"] == st.session_state["cl_old_val"] or
                    st.session_state["cl_new_ref"] == st.session_state["cl_new_val"]):
                st.error("Référence produit et M2 doivent être deux colonnes différentes.")
                st.stop()

            with st.spinner("Calcul en cours…"):
                def _concat(key):
                    return pd.concat([read_any(f) for f in st.session_state[f"cl_{key}_files"]],
                                     ignore_index=True).drop_duplicates()
                appair_df, missing_df = m2_service.build_appairage(
                    _concat("old"), _concat("new"), _concat("map"),
                    st.session_state["cl_old_ref"] - 1, st.session_state["cl_old_val"] - 1,
                    st.session_state["cl_new_ref"] - 1, st.session_state["cl_new_val"] - 1,
                    st.session_state["cl_map_ref"] - 1, st.session_state["cl_map_val"] - 1,
                    extra_cols,
                )
            dstr = _today()
            _repo.save_dataframe("cl", "result", appair_df)
            _repo.save_file("cl", f"appairage_M2_CodeFamilleClient_{dstr}.csv",
                            appair_df.to_csv(index=False, sep=";").encode(),
                            "text/csv", "⬇️ appairage_M2_famille.csv")
            _repo.save_file("cl", f"a_remplir_{dstr}.csv",
                            missing_df.to_csv(index=False, sep=";").encode(),
                            "text/csv", "⬇️ a_remplir.csv")
            st.success(f"✓ {len(appair_df):,} codes mappés · {len(missing_df):,} sans famille")

        _render_downloads("cl")
        _render_df("cl")


def page_classification() -> None:
    page_header("🧩 Classification Code",
                "Génère DFRXHYBRCMR et AFRXHYBRCMR à partir d'un appairage M2 → famille client")

    pair_file = st.file_uploader("📄 Fichier d'appairage", type=("csv", "xlsx", "xls"))
    if not pair_file:
        st.info("Chargez le fichier d'appairage pour continuer.")
        _render_downloads("cc")
        _render_df("cc")
        st.stop()

    try:
        pair_df = read_any(pair_file)
    except Exception as e:
        st.error(f"Impossible de lire le fichier : {e}")
        st.stop()

    with st.expander("📊 Aperçu du fichier chargé"):
        st.dataframe(pair_df.head(), use_container_width=True)

    max_cols = len(pair_df.columns)
    col1, col2, col3 = st.columns(3)
    with col1: idx_m2  = st.number_input("🔢 Index colonne Code M2",         1, max_cols, 1)
    with col2: idx_fam = st.number_input("🔢 Index colonne Code famille",     1, max_cols, 2)
    with col3: entreprise = st.text_input("🏢 Entreprise")

    if st.button("🚀 Générer les fichiers", key="class_generate"):
        if not entreprise:
            st.warning("Renseignez le champ Entreprise.")
            st.stop()
        try:
            with st.spinner("Génération…"):
                dstr = _today()
                df_out, dfrx_bytes, afrx_ack = classification_service.build_classification_output(
                    pair_df, int(idx_m2), int(idx_fam), entreprise, dstr
                )
        except (ValueError, IndexError) as e:
            st.error(str(e))
            st.stop()

        _repo.save_dataframe("cc", "result", df_out)
        _repo.save_file("cc", f"DFRXHYBRCMR{dstr}0000", dfrx_bytes,
                        "text/tab-separated-values", "📥 DFRXHYBRCMR")
        _repo.save_file("cc", f"AFRXHYBRCMR{dstr}0000", afrx_ack.encode(),
                        "text/plain", "📥 AFRXHYBRCMR")
        st.success(f"✓ {len(df_out):,} lignes générées")

    _render_downloads("cc")
    _render_df("cc")


def page_multiconnexion() -> None:
    page_header("📦 Multiconnexion",
                "Génère les fichiers PF1–PF6 pour la création de comptes B2B (OCI ou cXML)")

    integration_type = st.radio("Type d'intégration", ["cXML", "OCI"], horizontal=True)

    with st.expander("📑 Télécharger le template Multiconnexion"):
        cols_tpl = ["Numéro de compte", "Raison sociale", "Adresse", "Code agence"]
        buf_tpl  = io.BytesIO()
        pd.DataFrame([{c: "" for c in cols_tpl}]).to_excel(buf_tpl, index=False)
        buf_tpl.seek(0)
        st.download_button("📥 Template Excel", buf_tpl.getvalue(),
                           file_name="dfrecu_template.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="multi_template")

    up_file = st.file_uploader("📄 Fichier Multiconnexion", type=("csv", "xlsx", "xls"))
    if not up_file:
        _render_downloads("multi")
        _render_df("multi")
        st.stop()

    col1, col2 = st.columns(2)
    with col1: entreprise    = st.text_input("🏢 Entreprise").strip()
    with col2: punchout_user = st.text_input("👤 punchoutUserID")

    domain = identity = ""
    if integration_type == "cXML":
        col3, col4 = st.columns(2)
        with col3: domain   = st.selectbox("🌐 Domain", ["NetworkID", "DUNS"])
        with col4: identity = st.text_input("🆔 Identity")

    col5, col6 = st.columns(2)
    with col5: vm_choice  = st.radio("ViewMasterCatalog", ["True", "False"], horizontal=True)
    with col6: pc_enabled = st.radio("Personal Catalogue ?", ["True", "False"], horizontal=True)

    pc_name = ""
    if pc_enabled == "True":
        pc_name = st.text_input("Nom du catalogue (sans PC_)", placeholder="CATALOGUE").strip()

    if st.button("🚀 Générer les fichiers", key="multi_generate"):
        base_ok = all([entreprise, punchout_user, (pc_enabled == "False" or bool(pc_name))])
        cx_ok   = all([domain, identity]) if integration_type == "cXML" else True
        if not (base_ok and cx_ok):
            st.warning("Remplissez tous les champs requis.")
            st.stop()

        try:
            df_src = read_any(up_file)
        except Exception as e:
            st.error(f"Lecture impossible : {e}")
            st.stop()

        CANONICAL = {
            "numéro de compte": "Numéro de compte", "code agence":    "Code agence",
            "raison sociale":   "Raison sociale",    "adresse":        "Adresse",
        }
        norm_map = {c: c.strip().lower() for c in df_src.columns}
        missing  = [CANONICAL[lc] for lc in CANONICAL if lc not in norm_map.values()]
        if missing:
            st.error(f"Colonnes manquantes : {', '.join(missing)}")
            st.stop()

        rename = {orig: CANONICAL[norm] for orig, norm in norm_map.items() if norm in CANONICAL}
        df_src.rename(columns=rename, inplace=True)

        df_src["Numéro de compte"], bad_acc = sanitize_numeric(df_src["Numéro de compte"], 7)
        df_src["Code agence"],      bad_ag  = sanitize_numeric(df_src["Code agence"], 4)
        if bad_acc.any():
            st.error(f"{bad_acc.sum()} Numéro(s) de compte invalide(s).")
            st.dataframe(df_src.loc[bad_acc, "Numéro de compte"].to_frame(), use_container_width=True)
            st.stop()
        if bad_ag.any():
            st.error(f"{bad_ag.sum()} Code(s) agence invalide(s).")
            st.dataframe(df_src.loc[bad_ag, "Code agence"].to_frame(), use_container_width=True)
            st.stop()

        with st.spinner("Construction des tables PF…"):
            try:
                tables = multiconnexion_service.build_tables(
                    df_src, entreprise=entreprise, view_master_catalog=vm_choice,
                    punchout_user_id=punchout_user, domain=domain, identity=identity,
                    integration_type=integration_type,
                )
            except ValueError as e:
                st.error(str(e))
                st.stop()

        labels   = ["PF1", "PF2", "PF3", "PF4", "PF5"] + (["PF6"] if integration_type == "cXML" else [])
        file_map = {
            "PF1": f"B2B Units creation_{entreprise}.xlsx",
            "PF2": f"Table_chargement_adresse_{entreprise}.xlsx",
            "PF3": f"Table_PunchoutAccountAndBranchAssociation_{entreprise}.xlsx",
            "PF4": f"PunchoutBranchAliasAssociation_{entreprise}.xlsx",
            "PF5": f"Table_Attach_B2BUnitstoUsers_{entreprise}.xlsx",
            "PF6": f"PunchoutAccountSetup_{entreprise}.xlsx",
        }
        XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        for lbl, df in zip(labels, tables):
            _repo.save_file("multi", file_map[lbl], to_xlsx(df), XLSX_MIME, f"⬇️ {lbl}")
        _repo.save_dataframe("multi", "result", tables[0])
        st.success(f"✓ {len(labels)} fichiers générés pour {len(df_src):,} comptes")

    _render_downloads("multi")
    _render_df("multi")

    if IS_OUTLOOK and _repo.list_files("multi"):
        st.markdown("---")
        dest = st.text_input("Destinataire (Outlook)")
        if st.button("Ouvrir un brouillon Outlook", key="multi_outlook"):
            files_att = [(i["filename"], i["data"])
                         for i in _repo.list_files("multi").values()]
            _create_outlook_draft(files_att, to_=dest,
                                  subject=f"Fichiers PF – {entreprise} ({datetime.now():%Y-%m-%d %H:%M})")
            st.success("Brouillon Outlook ouvert.")
    elif not IS_OUTLOOK:
        st.caption("ℹ️ Automatisation Outlook non disponible sur cet environnement.")


def page_dfrx_pc() -> None:
    page_header("🗂️ Personal Catalogue",
                "Génère DFRXHYBRPCP, DFRXHYBRCMP et les ACK associés")
    nav = st.radio("Mode", ["Sans mise à jour Mach_2", "Avec mise à jour Mach_2"], horizontal=True)
    st.markdown("---")
    if nav == "Sans mise à jour Mach_2":
        _generator_pc()
    else:
        _generator_maj_m2()


def _generator_pc() -> None:
    col_a, col_b = st.columns(2)
    with col_a:
        codes_file = st.file_uploader("📄 Fichier codes Mach_2", type=("csv", "xlsx", "xls"), key="pc_codes")
        if codes_file: _preview_file(codes_file)
    with col_b:
        compte_file = st.file_uploader("📄 Fichier numéros de compte", type=("csv", "xlsx", "xls"), key="pc_comptes")
        if compte_file: _preview_file(compte_file)

    if not (codes_file and compte_file):
        _render_downloads("gen_pc")
        _render_df("gen_pc")
        return

    df_codes_meta = read_any(codes_file)
    df_comp_meta  = read_any(compte_file)
    col1, col2, col3, col4 = st.columns(4)
    with col1: ci = st.number_input("Colonne codes M2", 1, len(df_codes_meta.columns), 1, key="gpc_ci")
    with col2: cc = st.number_input("Colonne comptes",  1, len(df_comp_meta.columns),  1, key="gpc_cc")
    with col3: entreprise = st.text_input("🏢 Entreprise", key="gpc_ent")
    with col4: statut = st.selectbox("📌 Statut", ["", "INCLUDE", "EXCLUDE"], key="gpc_stat")

    if st.button("🚀 Générer PC", key="gen_pc_generate"):
        if not all([entreprise, statut]):
            st.warning("Renseignez l'entreprise et le statut.")
            st.stop()
        df_codes   = read_any(codes_file)
        df_comptes = read_any(compte_file)
        raw_codes  = df_codes.iloc[:, ci-1].dropna().astype(str).str.strip()
        comptes    = df_comptes.iloc[:, cc-1].dropna().astype(str).str.strip()
        sanitized  = raw_codes.apply(sanitize_code)
        if sanitized.isna().any():
            st.error(f"{sanitized.isna().sum()} code(s) M2 invalide(s).")
            st.dataframe(raw_codes[sanitized.isna()].to_frame("Code fourni"), use_container_width=True)
            st.stop()
        with st.spinner("Génération…"):
            dstr    = _today()
            df1     = pc_service.build_pc_profile(sanitized, entreprise, statut)
            files   = pc_service.build_pc_files_data(df1, comptes, entreprise, dstr)
        _repo.save_dataframe("gen_pc", "result", df1)
        _repo.save_files_dict("gen_pc", files)
        st.success(f"✓ {len(df1):,} codes · {len(comptes):,} comptes")

    _render_downloads("gen_pc")
    _render_df("gen_pc")


def _generator_maj_m2() -> None:
    col_a, col_b, col_c = st.columns(3)
    with col_a: codes_file  = st.file_uploader("📄 Fichier codes Mach_2",    type=("csv","xlsx","xls"), key="maj_codes")
    with col_b: compte_file = st.file_uploader("📄 Fichier numéros de compte",type=("csv","xlsx","xls"), key="maj_comp")
    with col_c: map_file    = st.file_uploader("📄 Fichier M2_MisAJour",      type=("csv","xlsx","xls"), key="maj_map")

    if not (codes_file and compte_file and map_file):
        st.info("Chargez les 3 fichiers pour continuer.")
        _render_downloads("majm2")
        return

    col1, col2, col3, col4 = st.columns(4)
    with col1: ci_codes = st.number_input("Colonne M2",        1, 50, 1, key="maj_ci")
    with col2: ci_comp  = st.number_input("Colonne comptes",   1, 50, 1, key="maj_cc")
    with col3: ci_old   = st.number_input("Colonne M2 ancien", 1, 50, 1, key="maj_old")
    with col4: ci_new   = st.number_input("Colonne M2 nouveau",1, 50, 2, key="maj_new")
    col5, col6 = st.columns(2)
    with col5: entreprise = st.text_input("🏢 Entreprise", key="maj_ent")
    with col6: statut = st.selectbox("📌 Statut", ["", "INCLUDE", "EXCLUDE"], key="maj_stat")

    if st.button("🚀 Générer MàJ", key="majm2_generate"):
        if not all([entreprise, statut]):
            st.warning("Renseignez l'entreprise et le statut.")
            st.stop()
        df_codes   = read_any(codes_file)
        df_comptes = read_any(compte_file)
        df_map     = read_any(map_file)
        raw_codes  = df_codes.iloc[:, ci_codes-1].dropna().astype(str).str.strip()
        comptes    = df_comptes.iloc[:, ci_comp-1].dropna().astype(str).str.strip()
        sanitized  = raw_codes.apply(sanitize_code)
        if sanitized.isna().any():
            st.error("Codes M2 invalides détectés.")
            st.dataframe(raw_codes[sanitized.isna()].to_frame("Code fourni"), use_container_width=True)
            st.stop()
        old_codes = df_map.iloc[:, ci_old-1].astype(str).apply(sanitize_code)
        new_codes = df_map.iloc[:, ci_new-1].astype(str).apply(sanitize_code)
        mapping   = (pd.DataFrame({"old": old_codes, "new": new_codes})
                     .dropna().drop_duplicates("old").set_index("old")["new"].to_dict())
        updated = sanitized.map(lambda c: mapping.get(c, c))
        with st.spinner("Génération…"):
            dstr  = _today()
            df1   = pc_service.build_pc_profile(updated, entreprise, statut)
            files = pc_service.build_pc_files_data(df1, comptes, entreprise, dstr)
        _repo.save_dataframe("majm2", "result", df1)
        _repo.save_files_dict("majm2", files)
        st.success(f"✓ {len(df1):,} codes · {len(mapping):,} substitutions appliquées")

    _render_downloads("majm2")
    _render_df("majm2")


def page_cpn() -> None:
    page_header("📑 CPN",
                "Génère DFRXHYBCPNA — produit cartésien Référence interne × Numéros de compte")

    MAX_ROWS = 500_000

    col_a, col_b = st.columns(2)
    with col_a:
        main_file = st.file_uploader("📄 Appairage Code produit / Référence interne",
                                     type=("csv","xlsx","xls"), key="cpn_main")
        if main_file: _preview_file(main_file)
    with col_b:
        cli_file = st.file_uploader("📄 Périmètre (numéros de compte)",
                                    type=("csv","xlsx","xls"), key="cpn_cli")
        if cli_file: _preview_file(cli_file)

    if not (main_file and cli_file):
        _render_downloads("cpn")
        _render_df("cpn")
        st.stop()

    df_main = read_any(main_file)
    df_cli  = read_any(cli_file)

    col1, col2, col3 = st.columns(3)
    with col1: col_int      = st.selectbox("Référence interne (8 chiffres)",
                                           range(1, len(df_main.columns)+1), index=0, key="cpn_ci")
    with col2: col_cli_prod = st.selectbox("Code produit client",
                                           range(1, len(df_main.columns)+1),
                                           index=min(1, len(df_main.columns)-1), key="cpn_cp")
    with col3: col_cli_acc  = st.selectbox("Numéro de compte (périmètre)",
                                           range(1, len(df_cli.columns)+1), index=0, key="cpn_ca")

    n_prod  = len(df_main.iloc[:, col_int-1].dropna())
    n_acc   = len(df_cli.iloc[:, col_cli_acc-1].dropna())
    n_total = n_prod * n_acc
    st.caption(f"📐 Produit cartésien estimé : **{n_total:,}** lignes ({n_prod} produits × {n_acc} comptes)")
    if n_total > MAX_ROWS:
        st.warning(f"⚠️ Volume trop grand ({n_total:,} > {MAX_ROWS:,} max). Réduisez le périmètre.")

    if st.button("🚀 Générer CPN", key="cpn_generate", disabled=(n_total > MAX_ROWS)):
        series_int      = df_main.iloc[:, col_int-1].astype(str).str.strip().reset_index(drop=True)
        series_cli_prod = df_main.iloc[:, col_cli_prod-1].astype(str).str.strip().reset_index(drop=True)
        series_cli_acc  = df_cli.iloc[:,  col_cli_acc-1].astype(str).str.strip().reset_index(drop=True)

        invalid = ~series_int.str.fullmatch(r"\d{8}")
        if invalid.any():
            st.error(f"{invalid.sum()} référence(s) invalide(s) — 8 chiffres attendus.")
            st.dataframe(series_int[invalid].to_frame("Référence"), use_container_width=True)
            st.stop()

        with st.spinner(f"Calcul du produit cartésien ({n_total:,} lignes)…"):
            df_out = cpn_service.build_cpn(series_int, series_cli_prod, series_cli_acc)

        dstr = _today()
        _repo.save_dataframe("cpn", "result", df_out)
        _repo.save_file("cpn", f"DFRXHYBCPNA{dstr}0000",
                        df_out.to_csv(sep="\t", index=False, header=False).encode(),
                        "text/tab-separated-values", "⬇️ DFRX (TSV)")
        _repo.save_file("cpn", f"AFRXHYBCPNA{dstr}0000",
                        cpn_service.build_cpn_ack(dstr).encode(),
                        "text/plain", "⬇️ AFRX (TXT)")
        st.success(f"✓ {len(df_out):,} lignes générées")

    _render_downloads("cpn")
    _render_df("cpn")


# ── Outlook helper (Windows uniquement) ──────────────────────────────────────
def _create_outlook_draft(att: List[Tuple[str, bytes]], to_: str, subject: str) -> None:
    if not IS_OUTLOOK:
        return
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To, mail.Subject = to_, subject
    mail.Body = "Bonjour,\n\nVeuillez trouver les fichiers PF en pièce jointe.\n"
    for name, data in att:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=name)
        tmp.write(data)
        tmp.close()
        mail.Attachments.Add(tmp.name)
    mail.Display()


# ══ NAVIGATION ════════════════════════════════════════════════════════════════

PAGES = {
    "🏠 Dashboard":        page_dashboard,
    "Mise à jour Mach_2":  page_update_m2,
    "Classification Code": page_classification,
    "Multiconnexion":      page_multiconnexion,
    "Personal Catalogue":  page_dfrx_pc,
    "CPN":                 page_cpn,
}

_inject_css()

with st.sidebar:
    st.markdown("""
    <div class="sidebar-logo">
        <div class="sidebar-logo-icon">R</div>
        <div>
            <p class="sidebar-logo-title">Rexel</p>
            <p class="sidebar-logo-sub">Multi-Outils B2B</p>
        </div>
    </div>""", unsafe_allow_html=True)

    if st.button("🔄 Réinitialiser", key="reset_btn"):
        reset_page()

    choice = st.radio(
        "Navigation", list(PAGES), index=0,
        key="nav_main", label_visibility="collapsed",
    )

PAGES[choice]()
