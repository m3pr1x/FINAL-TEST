# multi_tool_app.py
# ──────────────────────────────────────────────────────────
# Application « Boîte à outils » – 5 modules dans 1 seul Streamlit
# ──────────────────────────────────────────────────────────
from __future__ import annotations
import csv, io, re, tempfile, os, sys
from datetime import datetime
from itertools import product
from pathlib import Path
from typing import Dict, Tuple, List

import pandas as pd
import streamlit as st

# ═══════════ IMPORTS OPTIONNELS ═══════════
try:                                   # libpostal
    from postal.parser import parse_address        # type: ignore
    USE_POSTAL = True
except ImportError:
    USE_POSTAL = False

try:                                   # Outlook
    import win32com.client as win32                # type: ignore
    IS_OUTLOOK = True
except ImportError:
    IS_OUTLOOK = False

try:                                   # RAM
    import psutil                                  # type: ignore
    RAM = lambda: f"{psutil.Process().memory_info().rss/1_048_576:,.0f} Mo"
except ModuleNotFoundError:
    psutil = None
    RAM = lambda: "n/a"                            # type: ignore

TODAY = datetime.today().strftime("%y%m%d")
st.set_page_config("Boîte à outils", "🛠", layout="wide")

# ──────────────────────────── FONCTIONS UTILITAIRES GLOBALES ────────────────────────────
@st.cache_data(show_spinner=False, hash_funcs={io.BytesIO: lambda _: None})
def read_csv(buf: io.BytesIO) -> pd.DataFrame:
    """Lecture robuste CSV : détecte encodage & séparateur."""
    for enc in ("utf-8", "latin1", "cp1252"):
        buf.seek(0)
        try:
            sample = buf.read(2048).decode(enc, errors="ignore")
            sep = csv.Sniffer().sniff(sample, delimiters=";,|\t").delimiter
            buf.seek(0)
            return pd.read_csv(buf, sep=sep, encoding=enc, engine="python",
                               on_bad_lines="skip", dtype=str)
        except Exception:
            continue
    raise ValueError("CSV illisible (encodage ou séparateur)")

@st.cache_data(show_spinner=False, hash_funcs={io.BytesIO: lambda _: None})
def read_any(upload) -> pd.DataFrame:
    """Lecture CSV / Excel, avec caching."""
    name = upload.name.lower()
    if name.endswith(".csv"):
        return read_csv(io.BytesIO(upload.getvalue()))
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(upload, engine="openpyxl", dtype=str)
    raise ValueError("Extension non gérée")

def to_m2(s: pd.Series) -> pd.Series:
    return s.astype(str).str.zfill(6)

def sanitize_code(code: str) -> str | None:
    """Vérifie qu’un code M2 compte 6 chiffres (ou 5 → paddé à 6)."""
    s = str(code).strip()
    if not s.isdigit():
        return None
    if len(s) == 5:
        s = s.zfill(6)
    return s if len(s) == 6 else None

def sanitize_numeric(series: pd.Series, width: int) -> Tuple[pd.Series, pd.Series]:
    """Normalise champs numériques ; renvoie série nettoyée + masque invalides."""
    s = series.astype(str).str.strip()
    s_pad = s.apply(lambda x: x.zfill(width) if x.isdigit() and len(x) <= width else x)
    bad = ~s_pad.str.fullmatch(fr"\d{{{width}}}")
    return s_pad, bad

# ═══════════════════════════  PAGE 1 – MISE À JOUR M2 (PC & Appairage) ═══════════════════════════
def page_update_m2() -> None:
    """Outil unique : Mise à jour M2 + Appairage client."""
    st.header("🔄 Mise à jour des codes M2")
    tab_pc, tab_cli = st.tabs(["📂 Personal Catalogue", "🤝 Appairage client"])

    # ▸▸▸ Onglet 1 – Mise à jour (PC)
    with tab_pc:
        LOTS_PC = {
            "old": ("Données N‑1", "Ref produit", "M2 ancien"),
            "new": ("Données N"  , "Ref produit", "M2 nouveau"),
        }
        _uploader_state("pc", LOTS_PC)

        if st.button("Générer : M2_MisAJour.csv", key="btn_pc"):
            if not all(st.session_state[f"pc_{k}_files"] for k in LOTS_PC):
                st.warning("Chargez à la fois les fichiers N‑1 **et** N."); st.stop()
            maj_df = _build_m2_update("pc", LOTS_PC)
            st.download_button("⬇️ Télécharger M2_MisAJour.csv",
                               maj_df.to_csv(index=False, sep=";"),
                               file_name=f"M2_MisAJour_{TODAY}.csv",
                               mime="text/csv")
            st.dataframe(maj_df.head())

    # ▸▸▸ Onglet 2 – Appairage
    with tab_cli:
        LOTS_CL = {
            "old": ("Données N‑1", "Ref produit", "M2 ancien"),
            "new": ("Données N"  , "Ref produit", "M2 nouveau"),
            "map": ("Mapping"    , "M2 ancien",   "Code famille client"),
        }
        _uploader_state("cl", LOTS_CL)

        extra_cols = st.multiselect(
            "Colonnes additionnelles (pour « a_remplir.csv »)",
            options=st.session_state.get("cl_cols", []),
        )

        if st.button("Générer : fichiers d’appairage", key="btn_cl"):
            if not all(st.session_state[f"cl_{k}_files"] for k in LOTS_CL):
                st.warning("Chargez les **3** jeux de données (N‑1, N, Mapping)."); st.stop()

            # Garde‑fou : indices identiques
            if (st.session_state["cl_old_ref"] == st.session_state["cl_old_val"] or
                st.session_state["cl_new_ref"] == st.session_state["cl_new_val"]):
                st.error("« Ref produit » et « M2 » doivent être deux colonnes différentes.")
                st.stop()

            appair_df, missing_df = _build_appairage("cl", LOTS_CL, extra_cols)
            st.download_button("⬇️ appairage_M2_famille.csv",
                               appair_df.to_csv(index=False, sep=";"),
                               file_name=f"appairage_M2_CodeFamilleClient_{TODAY}.csv",
                               mime="text/csv")
            st.download_button("⬇️ a_remplir.csv",
                               missing_df.to_csv(index=False, sep=";"),
                               file_name=f"a_remplir_{TODAY}.csv",
                               mime="text/csv")
            st.dataframe(appair_df.head())

# ──────────────────────────── HELPERS (Mise à jour M2) ────────────────────────────
def _preview_file(upload) -> None:
    """Aperçu interactif : 5 lignes + liste des colonnes."""
    try:
        df = read_any(upload)
    except Exception as e:
        st.error(f"{upload.name} – lecture impossible : {e}")
        return
    with st.expander(f"📄 Aperçu – {upload.name}", expanded=False):
        st.dataframe(df.head())
        meta = pd.DataFrame({"N°": range(1, len(df.columns)+1),
                             "Nom de colonne": df.columns})
        st.table(meta)

def _uploader_state(prefix: str, lots: dict[str, tuple[str, str, str]]) -> None:
    """Widget upload + état mémoire + aperçu automatique."""
    for key in lots:
        st.session_state.setdefault(f"{prefix}_{key}_files", [])
        st.session_state.setdefault(f"{prefix}_{key}_names", [])

    cols = st.columns(len(lots))
    for (key, (title, lab_ref, lab_val)), col in zip(lots.items(), cols):
        with col:
            st.subheader(title)
            uploads = st.file_uploader("Déposer…", type=("csv", "xlsx"),
                                       accept_multiple_files=True,
                                       key=f"{prefix}_{key}_up")
            if uploads:
                new = [u for u in uploads
                       if u.name not in st.session_state[f"{prefix}_{key}_names"]]
                st.session_state[f"{prefix}_{key}_files"] += new
                st.session_state[f"{prefix}_{key}_names"] += [u.name for u in new]
                st.success(f"{len(new)} fichier(s) ajouté(s)")
                for up in new:            # ➜ aperçu instantané
                    _preview_file(up)

            st.number_input(lab_ref, 1, 50, 1,
                            key=f"{prefix}_{key}_ref",
                            help="Index de la colonne contenant la référence produit")
            st.number_input(lab_val, 1, 50, 2,
                            key=f"{prefix}_{key}_val",
                            help="Index de la colonne contenant le code M2")
            st.caption(f"{len(st.session_state[f'{prefix}_{key}_files'])} fichier(s) • RAM {RAM()}")

def _add_cols(df: pd.DataFrame, ref_i: int, m2_i: int,
              ref_label: str, m2_label: str) -> pd.DataFrame:
    sub = df.iloc[:, [ref_i-1, m2_i-1]].copy()
    sub.columns = [ref_label, m2_label]
    sub[m2_label] = to_m2(sub[m2_label])
    return sub

def _build_m2_update(prefix: str, lots: dict[str, tuple[str, str, str]]) -> pd.DataFrame:
    dfs = {k: pd.concat([read_any(f) for f in st.session_state[f"{prefix}_{k}_files"]],
                        ignore_index=True).drop_duplicates()
           for k in lots}

    old_df = _add_cols(dfs["old"], st.session_state[f"{prefix}_old_ref"],
                       st.session_state[f"{prefix}_old_val"], "Ref", "M2_ancien")
    new_df = _add_cols(dfs["new"], st.session_state[f"{prefix}_new_ref"],
                       st.session_state[f"{prefix}_new_val"], "Ref", "M2_nouveau")

    merged = new_df.merge(old_df[["Ref", "M2_ancien"]], on="Ref", how="left")
    return (merged.groupby("M2_nouveau")["M2_ancien"]
                  .agg(lambda s: s.value_counts().idxmax()
                       if s.notna().any() else pd.NA)
                  ).reset_index()

def _build_appairage(prefix: str, lots: dict[str, tuple[str, str, str]],
                     extra_cols: list[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    dfs = {k: pd.concat([read_any(f) for f in st.session_state[f"{prefix}_{k}_files"]],
                        ignore_index=True).drop_duplicates()
           for k in lots}

    old_df = _add_cols(dfs["old"], st.session_state[f"{prefix}_old_ref"],
                       st.session_state[f"{prefix}_old_val"],
                       "Ref", "M2_ancien")
    new_df = _add_cols(dfs["new"], st.session_state[f"{prefix}_new_ref"],
                       st.session_state[f"{prefix}_new_val"],
                       "Ref", "M2_nouveau")

    map_df = dfs["map"].iloc[:, [st.session_state[f"{prefix}_map_ref"]-1,
                                 st.session_state[f"{prefix}_map_val"]-1]].copy()
    map_df.columns = ["M2_ancien", "Code_famille_Client"]
    map_df["M2_ancien"] = to_m2(map_df["M2_ancien"])
    old_df["M2_ancien"] = to_m2(old_df["M2_ancien"])

    merged = (new_df.merge(old_df[["Ref", "M2_ancien"]], on="Ref", how="left")
                     .merge(map_df, on="M2_ancien", how="left"))
    st.session_state["cl_cols"] = list(merged.columns)

    fam = (merged.groupby("M2_nouveau")["Code_famille_Client"]
                 .agg(lambda s: s.value_counts().idxmax()
                      if s.notna().any() else pd.NA)
                 ).reset_index()

    missing = fam[fam["Code_famille_Client"].isna()].copy()
    if extra_cols:
        keep = [c for c in extra_cols if c in merged.columns]
        missing = missing.merge(merged[["M2_nouveau"] + keep].drop_duplicates(),
                                on="M2_nouveau", how="left")
    return fam, missing
# ═══════════════════  PAGE 2 – CLASSIFICATION CODE ═══════════════════
def page_classification():
    st.header("🧩 Classification Code")
    pair_file = st.file_uploader("1) Appairage M2 ➜ Code famille (CSV)", type="csv")
    if not pair_file:
        st.info("Commence par charger l'appairage M2."); st.stop()

    pair_df = read_csv(io.BytesIO(pair_file.getvalue()))
    exp_cols = {"M2", "Code_famille_Client"}
    if not exp_cols.issubset(pair_df.columns):
        st.error(f"Le fichier doit contenir : {exp_cols}"); st.stop()
    pair_df["M2"] = to_m2(pair_df["M2"])
    st.success(f"{len(pair_df)} lignes chargées")
    st.dataframe(pair_df.head())

    data_files = st.file_uploader("2) Fichiers à classifier (CSV/XLSX/XLS)",
                                  accept_multiple_files=True,
                                  type=("csv", "xlsx", "xls"))
    if not data_files:
        st.info("Ajoute un ou plusieurs fichiers à classifier."); st.stop()

    results = []
    for upl in data_files:
        df = read_any(upl)
        st.markdown(f"##### {upl.name}")
        cols = [f"{i+1} – {c}" for i, c in enumerate(df.columns)]
        idx = st.selectbox("Colonne M2", cols, key=f"m2col_{upl.name}")
        m2_col = df.columns[int(idx.split(' –')[0]) - 1]
        df["M2"] = to_m2(df[m2_col])
        merged = df.merge(pair_df[["M2", "Code_famille_Client"]], on="M2", how="left")
        st.write(f"→ {merged['Code_famille_Client'].notna().sum()} / {len(df)} lignes appariées")
        results.append(merged)
        with st.expander("Aperçu"):
            st.dataframe(merged.head())

    final = pd.concat(results, ignore_index=True)
    fname = f"DATA_CLASSIFIEE_{datetime.today().strftime('%y%m%d_%H%M%S')}.csv"
    st.download_button("⬇️ Télécharger les données classifiées", final.to_csv(index=False, sep=";"),
                       file_name=fname, mime="text/csv")
    st.success("Classification terminée !")

# ═══════════════════  PAGE 3 – PF1 → PF6 GENERATOR ═══════════════════
def page_multiconnexion():
    st.header("📦 Générateur PF1 → PF6 (Multiconnexion)")
    integration_type = st.radio("Type d’intégration", ["cXML", "OCI"], horizontal=True)

    st.markdown(
        "Téléchargez le template, remplissez‑le puis uploadez votre fichier.  \n"
        "Colonnes requises : **Numéro de compte** (7 chiffres), **Raison sociale**, "
        "**Adresse**, **ManagingBranch** (4 chiffres)."
    )

    # Template
    with st.expander("📑 Template dfrecu.xlsx"):
        tpl_cols = ["Numéro de compte", "Raison sociale", "Adresse", "ManagingBranch"]
        tpl_buf = io.BytesIO()
        pd.DataFrame([{c: "" for c in tpl_cols}]).to_excel(tpl_buf, index=False)
        tpl_buf.seek(0)
        st.download_button("📥 Télécharger le template", tpl_buf.getvalue(),
                           file_name="dfrecu_template.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    up_file = st.file_uploader("📄 Fichier dfrecu", type=("csv", "xlsx", "xls"))
    if not up_file:
        st.info("Charge un fichier dfrecu pour continuer."); st.stop()

    col1, col2, col3 = st.columns(3)
    with col1:  entreprise     = st.text_input("🏢 Entreprise").strip()
    with col2:  punchout_user  = st.text_input("👤 punchoutUserID")
    with col3:  domain         = st.selectbox("🌐 Domain", ["NetworkID", "DUNS"])
    identity      = st.text_input("🆔 Identity")
    vm_choice     = st.radio("ViewMasterCatalog", ["True", "False"], horizontal=True)
    pc_enabled    = st.radio("Personal Catalogue ?", ["True", "False"], horizontal=True)
    pc_name       = st.text_input("Nom du catalogue (sans PC_)", placeholder="CATALOGUE").strip() if pc_enabled == "True" else ""

    if st.button("🚀 Générer PF", key="btn_pf"):
        required = [entreprise, punchout_user, identity,
                    (pc_enabled == "False" or pc_name)]
        if not all(required):
            st.warning("Remplis tous les champs requis."); st.stop()

        df_src = read_any(up_file)
        if {"Numéro de compte", "ManagingBranch"} - set(df_src.columns):
            st.error("Colonnes manquantes dans le fichier."); st.stop()

        acc_series, bad_acc = sanitize_numeric(df_src["Numéro de compte"], 7)
        man_series, bad_man = sanitize_numeric(df_src["ManagingBranch"], 4)
        if bad_acc.any() or bad_man.any():
            st.error("Numéro de compte ou ManagingBranch invalide(s)."); st.stop()

        df_src["Numéro de compte"] = acc_series
        df_src["ManagingBranch"]   = man_series

        # build_tables vient du code original (non reproduit ici pour concision)
        tables: List[pd.DataFrame] = build_tables(df_src)  # type: ignore

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        labels = ["PF1", "PF2", "PF3", "PF4", "PF5"] + (["PF6"] if integration_type == "cXML" else [])
        files_bytes = {}
        for label, df in zip(labels, tables):
            data_bytes = to_xlsx(df)
            fname = f"{label}_{entreprise}_{ts}.xlsx"
            files_bytes[fname] = data_bytes
            st.download_button(f"⬇️ {label}", data=data_bytes, file_name=fname,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.success("✅ Fichiers prêts !")
        st.dataframe(tables[0].head())

        # Option Outlook
        st.markdown("---")
        st.subheader("📧 Exporter via Outlook Desktop")
        if IS_OUTLOOK:
            dest = st.text_input("Destinataire (optionnel)")
            subj = f"Fichiers PF – {entreprise} ({ts})"
            if st.button("Ouvrir un brouillon Outlook"):
                create_outlook_draft(list(files_bytes.items()), to_=dest, subject=subj)
                st.success("Brouillon Outlook ouvert.")
        else:
            st.info("Automatisation Outlook indisponible sur cet environnement.")

# ═══════════════════  PAGE 4 – DFRX/AFRX (PC & MàJ M2) ═══════════════════
def page_dfrx_pc():
    st.header("🛠️ Générateur PC + Mise à jour M2")
    nav2 = st.radio("Choisir l’outil", ["Générateur PC", "Mise à jour M2"], horizontal=True)
    if nav2 == "Générateur PC":
        generator_pc()
    else:
        generator_maj_m2()

# ► Sous‑composants simplifiés (extraits du code initial)
def generator_pc():
    st.subheader("Générateur PC")
    codes_file = st.file_uploader("Codes produit", type=("csv", "xlsx", "xls"))
    if not codes_file: st.stop()
    col_idx_codes = st.number_input("Colonne Codes M2 (1=A)", 1, 50, 1)
    compte_file = st.file_uploader("Numéros de compte", type=("csv", "xlsx", "xls"))
    if not compte_file: st.stop()
    col_idx_comptes = st.number_input("Colonne comptes (1=A)", 1, 50, 1)
    entreprise = st.text_input("Entreprise")
    statut     = st.selectbox("Statut", ["", "INCLUDE", "EXCLUDE"])
    if st.button("🚀 Générer PC"):
        df_codes   = read_any(codes_file)
        df_comptes = read_any(compte_file)
        codes_raw = df_codes.iloc[:, col_idx_codes-1].dropna()
        comptes   = df_comptes.iloc[:, col_idx_comptes-1].dropna().astype(str).str.strip()
        codes = codes_raw.astype(str).apply(sanitize_code)
        if codes.isna().any():
            st.error("Codes M2 invalides."); st.stop()
        dstr = TODAY
        df1 = pd.DataFrame({
            0: [f"PC_PROFILE_{entreprise}"] * len(codes),
            1: [statut] * len(codes),
            2: [None] * len(codes),
            3: [f"M2_{c}" for c in codes],
            4: ["frxProductCatallog:Online"] * len(codes),
        }).drop_duplicates()
        st.download_button("DFRXHYBRPCP", df1.to_csv(index=False, header=False, sep=";"),
                           file_name=f"DFRXHYBRPCP{dstr}0000", mime="text/plain")
        st.success("Fichiers générés.")

def generator_maj_m2():
    st.subheader("Mise à jour M2 avant génération")
    codes_file = st.file_uploader("Codes produit", type=("csv", "xlsx", "xls"), key="maj_codes")
    if not codes_file: st.stop()
    col_idx_codes = st.number_input("Colonne Codes M2 (1=A)", 1, 50, 1)
    compte_file = st.file_uploader("Numéros de compte", type=("csv", "xlsx", "xls"), key="maj_comptes")
    if not compte_file: st.stop()
    col_idx_comptes = st.number_input("Colonne comptes (1=A)", 1, 50, 1)
    map_file = st.file_uploader("Fichier M2_MisAJour", type=("csv", "xlsx", "xls"))
    if not map_file: st.stop()
    col_idx_old = st.number_input("Colonne M2 ancien", 1, 50, 1)
    col_idx_new = st.number_input("Colonne M2 nouveau", 1, 50, 2)
    entreprise = st.text_input("Entreprise")
    statut     = st.selectbox("Statut", ["", "INCLUDE", "EXCLUDE"])

    if st.button("🚀 Générer MàJ"):
        df_codes   = read_any(codes_file)
        df_comptes = read_any(compte_file)
        df_map     = read_any(map_file)
        raw_codes  = df_codes.iloc[:, col_idx_codes-1]
        comptes    = df_comptes.iloc[:, col_idx_comptes-1].dropna().astype(str).str.strip()
        codes      = raw_codes.astype(str).apply(sanitize_code)
        old_codes  = df_map.iloc[:, col_idx_old-1].astype(str).apply(sanitize_code)
        new_codes  = df_map.iloc[:, col_idx_new-1].astype(str).apply(sanitize_code)
        mapping = pd.Series(new_codes.values, index=old_codes).dropna().to_dict()
        updated = codes.map(lambda c: mapping.get(c, c))
        dstr = TODAY
        df1 = pd.DataFrame({
            0: [f"PC_PROFILE_{entreprise}"] * len(updated),
            1: [statut] * len(updated),
            2: [None] * len(updated),
            3: [f"M2_{c}" for c in updated],
            4: ["frxProductCatallog:Online"] * len(updated),
        }).drop_duplicates()
        st.download_button("DFRXHYBRPCP", df1.to_csv(index=False, header=False, sep=";"),
                           file_name=f"DFRXHYBRPCP{dstr}0000", mime="text/plain")
        st.success("Fichiers générés.")

# ═══════════════════  PAGE 5 – CPN GENERATOR ═══════════════════
def page_cpn():
    st.header("📑 Générateur CPN (DFRXHYBCPNA / AFRXHYBCPNA)")
    colA, colB = st.columns(2)
    with colA:
        main_file = st.file_uploader("Appairage client", type=("csv", "xlsx", "xls"))
    with colB:
        cli_file  = st.file_uploader("Périmètre (comptes client)", type=("csv", "xlsx", "xls"))
    if not (main_file and cli_file):
        st.stop()

    df_main = read_any(main_file)
    max_cols = len(df_main.columns)
    col_int = st.selectbox("Colonne Réf. interne", range(1, max_cols+1), 0)
    col_cli = st.selectbox("Colonne Réf. client", range(1, max_cols+1), 1 if max_cols > 1 else 0)

    if st.button("🚀 Générer CPN"):
        df_cli = read_any(cli_file)
        series_int = df_main.iloc[:, col_int-1].astype(str).str.strip()
        invalid = ~series_int.str.fullmatch(r"\d{8}")
        if invalid.any():
            st.error("Réf. interne invalide (doit contenir 8 chiffres).")
            st.dataframe(series_int[invalid]); st.stop()
        series_cli = df_cli.iloc[:, 0].astype(str).str.strip()
        pf = pd.DataFrame(product(series_int, series_cli),
                          columns=["1", "2"])
        pf["3"] = pf["1"]
        today = TODAY
        dfrx_name = f"DFRXHYBCPNA{today}0000"
        afrx_name = f"AFRXHYBCPNA{today}0000"
        afrx_txt = (f"DFRXHYBCPNA{today}000148250201IT"
                    f"DFRXHYBCPNA{today}CPNAHYBFRX                    OK000000")
        st.download_button("⬇️ DFRX (TSV)", pf.to_csv(sep="\t", index=False, header=False).encode(),
                           file_name=dfrx_name, mime="text/tab-separated-values")
        st.download_button("⬇️ AFRX (TXT)", afrx_txt, file_name=afrx_name, mime="text/plain")
        st.success("Fichiers générés.")
        st.dataframe(pf.head())

# ═══════════════════════════  MENU PRINCIPAL ═══════════════════════════
PAGES = {
    "Mise à jour M2": page_update_m2,
    "Classification Code": page_classification,
    "PF1 → PF6 Generator": page_multiconnexion,
    "Générateur PC / MàJ M2": page_dfrx_pc,
    "CPN Generator": page_cpn,
}
choice = st.sidebar.radio("Navigation", list(PAGES.keys()))
PAGES[choice]()
