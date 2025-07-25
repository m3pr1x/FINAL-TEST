# -*- coding: utf-8 -*-
"""
multi_tool_app.py – Application « Boîte à outils » (Streamlit)

Correctifs 07/2025
    • PF1→PF6 : noms explicites (plus de timestamp)
    • Générateur PC & MàJ M2 : ajout du fichier AFRXHYBRPCP<date>0000.txt
    • Correction d’une parenthèse non fermée (Outlook)
    • Clé 'nav_main' pour le menu (évite DuplicateElementId)
    • **Bouton « Réinitialiser la page » corrigé** → nouvelle fonction
      `reset_page()` qui vide `st.session_state` puis appelle `st.rerun()`.
      Dorénavant :
          st.button(..., on_click=reset_page)
      (`st.experimental_rerun()` était déprécié depuis Streamlit ≥ 1.27).

Améliorations 07/2025 — PERSISTANCE DES FICHIERS GÉNÉRÉS
    • Les objets créés (DataFrame / bytes) sont conservés dans st.session_state
      tant que l’utilisateur ne ré‑initialise pas la page.
    • Les st.download_button sont affichés hors du bloc « clic »,
      uniquement si le fichier correspondant est présent en mémoire.
    • Chaque download_button reçoit une key= fixe basée sur le nom de
      fichier, ce qui empêche la recréation de widgets.
    • Un bouton global « 🔄 Réinitialiser la page » est ajouté dans la sidebar ;
      il exécute `reset_page()`.
    • Même logique appliquée à toutes les pages génératrices de fichiers :
      Mise à jour M2, Classification Code, Multiconnexion, Personal Catalogue,
      Mise à jour M2 (PC) et CPN.

Nouveau 07/2025 — **Téléchargement groupé ergonomique**
    • Dès qu’une section contient au moins un fichier :
        – champ texte « 📁 Nom du dossier » (pré‑rempli) ;
        – bouton **📦 Télécharger tous les fichiers** qui livre une
          archive ZIP `<NomDuDossier>.zip` contenant chaque fichier
          dans le dossier du même nom.
"""

from __future__ import annotations
import csv, io, re, tempfile, os, sys, zipfile
from datetime import datetime
from itertools import product
from typing import Dict, Tuple, List
from streamlit_option_menu import option_menu

import pandas as pd
import streamlit as st

# ═══════════ IMPORTS OPTIONNELS ═══════════
try:                                   # libpostal
    from postal.parser import parse_address          # type: ignore
    USE_POSTAL = True
except ImportError:
    USE_POSTAL = False

try:                                   # Outlook
    import win32com.client as win32                  # type: ignore
    IS_OUTLOOK = True
except ImportError:
    IS_OUTLOOK = False

try:                                   # RAM indicator
    import psutil                                    # type: ignore
    RAM = lambda: f"{psutil.Process().memory_info().rss/1_048_576:,.0f} Mo"
except ModuleNotFoundError:
    psutil = None
    RAM = lambda: "n/a"                              # type: ignore

TODAY = datetime.today().strftime("%y%m%d")
st.set_page_config(page_title="Boîte à outils", page_icon="🛠", layout="wide")

# ════════════ PERSISTENCE : OUTILS GÉNÉRIQUES ════════════

def _save_file(section: str, label: str, data: bytes | str, filename: str, mime: str) -> None:
    if isinstance(data, str):
        data = data.encode()

    files_key = f"{section}_files"
    st.session_state.setdefault(files_key, {})

    # écrase / met à jour l'entrée (pas d'accumulation => pas de DuplicateWidgetID)
    st.session_state[files_key][filename] = {
        "label": label,
        "data": data,
        "filename": filename,
        "mime": mime,
    }


def _save_df(section: str, df: pd.DataFrame) -> None:
    st.session_state[f"{section}_df"] = df


def _build_zip(files: List[dict], folder_name: str) -> bytes:
    """Construit une archive ZIP en mémoire contenant <files> dans <folder_name>/..."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for info in files:
            zf.writestr(os.path.join(folder_name, info["filename"]), info["data"])
    buf.seek(0)
    return buf.getvalue()


def _render_downloads(section: str) -> None:
    store = st.session_state.get(f"{section}_files", {})
    # compat rétro (si liste héritée d'une ancienne session)
    if isinstance(store, list):
        # convertir à la volée
        store = {item["filename"]: item for item in store}
        st.session_state[f"{section}_files"] = store

    for info in store.values():
        st.download_button(
            info["label"],
            info["data"],
            file_name=info["filename"],
            mime=info["mime"],
            key=f"{section}_{info['filename']}",
        )

def _render_df(section: str, rows: int = 5) -> None:
    df = st.session_state.get(f"{section}_df")
    if df is not None:
        st.dataframe(df.head(rows))

# ──────────────────────────── UTILITAIRES GLOBAUX ────────────────────────────

def read_csv(buf: io.BytesIO) -> pd.DataFrame:
    """Lecture robuste CSV : détection encodage + séparateur."""
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


def read_any(upload) -> pd.DataFrame:
    name = upload.name.lower()
    if name.endswith(".csv"):
        df = read_csv(io.BytesIO(upload.getvalue()))
    elif name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(upload, engine="openpyxl", dtype=str)
    else:
        raise ValueError("Extension non gérée")

    # ⬇️  supprime les colonnes d’index ajoutées par Excel/CSV
    df = df.loc[:, ~df.columns.str.match(r'^Unnamed')]

    return df


def to_m2(s: pd.Series) -> pd.Series:
    return s.astype(str).str.zfill(6)


def sanitize_code(code: str) -> str | None:
    """Valide un code M2 de 5–6 chiffres, retourne None sinon."""
    s = str(code).strip()
    if not s.isdigit():
        return None
    return s.zfill(6) if len(s) <= 6 else None


def sanitize_numeric(series: pd.Series, width: int) -> Tuple[pd.Series, pd.Series]:
    """Pad numérique à largeur fixe + renvoie masque erreurs."""
    s = series.astype(str).str.strip()
    s_pad = s.apply(lambda x: x.zfill(width) if x.isdigit() and len(x) <= width else x)
    bad = ~s_pad.str.fullmatch(fr"\d{{{width}}}")
    return s_pad, bad

# ═══════════════════ BOUTON RÉINITIALISER (FIX) ═══════════════════

def reset_page() -> None:
    """Vide complètement st.session_state puis relance l’application."""
    st.session_state.clear()
    # Utilise désormais l’API stable – `st.rerun()`
    st.rerun()
# ═══════════════════ PAGE 1 – MISE À JOUR M2 (PC & Appairage) ═══════════════════

def _preview_file(upload) -> None:
    """Aperçu interactif : 5 lignes + liste des colonnes."""
    try:
        df = read_any(upload)
    except Exception as e:
        st.error(f"{upload.name} – lecture impossible : {e}")
        return
    with st.expander(f"📄 Aperçu – {upload.name}", expanded=False):
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
            uploads = st.file_uploader("Déposer votre fichier", type=("csv", "xlsx"),
                                       accept_multiple_files=True,
                                       key=f"{prefix}_{key}_up")
            if uploads:
                new = [u for u in uploads
                       if u.name not in st.session_state[f"{prefix}_{key}_names"]]
                st.session_state[f"{prefix}_{key}_files"] += new
                st.session_state[f"{prefix}_{key}_names"] += [u.name for u in new]
                st.success(f"{len(new)} fichier(s) ajouté(s)")
                for up in new:
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

    old_df = _add_cols(dfs["old"],
                       st.session_state[f"{prefix}_old_ref"],
                       st.session_state[f"{prefix}_old_val"],
                       "Ref", "M2_ancien")

    new_df = _add_cols(dfs["new"],
                       st.session_state[f"{prefix}_new_ref"],
                       st.session_state[f"{prefix}_new_val"],
                       "Ref", "M2_nouveau")

    merged = new_df.merge(old_df[["Ref", "M2_ancien"]], on="Ref", how="left")
    return (merged.groupby("M2_nouveau")["M2_ancien"]
                  .agg(lambda s: s.value_counts().idxmax()
                       if s.notna().any() else pd.NA)
                  ).reset_index()


def _build_appairage(prefix: str,
                     lots: dict[str, tuple[str, str, str]],
                     extra_cols: list[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    # ─── 0) Lecture + concat ────────────────────────────────────────────────
    dfs = {k: pd.concat([read_any(f) for f in st.session_state[f"{prefix}_{k}_files"]],
                        ignore_index=True).drop_duplicates()
           for k in lots}

    # ─── 1) Plan N‑1 : on garde seulement Réf + M2 ─────────────────────────
    old_df = _add_cols(dfs["old"],
                       st.session_state[f"{prefix}_old_ref"],
                       st.session_state[f"{prefix}_old_val"],
                       "Ref", "M2_ancien")

    # ─── 2) Plan N (2025) : on garde TOUTES les colonnes + normalisation M2 ─
    ref_i_new = st.session_state[f"{prefix}_new_ref"] - 1
    m2_i_new  = st.session_state[f"{prefix}_new_val"] - 1

    new_full = dfs["new"].copy()

    # On extrait les bonnes colonnes AVANT d’insérer quoi que ce soit
    ref_series = new_full.iloc[:, ref_i_new].astype(str).str.strip()
    m2_series  = new_full.iloc[:, m2_i_new]

    new_full.insert(0, "Ref", ref_series)
    new_full.insert(1, "M2_nouveau", to_m2(m2_series))

    new_df = new_full  # alias lisible

    # ─── 3) Mapping M2_ancien → Code famille client ────────────────────────
    map_df = dfs["map"].iloc[:, [st.session_state[f"{prefix}_map_ref"]-1,
                                 st.session_state[f"{prefix}_map_val"]-1]].copy()
    map_df.columns = ["M2_ancien", "Code_famille_Client"]
    map_df["M2_ancien"] = to_m2(map_df["M2_ancien"])
    old_df["M2_ancien"] = to_m2(old_df["M2_ancien"])

    merged = (new_df
              .merge(old_df[["Ref", "M2_ancien"]], on="Ref", how="left")
              .merge(map_df, on="M2_ancien", how="left"))

    # ─── 4) Table principale + table des codes sans famille ────────────────
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


def page_update_m2() -> None:
    st.header("🔄 Mise à jour des codes Mach_2")

    # Les deux onglets
    tab_pc, tab_cli = st.tabs(["📂 Personal Catalogue", "🤝 Classification Code"])

    # ------------------------------------------------------------------
    # 1) Onglet Personal Catalogue
    # ------------------------------------------------------------------
    with tab_pc:
        LOTS_PC = {
            "old": ("Ancien plan d'offre", "Référence produit", "Ancien code Mach_2"),
            "new": ("Nouveau plan d'offre", "Référence produit", "Nouveau code Mach_2"),
        }
        _uploader_state("pc", LOTS_PC)

        if st.button("🚀 Générer le fichier", key="pc_generate"):
            if not all(st.session_state[f"pc_{k}_files"] for k in LOTS_PC):
                st.warning("Chargez à la fois les fichiers N‑1 **et** N.")
                st.stop()

            maj_df = _build_m2_update("pc", LOTS_PC)
            _save_df("pc_maj", maj_df)
            _save_file("pc_maj", "⬇️ Télécharger M2_MisAJour.csv",
                       maj_df.to_csv(index=False, sep=";"),
                       f"M2_MisAJour_{TODAY}.csv", "text/csv")

        _render_downloads("pc_maj")
        _render_df("pc_maj")

    # ------------------------------------------------------------------
    # 2) Onglet Appairage client
    # ------------------------------------------------------------------
    with tab_cli:
        LOTS_CL = {
            "old": ("Ancien plan d'offre", "Référence produit", "Ancien code Mach_2"),
            "new": ("Nouveau plan d'offre", "Référence produit", "Nouveau code Mach_2"),
            "map": ("Appairage Client",
                    "Ancien code Mach2", "Code famille client"),
        }

        _uploader_state("cl", LOTS_CL)

        # Pré‑renseigner les colonnes disponibles pour le multiselect
        if (not st.session_state.get("cl_cols")
                and st.session_state.get("cl_new_files")):
            cols_new = []
            for f in st.session_state["cl_new_files"]:
                cols_new += read_any(f).columns.tolist()
            st.session_state["cl_cols"] = sorted(set(cols_new))

        extra_cols = st.multiselect(
            "Colonnes additionnelles (pour « a_remplir.csv »)",
            options=st.session_state.get("cl_cols", []),
        )

        if st.button("🚀 Générer : fichiers d’appairage", key="cl_generate"):
            if not all(st.session_state[f"cl_{k}_files"] for k in LOTS_CL):
                st.warning("Chargez les **3** jeux de données (N‑1, N, Mapping).")
                st.stop()

            if (st.session_state["cl_old_ref"] == st.session_state["cl_old_val"] or
                st.session_state["cl_new_ref"] == st.session_state["cl_new_val"]):
                st.error("« Ref produit » et « M2 » doivent être deux colonnes différentes.")
                st.stop()

            appair_df, missing_df = _build_appairage("cl", LOTS_CL, extra_cols)

            _save_df("cl", appair_df)
            _save_file("cl", "⬇️ appairage_M2_famille.csv",
                       appair_df.to_csv(index=False, sep=";"),
                       f"appairage_M2_CodeFamilleClient_{TODAY}.csv", "text/csv")
            _save_file("cl", "⬇️ a_remplir.csv",
                       missing_df.to_csv(index=False, sep=";"),
                       f"a_remplir_{TODAY}.csv", "text/csv")

        _render_downloads("cl")
        _render_df("cl")

# ═══════════════════ PAGE 2 – CLASSIFICATION CODE ═══════════════════
def page_classification():
    """Génère DFRXHYBRCMR & AFRXHYBRCMR à partir d’un appairage."""
    st.header("🧩 Classification Code ")

    # --------- 1) Appairage obligatoire ---------
    pair_file = st.file_uploader(
        "📄 Déposer le fichier d'appairage Code Mach_2/Code famille client (CSV / Excel)"
    )
    if not pair_file:
        st.info("Charger d’abord le fichier d’appairage Mach_2 → Code famille.")
        _render_downloads("cc")
        _render_df("cc")
        st.stop()

    pair_df = read_any(pair_file)
    st.dataframe(pair_df.head())

    max_cols = len(pair_df.columns)
    idx_m2  = st.number_input("🔢 Index colonne Code Mach_2", 1, max_cols, 1)
    idx_fam = st.number_input("🔢 Index colonne Code famille client", 1, max_cols, 2)

    entreprise = st.text_input("🏢 Entreprise")

    if st.button("🚀 Générer les fichiers", key="class_generate"):
        if not entreprise:
            st.warning("Renseigne le champ Entreprise.")
            st.stop()

        try:
            col_m2  = pair_df.columns[int(idx_m2)  - 1]
            col_fam = pair_df.columns[int(idx_fam) - 1]
        except IndexError:
            st.error("Indice de colonne hors plage.")
            st.stop()

        # ---- Construction du dataframe ----
        df_out = pair_df[[col_fam, col_m2]].rename(columns={
            col_fam: "Code famille Client",
            col_m2:  "M2",
        }).copy()

        raw_m2 = df_out["M2"].astype(str).str.strip()
        sanitized = raw_m2.apply(sanitize_code)
        invalid_mask = sanitized.isna()
        if invalid_mask.any():
            st.error(f"{invalid_mask.sum()} code(s) M2 invalides – uniquement 5 ou 6 chiffres.")
            st.dataframe(raw_m2[invalid_mask].to_frame("Code fourni"))
            st.stop()

        df_out["M2"] = sanitized.map(lambda x: f"M2_{x}")
        df_out["onsenfou"]   = None
        df_out["Entreprises"] = entreprise

        # ---- Ajout colonne vide en première position ----
        df_out.insert(0, "Empty", "")

        # ---- Réordonnage final ----
        df_out = df_out[["Empty", "Code famille Client", "onsenfou", "Entreprises", "M2"]]

        dstr       = TODAY
        dfrx_name  = f"DFRXHYBRCMR{dstr}0000"
        afrx_name  = f"AFRXHYBRCMR{dstr}0000"   # <-- plus d’extension .txt

        _save_df("cc", df_out)
        _save_file(
            "cc",
            "📥 DFRXHYBRCMR",
            df_out.to_csv(sep="\t", index=False, header=False).encode(),
            dfrx_name,
            "text/tab-separated-values"
        )

        ack_txt = (
            f"DFRXHYBRCMR{dstr}000068230116IT"
            f"DFRXHYBRCMR{dstr}RCMRHYBFRX                    OK000000"
        )
        _save_file(
            "cc",
            "📥 AFRXHYBRCMR",
            ack_txt,
            afrx_name,
            "text/plain"
        )

    _render_downloads("cc")
    _render_df("cc")

# ═══════════════════ PAGE 3 – Multiconnexion ═══════════════════

def to_xlsx(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    buf.seek(0)
    return buf.getvalue()


# ───────────────────────── Helper adresse → dict ────────────────────
# ───────────────────────── Helper adresse → dict ────────────────────
def split_address(addr: str) -> dict:
    """
    Tente d’analyser l’adresse.
    1) Utilise libpostal si disponible.
    2) Sinon, regex tolérante : prend en charge « 10 bis rue X, 75000 Paris »,
       « 10bis rue X 75000 PARIS », etc.
    Renvoie toujours un dict : num, voie, cp, ville, pays.
    """
    parts = {"num": "", "voie": "", "cp": "", "ville": "", "pays": "FR"}

    if USE_POSTAL:                                  # libpostal disponible
        for val, label in parse_address(addr or ""):
            if label == "house_number":
                parts["num"] = val
            elif label in {"road", "pedestrian", "path", "footway"}:
                parts["voie"] = val
            elif label == "postcode":
                parts["cp"] = val
            elif label in {"city", "town", "village", "suburb"}:
                parts["ville"] = val
            elif label == "country":
                parts["pays"] = val
        return parts

    # -------- Fallback regex amélioré --------
    import re
    pattern = re.compile(
        r"""
        ^\s*
        (?P<num>[\d\w\-]*)          # numéro : 10, 10B, 5-7, etc. (facultatif)
        \s*
        (?P<voie>[^,]+?)            # libellé de voie jusqu'à virgule ou CP
        [,\s]+
        (?P<cp>\d{2}\s?\d{3})       # code postal (75008 ou 75 008)
        \s+
        (?P<ville>.+?)              # ville
        \s*$
        """,
        re.VERBOSE | re.IGNORECASE,
    )
    m = pattern.match(addr or "")
    if m:
        parts.update(m.groupdict())
    return parts


def build_tables(
    df_src: pd.DataFrame,
    *,
    entreprise: str,
    view_master_catalog: str,
    punchout_user_id: str,
    domain: str,
    identity: str,
    integration_type: str = "OCI",   # ou "cXML"
) -> list[pd.DataFrame]:
    """
    Construit PF1 → PF5 (+ PF6 si cXML) au format attendu.
    Le DataFrame d’entrée doit contenir :
        « Numéro de compte », « Raison sociale », « Adresse », « Code agence ».
    """
    pf1_cols = [
        "uid", "name", "locName",
        "CXmIAssignedConfiguration",
        "pcCompoundProfile",
        "ViewMasterCatalog",
    ]
    pf2_cols = [
        "B2B Unit",
        "ADRESSE / Numéro de rue", "ADRESSE / rue",
        "ADRESSE / Code postal",   "ADRESSE / Ville",
        "ADRESSE / Pays/Région",
        "INFORMATIONS D'ADRESSE SUPPLÉMENTAIRES / Téléphone 1",
    ]
    pf3_cols = ["B2BUnitID", "itemtype", "managingBranches", "punchoutUserID", "sealed"]
    pf4_cols = ["aliasName", "branch", "punchoutUserID", "sealed"]
    pf5_cols = ["B2BUnitID", "punchoutUserID"]
    pf6_cols = ["number", "domain", "identity"]

    pf1 = pd.DataFrame(columns=pf1_cols)
    pf2 = pd.DataFrame(columns=pf2_cols)
    pf3 = pd.DataFrame(columns=pf3_cols)
    pf4 = pd.DataFrame(columns=pf4_cols)
    pf5 = pd.DataFrame(columns=pf5_cols)
    pf6 = pd.DataFrame(columns=pf6_cols)  # utilisé seulement en cXML

    sealed_val = "false"

    for _, row in df_src.iterrows():
        account   = row["Numéro de compte"]
        company   = row["Raison sociale"]
        full_addr = (row["Adresse"] or "").strip()  # <-- trim pour éviter NaN/espaces

        branch    = row["Code agence"]      # ← remplace ManagingBranch

        # PF1 : locName == name (plus d'adresse)
        pf1.loc[len(pf1)] = [
            account,
            company,
            company,                        # locName = company
            f"frx-variant-{entreprise}-configuration-set",
            f"PC_{entreprise}",
            view_master_catalog,
        ]

        # PF2 : adresse détaillée
        addr = split_address(full_addr)
        pf2.loc[len(pf2)] = [
            account, addr["num"], addr["voie"],
            addr["cp"], addr["ville"], addr["pays"], ""
        ]

        # PF3
        pf3.loc[len(pf3)] = [
            account,
            "PunchoutAccountAndBranchAssociation",
            branch,
            punchout_user_id,
            sealed_val,
        ]

        # PF4
        pf4.loc[len(pf4)] = [
            branch,
            branch,
            punchout_user_id,
            sealed_val,
        ]

        # PF5
        pf5.loc[len(pf5)] = [account, punchout_user_id]

        # PF6 (cXML uniquement)
        if integration_type == "cXML":
            pf6.loc[len(pf6)] = [account, domain, identity]

    tables = [pf1, pf2, pf3, pf4, pf5]
    if integration_type == "cXML":
        tables.append(pf6)
    return tables

def create_outlook_draft(att: List[Tuple[str, bytes]],
                         to_: str, subject: str) -> None:
    if not IS_OUTLOOK:
        return
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # olMailItem
    mail.To = to_
    mail.Subject = subject
    mail.Body = "Bonjour,\n\nVeuillez trouver les fichiers PF en pièce jointe.\n"
    for name, data in att:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=name)
        tmp.write(data)
        tmp.close()
        mail.Attachments.Add(tmp.name)
    mail.Display()


def page_multiconnexion():
    st.header("📦 Multiconnexion")

    # --- 1) Choix du type d’intégration ---
    integration_type = st.radio(
        "Type d’intégration",
        ["cXML", "OCI"],
        horizontal=True
    )

    st.markdown(
        "Télécharger le modèle, le compléter, puis uploader le fichier.  \n"
        "Colonnes requises : **Numéro de compte** (7 chiffres), **Raison sociale**, "
        "**Adresse**, **Code d'agence** (4 chiffres)."
    )

    # 2) Template vierge ----------------------------------------------------------------
    with st.expander("📑 Template Multiconnexion.xlsx"):
        cols_tpl = ["Numéro de compte", "Raison sociale", "Adresse", "Code agence"]
        buf_tpl = io.BytesIO()
        pd.DataFrame([{c: "" for c in cols_tpl}]).to_excel(buf_tpl, index=False)
        buf_tpl.seek(0)
        st.download_button(
            "📥 Télécharger le template",
            buf_tpl.getvalue(),
            file_name="dfrecu_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="multi_template"
        )

    # 3) Fichier source -----------------------------------------------------------------
    up_file = st.file_uploader(
        "📄 Déposer Fichier Multiconnexion",
        type=("csv", "xlsx", "xls")
    )
    if not up_file:
        _render_downloads("multi")
        _render_df("multi")
        st.stop()

    # 4) Paramètres d’en‑tête ------------------------------------------------------------
    col1, col2 = st.columns(2)
    with col1:
        entreprise   = st.text_input("🏢 Entreprise").strip()
    with col2:
        punchout_user = st.text_input("👤 punchoutUserID")

    # Domain & Identity uniquement en mode cXML
    if integration_type == "cXML":
        col3, col4 = st.columns(2)
        with col3:
            domain = st.selectbox("🌐 Domain", ["NetworkID", "DUNS"])
        with col4:
            identity = st.text_input("🆔 Identity")
    else:
        domain = ""
        identity = ""

    # Options catalogue perso
    vm_choice   = st.radio("ViewMasterCatalog", ["True", "False"], horizontal=True)
    pc_enabled  = st.radio("Personal Catalogue ?", ["True", "False"], horizontal=True)
    pc_name     = st.text_input(
        "Nom du catalogue (sans PC_)",
        placeholder="CATALOGUE"
    ).strip() if pc_enabled == "True" else ""

    # 5) Bouton génération ---------------------------------------------------------------
    if st.button("🚀 Générer les fichiers", key="multi_generate"):

        # Champs requis selon le mode
        base_required = [entreprise, punchout_user, (pc_enabled == "False" or pc_name)]
        cx_required   = [domain, identity] if integration_type == "cXML" else []
        if not all(base_required + cx_required):
            st.warning("Remplir tous les champs requis.")
            st.stop()

        # --- Lecture + normalisation colonnes ---
        df_src = read_any(up_file)

        # tolérance casse / espaces
               # --- tolérance casse / espaces ---
        norm_map = {c: c.strip().lower() for c in df_src.columns}
        expected = {
            "Numéro de compte": "numéro de compte",
            "Code agence":      "code agence",
        }
        missing = [orig for orig, norm in expected.items() if norm not in norm_map.values()]
        if missing:
            st.error(f"Colonnes manquantes ou mal orthographiées : {', '.join(missing)}")
            st.stop()

        # Renomme pour avoir exactement les libellés attendus ensuite
        for orig, norm in norm_map.items():
            if norm == "numéro de compte":
                df_src.rename(columns={orig: "Numéro de compte"}, inplace=True)
            elif norm == "code agence":
                df_src.rename(columns={orig: "Code agence"}, inplace=True)


        # --- Contrôles format compte / branche ---
                df_src["Numéro de compte"], bad_acc = sanitize_numeric(df_src["Numéro de compte"], 7)
        df_src["Code agence"],      bad_ag  = sanitize_numeric(df_src["Code agence"], 4)

        if bad_acc.any() or bad_ag.any():
            st.error("Numéro de compte ou Code agence invalide(s).")
            st.stop()

        # --- Construction des tables PF ---
        tables = build_tables(
            df_src,
            entreprise=entreprise,
            view_master_catalog=vm_choice,
            punchout_user_id=punchout_user,
            domain=domain,
            identity=identity,
            integration_type=integration_type,
        )

        # --- Export XLSX + boutons ---
        file_map = {
            "PF1": f"B2B Units creation_{entreprise}.xlsx",
            "PF2": f"Table_chargement_adresse_{entreprise}.xlsx",
            "PF3": f"Table_PunchoutAccountAndBranchAssociation_{entreprise}.xlsx",
            "PF4": f"PunchoutBranchAliasAssociation_{entreprise}.xlsx",
            "PF5": f"Table_Attach_B2BUnitstoUsers_{entreprise}.xlsx",
            "PF6": f"PunchoutAccountSetup_{entreprise}.xlsx",
        }
        labels = ["PF1", "PF2", "PF3", "PF4", "PF5"] + (
            ["PF6"] if integration_type == "cXML" else []
        )

        for lbl, df in zip(labels, tables):
            data = to_xlsx(df)
            _save_file(
                "multi",
                f"⬇️ {lbl}",
                data,
                file_map[lbl],
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        _save_df("multi", tables[0])

    _render_downloads("multi")
    _render_df("multi")

    # 6) Option Outlook ------------------------------------------------------------------
    if IS_OUTLOOK and st.session_state.get("multi_files"):
        st.markdown("---")
        dest = st.text_input("Destinataire (Outlook)")
        if st.button("Ouvrir un brouillon Outlook", key="multi_outlook"):
            subj = f"Fichiers PF – {entreprise} ({datetime.now():%Y-%m-%d %H:%M})"
            files_att = [
                (info["filename"], info["data"]) for info in st.session_state["multi_files"]
            ]
            create_outlook_draft(files_att, to_=dest, subject=subj)
            st.success("Brouillon Outlook ouvert.")
    elif not IS_OUTLOOK:
        st.info("Automatisation Outlook indisponible sur cet environnement.")

# ═══════════════════ PAGE 4 – GÉNÉRATEUR PC / MàJ M2 (AFRX ajouté) ═══════════════════

def generator_pc_common(codes: pd.Series, entreprise: str, statut: str) -> pd.DataFrame:
    return pd.DataFrame({
        0: [f"PC_PROFILE_{entreprise}"] * len(codes),
        1: [statut] * len(codes),
        2: [None] * len(codes),
        3: [f"M2_{c}" for c in codes],
        4: ["frxProductCatalog:Online"] * len(codes),
    }).drop_duplicates()


def export_pc_files(section: str, df1: pd.DataFrame,
                    comptes: pd.Series,
                    entreprise: str,
                    dstr: str = TODAY) -> None:
    """Crée les 4 entrepôts dans le state et donc les boutons persistants."""

    # 1‑ profil PC
    _save_file(section, "⬇️ DFRXHYBRPCP",
               df1.to_csv(sep=";", index=False, header=False),
               f"DFRXHYBRPCP{dstr}0000", "text/plain")

    # 2‑ ACK CMP
    ack_cmp = (
        f"DFRXHYBRCMP{dstr}000068240530IT"
        f"DFRXHYBRCMP{dstr}CCMGHYBFRX                    OK000000"
    )
    _save_file(section, "⬇️ AFRXHYBRCMP",
               ack_cmp,
               f"AFRXHYBRCMP{dstr}0000.txt", "text/plain")

    # 3‑ rattachement comptes → profil
    cmp_content = (
        f"PC_{entreprise};PC_{entreprise};PC_PROFILE_{entreprise};"
        f"{','.join(comptes)};frxProductCatalog:Online"
    )
    _save_file(section, "⬇️ DFRXHYBRCMP",
               cmp_content,
               f"DFRXHYBRCMP{dstr}0000", "text/plain")

    # 4‑ ACK PCP
    ack_pcp = (
        f"DFRXHYBRPCP{dstr}000068200117IT"
        f"DFRXHYBRPCP{dstr}RCMRHYBFRX                    OK000000"
    )
    _save_file(section, "⬇️ AFRXHYBRPCP",
               ack_pcp,
               f"AFRXHYBRPCP{dstr}0000.txt", "text/plain")

# ─────────────────────────  GÉNÉRATEUR PC  ─────────────────────────

def generator_pc():
    st.subheader("Personal Catalogue")

    # ⬇️ ancien : "📄 Codes produit"
    codes_file = st.file_uploader(
        "📄 Fichier contenant la colonne Mach_2 (CSV / Excel)",
        type=("csv", "xlsx", "xls"),
        key="pc_codes",
    )
    if codes_file:
        df_codes_tmp = read_any(codes_file)
        with st.expander("Aperçu – codes produit"):
            st.dataframe(df_codes_tmp.head())
            st.write("Index colonnes :", [f"{i+1} – {c}" for i, c in enumerate(df_codes_tmp.columns)])

    compte_file = st.file_uploader(
        "📄 Déposer le Fichier des numéros de compte (CSV / Excel)",
        type=("csv", "xlsx", "xls"),
        key="pc_comptes",
    )
    if compte_file:
        df_comptes_tmp = read_any(compte_file)
        with st.expander("Aperçu – numéros de compte"):
            st.dataframe(df_comptes_tmp.head())
            st.write("Index colonnes :", [f"{i+1} – {c}" for i, c in enumerate(df_comptes_tmp.columns)])

    if codes_file and compte_file:
        nb_cols_codes = len(read_any(codes_file).columns)
        nb_cols_comp  = len(read_any(compte_file).columns)

        col_idx_codes = st.number_input("🔢 Colonne contenant les codes M2", 1, nb_cols_codes, 1)
        col_idx_comptes = st.number_input("🔢 Colonne contenant les numéros de compte", 1, nb_cols_comp, 1)

        entreprise = st.text_input("🏢 Entreprise")
        statut     = st.selectbox("📌 Statut", ["", "INCLUDE", "EXCLUDE"])

        if st.button("🚀 Générer PC", key="gen_pc_generate"):
            if not all([entreprise, statut]):
                st.warning("Renseigner l’entreprise et le statut.")
                st.stop()

            df_codes   = read_any(codes_file)
            df_comptes = read_any(compte_file)

            try:
                raw_codes = df_codes.iloc[:, col_idx_codes - 1].dropna().astype(str).str.strip()
                comptes   = df_comptes.iloc[:, col_idx_comptes - 1].dropna().astype(str).str.strip()
            except IndexError:
                st.error("Indice de colonne hors plage.")
                st.stop()

            sanitized = raw_codes.apply(sanitize_code)
            invalid_mask = sanitized.isna()

            if invalid_mask.any():
                st.error(f"{invalid_mask.sum()} code(s) M2 invalides – uniquement 5 ou 6 chiffres.")
                st.dataframe(raw_codes[invalid_mask].to_frame("Code fourni"))
                st.stop()

            codes = sanitized
            dstr  = TODAY

            df1 = generator_pc_common(codes, entreprise, statut)
            _save_df("gen_pc", df1)
            export_pc_files("gen_pc", df1, comptes, entreprise, dstr)

    _render_downloads("gen_pc")
    _render_df("gen_pc")

# ─────────────────────────  MISE À JOUR M2 (avant génération PC) ─────────────────────────

def generator_maj_m2():
    st.subheader("Mise à jour Mach_2 avant génération")

    # ⬇️ ancien : "📄 Codes produit"
    codes_file = st.file_uploader(
        "📄 Fichier contenant la colonne Mach_2 (CSV / Excel)",
        type=("csv", "xlsx", "xls"))
    col_idx_codes = st.number_input("🔢 Colonne Codes Mach_2", 1, 50, 1) if codes_file else None

    compte_file = st.file_uploader("📄 Numéros de compte", type=("csv", "xlsx", "xls"))
    col_idx_comptes = st.number_input("🔢 Colonne comptes (1=A)", 1, 50, 1) if compte_file else None

    map_file = st.file_uploader("📄 Fichier Mach_2_MisAJour", type=("csv", "xlsx", "xls"))
    if map_file:
        col_idx_old = st.number_input("🔢 Colonne M2 ancien", 1, 50, 1)
        col_idx_new = st.number_input("🔢 Colonne M2 nouveau", 1, 50, 2)
    else:
        col_idx_old = col_idx_new = None

    entreprise = st.text_input("🏢 Entreprise")
    statut     = st.selectbox("📌 Statut", ["", "INCLUDE", "EXCLUDE"])

    if st.button("🚀 Générer MàJ", key="majm2_generate"):
        required = [codes_file, compte_file, map_file, entreprise, statut,
                    col_idx_codes, col_idx_comptes, col_idx_old, col_idx_new]
        if not all(required):
            st.warning("Remplir tous les champs et joins les 3 fichiers.")
            st.stop()

        df_codes   = read_any(codes_file)
        df_comptes = read_any(compte_file)
        df_map     = read_any(map_file)

        try:
            raw_codes = df_codes.iloc[:, col_idx_codes-1].dropna().astype(str).str.strip()
            comptes   = df_comptes.iloc[:, col_idx_comptes-1].dropna().astype(str).str.strip()
        except IndexError:
            st.error("Indice colonne hors plage.")
            st.stop()

        sanitized = raw_codes.apply(sanitize_code)
        if sanitized.isna().any():
            st.error("Codes M2 invalides détectés.")
            st.dataframe(raw_codes[sanitized.isna()].to_frame("Code fourni"))
            st.stop()

        try:
            old_codes = df_map.iloc[:, col_idx_old-1].astype(str).apply(sanitize_code)
            new_codes = df_map.iloc[:, col_idx_new-1].astype(str).apply(sanitize_code)
        except IndexError:
            st.error("Indice colonne mapping hors plage.")
            st.stop()

        mapping = (pd.DataFrame({"old": old_codes, "new": new_codes})
                   .dropna()
                   .drop_duplicates("old")
                   .set_index("old")["new"].to_dict())

        updated_codes = sanitized.map(lambda c: mapping.get(c, c))

        dstr = TODAY
        df1 = generator_pc_common(updated_codes, entreprise, statut)
        _save_df("majm2", df1)
        export_pc_files("majm2", df1, comptes, entreprise, dstr)

    _render_downloads("majm2")
    _render_df("majm2")


# ───────── page_dfrx_pc : navigation corrigée ─────────
def page_dfrx_pc():
    st.header("🛠️ Personal Catalogue")
    nav = st.radio(
        "Choisir l’outil",
        ["Sans mise à jour Mach_2", "Avec mise à jour Mach_2"],
        horizontal=True
    )

    if nav == "Sans mise à jour Mach_2":
        generator_pc()
    else:
        generator_maj_m2()

# ═══════════════════ PAGE 5 – CPN GENERATOR ═══════════════════
def page_cpn():
    st.header("📑 CPN")

    colA, colB = st.columns(2)

    # ─── Upload fichiers ───
    with colA:
        main_file = st.file_uploader(
            "📄 Appairage Code produit client / Référence interne",
            type=("csv", "xlsx", "xls")
        )
        if main_file:
            _preview_file(main_file)

    with colB:
        cli_file = st.file_uploader(
            "📄 Périmètre (comptes client)",
            type=("csv", "xlsx", "xls")
        )
        if cli_file:
            _preview_file(cli_file)

    if not (main_file and cli_file):
        _render_downloads("cpn")
        _render_df("cpn")
        st.stop()

    df_main = read_any(main_file)
    df_cli  = read_any(cli_file)

    # ─── Choix des colonnes ───
    col_int = st.selectbox(
        "Colonne Référence produit interne",
        range(1, len(df_main.columns) + 1),
        index=0
    )
    col_cli_prod = st.selectbox(
        "Colonne Code produit client",
        range(1, len(df_main.columns) + 1),
        index=1 if len(df_main.columns) > 1 else 0
    )
    col_cli_acc = st.selectbox(
        "Colonne Numéro de compte client (périmètre)",
        range(1, len(df_cli.columns) + 1),
        index=0
    )

    # ─── Génération ───
        # ─── Génération ───
    if st.button("🚀 Générer CPN", key="cpn_generate"):
        series_int = df_main.iloc[:, col_int - 1].astype(str).str.strip()
        if (~series_int.str.fullmatch(r"\d{8}")).any():
            st.error("Réf. interne invalide : doit contenir exactement 8 chiffres.")
            st.dataframe(series_int[~series_int.str.fullmatch(r'\d{8}')])
            st.stop()

        series_cli_prod = df_main.iloc[:, col_cli_prod - 1].astype(str).str.strip()
        series_cli_acc  = df_cli.iloc[:,  col_cli_acc  - 1].astype(str).str.strip()

        # Produit cartésien InternalItemID × AccountNumber
        pf = pd.DataFrame(
            product(series_int, series_cli_acc),
            columns=["InternalItemID", "AccountNumber"]
        )
        pf["CustomerItemId"] = series_cli_prod.repeat(len(series_cli_acc)).values

        # Ré‑ordonnage final (sans colonne vide)
        df_out = pf[["CustomerItemId", "AccountNumber", "InternalItemID"]]

        # Nomenclature fichiers
        today     = TODAY
        dfrx_name = f"DFRXHYBCPNA{today}0000"
        afrx_name = f"AFRXHYBCPNA{today}0000"

        ack_txt = (
            f"DFRXHYBCPNA{today}000148250201IT"
            f"DFRXHYBCPNA{today}CPNAHYBFRX                    OK000000"
        )

        # Sauvegarde & téléchargements
        _save_df("cpn", df_out)
        _save_file(
            "cpn",
            "⬇️ DFRX (TSV)",
            df_out.to_csv(sep="\t", index=False, header=False).encode(),
            dfrx_name,
            "text/tab-separated-values"
        )
        _save_file("cpn", "⬇️ AFRX (TXT)", ack_txt, afrx_name, "text/plain")



# ═══════════════════════════  MENU PRINCIPAL ═══════════════════════════
# ═══════════════════════════  MENU PRINCIPAL ═══════════════════════════
PAGES = {
    "Mise à jour Mach_2": page_update_m2,
    "Classification Code": page_classification,
    "Multiconnexion": page_multiconnexion,
    "Personal Catalogue": page_dfrx_pc,
    "CPN": page_cpn,
}

with st.sidebar:
    # ▸ Bouton global de réinitialisation
    if st.button("🔄 Réinitialiser la page", key="reset_page_button"):
        reset_page()
        st.rerun()

    # ▸ Menu de navigation principal
    choice = st.radio(
        "Navigation", list(PAGES), index=0, key="nav_main",
        label_visibility="collapsed",
    )

# — exécution de la page choisie —
PAGES[choice]()
