# -*- coding: utf-8 -*-
"""
multi_tool_app.py – Application « Boîte à outils » (Streamlit)

Correctifs 07/2025
    • PF1→PF6 : noms explicites (plus de timestamp)
    • Générateur PC & MàJ M2 : ajout du fichier AFRXHYBRPCP<date>0000.txt
    • Correction d’une parenthèse non fermée (section Outlook)
"""

from __future__ import annotations
import csv, io, re, tempfile, os, sys
from datetime import datetime
from itertools import product
from typing import Dict, Tuple, List

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

# ──────────────────────────── UTILITAIRES GLOBAUX ────────────────────────────
@st.cache_data(show_spinner=False, hash_funcs={io.BytesIO: lambda _: None})
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

@st.cache_data(show_spinner=False, hash_funcs={io.BytesIO: lambda _: None})
def read_any(upload) -> pd.DataFrame:
    """Lecture CSV ou Excel (XLSX/XLS) avec cache."""
    name = upload.name.lower()
    if name.endswith(".csv"):
        return read_csv(io.BytesIO(upload.getvalue()))
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(upload, engine="openpyxl", dtype=str)
    raise ValueError("Extension non gérée")

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

# ═══════════════════ PAGE 1 – MISE À JOUR M2 (PC & Appairage) ═══════════════════
# (code d’origine inchangé : seul l’onglet PF était concerné par la correction)

def _preview_file(upload):  # .. helper (inchangé)
    ...

def _uploader_state(prefix: str, lots: dict[str, tuple[str, str, str]]):
    ...

def _add_cols(df: pd.DataFrame, ref_i: int, m2_i: int,
              ref_label: str, m2_label: str) -> pd.DataFrame:
    ...

def _build_m2_update(prefix: str, lots):    ...
def _build_appairage(prefix: str, lots, extra_cols): ...

def page_update_m2():  # (code d’origine)
    ...

# ═══════════════════ PAGE 2 – CLASSIFICATION CODE ═══════════════════
def page_classification():  # (inchangé)
    ...

# ═══════════════════ PAGE 3 – PF1 → PF6 GENERATOR (corrigé) ═══════════════════
def to_xlsx(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    buf.seek(0)
    return buf.getvalue()

def build_tables(df_src: pd.DataFrame) -> List[pd.DataFrame]:
    """Placeholder : implémentez votre logique métier ici."""
    raise NotImplementedError

def create_outlook_draft(attachments: List[Tuple[str, bytes]],
                         to_: str = "", subject: str = "", body: str = ""):
    if not IS_OUTLOOK:
        raise RuntimeError("Outlook COM indisponible.")
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to_
    mail.Subject = subject
    mail.Body = body or "Bonjour,\n\nVeuillez trouver les fichiers PF en pièce jointe.\n"
    for name, data in attachments:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=name)
        tmp.write(data)
        tmp.close()
        mail.Attachments.Add(tmp.name)
    mail.Display()

def page_multiconnexion():
    st.header("📦 Générateur PF1 → PF6 (Multiconnexion)")
    integration_type = st.radio("Type d’intégration", ["cXML", "OCI"], horizontal=True)

    st.markdown(
        "Téléchargez le template, remplissez‑le puis uploadez votre fichier.  \n"
        "Colonnes requises : **Numéro de compte** (7 chiffres), **Raison sociale**, "
        "**Adresse**, **ManagingBranch** (4 chiffres)."
    )

    # Template vierge
    with st.expander("📑 Template dfrecu.xlsx"):
        cols_tpl = ["Numéro de compte", "Raison sociale", "Adresse", "ManagingBranch"]
        buf_tpl = io.BytesIO()
        pd.DataFrame([{c: "" for c in cols_tpl}]).to_excel(buf_tpl, index=False)
        buf_tpl.seek(0)
        st.download_button("📥 Télécharger le template", buf_tpl.getvalue(),
                           file_name="dfrecu_template.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    up_file = st.file_uploader("📄 Fichier dfrecu", type=("csv", "xlsx", "xls"))
    if not up_file:
        st.stop()

    col1, col2, col3 = st.columns(3)
    with col1:
        entreprise = st.text_input("🏢 Entreprise").strip()
    with col2:
        punchout_user = st.text_input("👤 punchoutUserID")
    with col3:
        domain = st.selectbox("🌐 Domain", ["NetworkID", "DUNS"])

    identity = st.text_input("🆔 Identity")
    vm_choice = st.radio("ViewMasterCatalog", ["True", "False"], horizontal=True)
    pc_enabled = st.radio("Personal Catalogue ?", ["True", "False"], horizontal=True)
    pc_name = st.text_input("Nom du catalogue (sans PC_)", placeholder="CATALOGUE").strip() \
              if pc_enabled == "True" else ""

    if st.button("🚀 Générer PF"):
        if not all([entreprise, punchout_user, identity, (pc_enabled == "False" or pc_name)]):
            st.warning("Remplis tous les champs requis.")
            st.stop()

        df_src = read_any(up_file)
        if {"Numéro de compte", "ManagingBranch"} - set(df_src.columns):
            st.error("Colonnes manquantes dans le fichier.")
            st.stop()

        df_src["Numéro de compte"], bad_acc = sanitize_numeric(df_src["Numéro de compte"], 7)
        df_src["ManagingBranch"], bad_man = sanitize_numeric(df_src["ManagingBranch"], 4)
        if bad_acc.any() or bad_man.any():
            st.error("Numéro de compte ou ManagingBranch invalide(s).")
            st.stop()

        tables: List[pd.DataFrame] = build_tables(df_src)  # votre logique métier

        file_map = {
            "PF1": f"B2B Units creation_{entreprise}.xlsx",
            "PF2": f"Table_chargement_adresse_{entreprise}.xlsx",
            "PF3": f"Table_PunchoutAccountAndBranchAssociation_{entreprise}.xlsx",
            "PF4": f"PunchoutBranchAliasAssociation_{entreprise}.xlsx",
            "PF5": f"Table_Attach_B2BUnitstoUsers_{entreprise}.xlsx",
            "PF6": f"PunchoutAccountSetup_{entreprise}.xlsx",
        }

        labels = ["PF1", "PF2", "PF3", "PF4", "PF5"] + (["PF6"] if integration_type == "cXML" else [])
        files_bytes: Dict[str, bytes] = {}
        for lbl, df in zip(labels, tables):
            fname = file_map[lbl]
            data = to_xlsx(df)
            files_bytes[fname] = data
            st.download_button(f"⬇️ {lbl}", data, file_name=fname,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.dataframe(tables[0].head())
        st.success("✅ Fichiers prêts !")

        # Option Outlook
        if IS_OUTLOOK:
            st.markdown("---")
            dest = st.text_input("Destinataire (Outlook)")
            if st.button("Ouvrir un brouillon Outlook"):
                subj = f"Fichiers PF – {entreprise} ({datetime.now():%Y-%m-%d %H:%M})"
                create_outlook_draft(list(files_bytes.items()), to_=dest, subject=subj)
                st.success("Brouillon Outlook ouvert.")
        else:
            st.info("Automatisation Outlook indisponible sur cet environnement.")

# ═══════════════════ PAGE 4 – GÉNÉRATEUR PC / MàJ M2 (AFRX ajouté) ═══════════════════
def generator_pc_common(codes: pd.Series, entreprise: str, statut: str) -> pd.DataFrame:
    return pd.DataFrame({
        0: [f"PC_PROFILE_{entreprise}"] * len(codes),
        1: [statut] * len(codes),
        2: [None] * len(codes),
        3: [f"M2_{c}" for c in codes],
        4: ["frxProductCatallog:Online"] * len(codes),
    }).drop_duplicates()

def export_pc_files(df1: pd.DataFrame, dstr: str):
    dfrx_name = f"DFRXHYBRPCP{dstr}0000"
    afrx_name = f"AFRXHYBRPCP{dstr}0000.txt"
    afrx_txt = (
        f"DFRXHYBRPCP{dstr}000068200117IT"
        f"DFRXHYBRPCP{dstr}RCMRHYBFRX                    OK000000"
    )
    st.download_button("⬇️ DFRXHYBRPCP", df1.to_csv(index=False, header=False, sep=";"),
                       file_name=dfrx_name, mime="text/plain")
    st.download_button("⬇️ AFRXHYBRPCP", afrx_txt,
                       file_name=afrx_name, mime="text/plain")

def generator_pc():
    st.subheader("Générateur PC")
    codes_file = st.file_uploader("Codes produit", type=("csv", "xlsx", "xls"))
    if not codes_file: st.stop()
    col_idx_codes = st.number_input("Colonne Codes M2 (1=A)", 1, 50, 1)

    compte_file = st.file_uploader("Numéros de compte", type=("csv", "xlsx", "xls"))
    if not compte_file: st.stop()
    col_idx_comptes = st.number_input("Colonne comptes (1=A)", 1, 50, 1)

    entreprise = st.text_input("Entreprise")
    statut = st.selectbox("Statut", ["", "INCLUDE", "EXCLUDE"])

    if st.button("🚀 Générer PC"):
        df_codes = read_any(codes_file)
        codes_raw = df_codes.iloc[:, col_idx_codes-1].dropna()
        codes = codes_raw.astype(str).apply(sanitize_code)
        if codes.isna().any():
            st.error("Codes M2 invalides."); st.stop()

        df1 = generator_pc_common(codes, entreprise, statut)
        export_pc_files(df1, TODAY)
        st.success("Fichiers générés.")

def generator_maj_m2():
    st.subheader("Mise à jour M2 avant génération")
    codes_file = st.file_uploader("Codes produit", type=("csv", "xlsx", "xls"))
    if not codes_file: st.stop()
    col_idx_codes = st.number_input("Colonne Codes M2 (1=A)", 1, 50, 1)

    map_file = st.file_uploader("Fichier M2_MisAJour", type=("csv", "xlsx", "xls"))
    if not map_file: st.stop()
    col_idx_old = st.number_input("Colonne M2 ancien", 1, 50, 1)
    col_idx_new = st.number_input("Colonne M2 nouveau", 1, 50, 2)

    entreprise = st.text_input("Entreprise")
    statut = st.selectbox("Statut", ["", "INCLUDE", "EXCLUDE"])

    if st.button("🚀 Générer MàJ"):
        df_codes = read_any(codes_file)
        df_map = read_any(map_file)

        raw_codes = df_codes.iloc[:, col_idx_codes-1]
        codes = raw_codes.astype(str).apply(sanitize_code)

        old_codes = df_map.iloc[:, col_idx_old-1].astype(str).apply(sanitize_code)
        new_codes = df_map.iloc[:, col_idx_new-1].astype(str).apply(sanitize_code)
        mapping = pd.Series(new_codes.values, index=old_codes).dropna().to_dict()
        updated = codes.map(lambda c: mapping.get(c, c))

        df1 = generator_pc_common(updated, entreprise, statut)
        export_pc_files(df1, TODAY)
        st.success("Fichiers générés.")

def page_dfrx_pc():
    st.header("🛠️ Générateur PC + Mise à jour M2")
    nav = st.radio("Choisir l’outil", ["Générateur PC", "Mise à jour M2"], horizontal=True)
    (generator_pc if nav == "Générateur PC" else generator_maj_m2)()

# ═══════════════════ PAGE 5 – CPN GENERATOR (inchangé) ═══════════════════
def page_cpn(): ...
# (code d’origine)

# ═══════════════════ MENU PRINCIPAL ═══════════════════
PAGES = {
    "Mise à jour M2": page_update_m2,
    "Classification Code": page_classification,
    "PF1 → PF6 Generator": page_multiconnexion,
    "Générateur PC / MàJ M2": page_dfrx_pc,
    "CPN Generator": page_cpn,
}
choice = st.sidebar.radio("Navigation", list(PAGES.keys()))
PAGES[choice]()
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
