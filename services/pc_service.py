# -*- coding: utf-8 -*-
"""
services/pc_service.py
Logique métier : Personal Catalogue (DFRXHYBRPCP et fichiers associés).
Aucune dépendance Streamlit.
"""

from __future__ import annotations

import pandas as pd


def build_pc_profile(
    codes: pd.Series,
    entreprise: str,
    statut: str,
) -> pd.DataFrame:
    """
    Construit le DataFrame profil PC.

    Retourne
    --------
    DataFrame à 5 colonnes (format DFRXHYBRPCP).
    """
    return pd.DataFrame({
        0: [f"PC_PROFILE_{entreprise}"] * len(codes),
        1: [statut] * len(codes),
        2: [None] * len(codes),
        3: [f"M2_{c}" for c in codes],
        4: ["frxProductCatalog:Online"] * len(codes),
    }).drop_duplicates()


def build_pc_files_data(
    df_profile: pd.DataFrame,
    comptes: pd.Series,
    entreprise: str,
    dstr: str,
) -> dict[str, dict]:
    """
    Génère le contenu binaire des 4 fichiers PC.

    Retourne
    --------
    {filename: {"label": str, "data": bytes, "mime": str}}
    — prêt à être passé au repository de persistence.
    """
    files: dict[str, dict] = {}

    def _add(label: str, content: str | bytes, filename: str, mime: str) -> None:
        data = content.encode() if isinstance(content, str) else content
        files[filename] = {"label": label, "data": data, "filename": filename, "mime": mime}

    _add(
        "⬇️ DFRXHYBRPCP",
        df_profile.to_csv(sep=";", index=False, header=False),
        f"DFRXHYBRPCP{dstr}0000",
        "text/plain",
    )
    _add(
        "⬇️ AFRXHYBRCMP",
        (f"DFRXHYBRCMP{dstr}000068240530IT"
         f"DFRXHYBRCMP{dstr}CCMGHYBFRX                    OK000000"),
        f"AFRXHYBRCMP{dstr}0000.txt",
        "text/plain",
    )
    _add(
        "⬇️ DFRXHYBRCMP",
        (f"PC_{entreprise};PC_{entreprise};PC_PROFILE_{entreprise};"
         f"{','.join(comptes)};frxProductCatalog:Online"),
        f"DFRXHYBRCMP{dstr}0000",
        "text/plain",
    )
    _add(
        "⬇️ AFRXHYBRPCP",
        (f"DFRXHYBRPCP{dstr}000068200117IT"
         f"DFRXHYBRPCP{dstr}RCMRHYBFRX                    OK000000"),
        f"AFRXHYBRPCP{dstr}0000.txt",
        "text/plain",
    )
    return files
