# -*- coding: utf-8 -*-
"""
services/io_service.py
Lecture robuste de fichiers CSV / Excel — aucune dépendance Streamlit.
Le cache @st.cache_data est appliqué dans la couche UI (multi_tool_app.py).
"""

from __future__ import annotations
import csv
import io

import pandas as pd


def read_csv(buf: io.BytesIO) -> pd.DataFrame:
    """Détecte automatiquement encodage et séparateur."""
    for enc in ("utf-8", "latin1", "cp1252"):
        buf.seek(0)
        try:
            sample = buf.read(2048).decode(enc, errors="ignore")
            try:
                sep = csv.Sniffer().sniff(sample, delimiters=";,|\t").delimiter
            except csv.Error:
                sep = ";"
            buf.seek(0)
            return pd.read_csv(
                buf, sep=sep, encoding=enc,
                engine="python", on_bad_lines="skip", dtype=str,
            )
        except Exception:
            continue
    raise ValueError("CSV illisible (encodage ou séparateur non détecté)")


def read_file(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """
    Lit un fichier CSV / XLSX / XLS depuis ses bytes bruts.
    Peut être wrappé avec @st.cache_data côté UI pour éviter les re-lectures.
    """
    name = filename.lower()
    if name.endswith(".csv"):
        df = read_csv(io.BytesIO(file_bytes))
    elif name.endswith(".xlsx"):
        df = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl", dtype=str)
    elif name.endswith(".xls"):
        try:
            df = pd.read_excel(io.BytesIO(file_bytes), engine="xlrd", dtype=str)
        except Exception:
            raise ValueError("Fichier .xls illisible — convertissez-le en .xlsx d'abord.")
    else:
        raise ValueError(f"Extension non gérée : {filename}")

    return df.loc[:, ~df.columns.str.match(r"^Unnamed")]
