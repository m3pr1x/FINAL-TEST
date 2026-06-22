# -*- coding: utf-8 -*-
"""
services/transforms.py
Transformations atomiques réutilisables — aucune dépendance Streamlit.
"""

from __future__ import annotations
import io
from typing import Tuple

import pandas as pd


def to_m2(s: pd.Series) -> pd.Series:
    """Zero-pad une série de codes M2 à 6 chiffres."""
    return s.astype(str).str.zfill(6)


def sanitize_code(code: str) -> str | None:
    """Valide un code M2 de 5–6 chiffres. Retourne None si invalide."""
    s = str(code).strip()
    if not s.isdigit():
        return None
    return s.zfill(6) if len(s) <= 6 else None


def sanitize_numeric(series: pd.Series, width: int) -> Tuple[pd.Series, pd.Series]:
    """
    Pad numérique à largeur fixe.
    Retourne (série paddée, masque booléen des valeurs invalides).
    """
    s = series.astype(str).str.strip()
    s_pad = s.apply(lambda x: x.zfill(width) if x.isdigit() and len(x) <= width else x)
    bad = ~s_pad.str.fullmatch(rf"\d{{{width}}}")
    return s_pad, bad


def to_xlsx(df: pd.DataFrame) -> bytes:
    """Sérialise un DataFrame en bytes .xlsx (openpyxl)."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    buf.seek(0)
    return buf.getvalue()
