# -*- coding: utf-8 -*-
"""
services/m2_service.py
Logique métier : Mise à jour M2 et Appairage client.
Entrée/sortie : DataFrames pandas — aucune dépendance Streamlit.
"""

from __future__ import annotations
from typing import List

import pandas as pd

from services.transforms import to_m2


def build_m2_update(
    old_df: pd.DataFrame,
    new_df: pd.DataFrame,
    ref_i_old: int,
    val_i_old: int,
    ref_i_new: int,
    val_i_new: int,
) -> pd.DataFrame:
    """
    Mappe M2_ancien → M2_nouveau en croisant plan d'offre N‑1 et N.

    Paramètres
    ----------
    old_df / new_df : DataFrames lus depuis les fichiers uploadés
    ref_i_old/new   : index 0-based de la colonne Référence produit
    val_i_old/new   : index 0-based de la colonne Code M2

    Retourne
    --------
    DataFrame(M2_nouveau, M2_ancien)
    """
    def _extract(df: pd.DataFrame, ri: int, vi: int,
                 ref_lbl: str, val_lbl: str) -> pd.DataFrame:
        sub = df.iloc[:, [ri, vi]].copy()
        sub.columns = [ref_lbl, val_lbl]
        sub[val_lbl] = to_m2(sub[val_lbl])
        return sub

    old = _extract(old_df, ref_i_old, val_i_old, "Ref", "M2_ancien")
    new = _extract(new_df, ref_i_new, val_i_new, "Ref", "M2_nouveau")

    merged = new.merge(old[["Ref", "M2_ancien"]], on="Ref", how="left")
    return (
        merged.groupby("M2_nouveau")["M2_ancien"]
              .agg(lambda s: s.value_counts().idxmax() if s.notna().any() else pd.NA)
    ).reset_index()


def build_appairage(
    old_df: pd.DataFrame,
    new_df: pd.DataFrame,
    map_df: pd.DataFrame,
    ref_i_old: int,
    val_i_old: int,
    ref_i_new: int,
    val_i_new: int,
    ref_i_map: int,
    val_i_map: int,
    extra_cols: List[str],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Construit l'appairage M2_nouveau → Code famille client.

    Retourne
    --------
    (fam_df, missing_df)
    fam_df     : DataFrame(M2_nouveau, Code_famille_Client)
    missing_df : lignes sans famille client associée
    """
    old = old_df.iloc[:, [ref_i_old, val_i_old]].copy()
    old.columns = ["Ref", "M2_ancien"]
    old["M2_ancien"] = to_m2(old["M2_ancien"])

    new_full = new_df.copy()
    ref_series = new_full.iloc[:, ref_i_new].astype(str).str.strip()
    m2_series  = new_full.iloc[:, val_i_new]
    new_full.insert(0, "Ref", ref_series)
    new_full.insert(1, "M2_nouveau", to_m2(m2_series))

    mapping = map_df.iloc[:, [ref_i_map, val_i_map]].copy()
    mapping.columns = ["M2_ancien", "Code_famille_Client"]
    mapping["M2_ancien"] = to_m2(mapping["M2_ancien"])
    old["M2_ancien"]     = to_m2(old["M2_ancien"])

    merged = (
        new_full
        .merge(old[["Ref", "M2_ancien"]], on="Ref", how="left")
        .merge(mapping, on="M2_ancien", how="left")
    )

    fam = (
        merged.groupby("M2_nouveau")["Code_famille_Client"]
              .agg(lambda s: s.value_counts().idxmax() if s.notna().any() else pd.NA)
    ).reset_index()

    missing = fam[fam["Code_famille_Client"].isna()].copy()
    if extra_cols:
        keep = [c for c in extra_cols if c in merged.columns]
        missing = missing.merge(
            merged[["M2_nouveau"] + keep].drop_duplicates(),
            on="M2_nouveau", how="left",
        )
    return fam, missing
