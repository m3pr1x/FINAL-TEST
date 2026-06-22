# -*- coding: utf-8 -*-
"""
services/classification_service.py
Logique métier : Classification Code (DFRXHYBRCMR / AFRXHYBRCMR).
Aucune dépendance Streamlit.
"""

from __future__ import annotations

import pandas as pd

from services.transforms import sanitize_code


def build_classification_output(
    pair_df: pd.DataFrame,
    idx_m2: int,
    idx_fam: int,
    entreprise: str,
    dstr: str,
) -> tuple[pd.DataFrame, bytes, str]:
    """
    Construit les fichiers DFRXHYBRCMR et AFRXHYBRCMR.

    Paramètres
    ----------
    pair_df    : DataFrame source (appairage M2 ↔ famille client)
    idx_m2     : index 1-based de la colonne Code M2
    idx_fam    : index 1-based de la colonne Code famille client
    entreprise : libellé client
    dstr       : date au format YYMMDD

    Retourne
    --------
    (df_out, dfrx_bytes, afrx_ack_str)

    Lève
    ----
    ValueError  si des codes M2 sont invalides (avec la liste des codes fautifs)
    IndexError  si un index de colonne est hors plage
    """
    col_m2  = pair_df.columns[idx_m2  - 1]
    col_fam = pair_df.columns[idx_fam - 1]

    df_out = pair_df[[col_fam, col_m2]].rename(
        columns={col_fam: "Code famille Client", col_m2: "M2"}
    ).copy()

    raw_m2    = df_out["M2"].astype(str).str.strip()
    sanitized = raw_m2.apply(sanitize_code)
    bad_mask  = sanitized.isna()
    if bad_mask.any():
        bad_samples = raw_m2[bad_mask].head(10).tolist()
        raise ValueError(
            f"{bad_mask.sum()} code(s) M2 invalide(s) — 5 ou 6 chiffres attendus. "
            f"Exemples : {bad_samples}"
        )

    df_out["M2"]          = sanitized.map(lambda x: f"M2_{x}")
    df_out["_reserved"]   = None
    df_out["Entreprises"] = entreprise
    df_out.insert(0, "Empty", "")
    df_out = df_out[["Empty", "Code famille Client", "_reserved", "Entreprises", "M2"]]

    dfrx_bytes = df_out.to_csv(sep="\t", index=False, header=False).encode()
    afrx_ack   = (
        f"DFRXHYBRCMR{dstr}000068230116IT"
        f"DFRXHYBRCMR{dstr}RCMRHYBFRX                    OK000000"
    )
    return df_out, dfrx_bytes, afrx_ack
