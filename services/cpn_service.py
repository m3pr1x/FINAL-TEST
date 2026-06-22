# -*- coding: utf-8 -*-
"""
services/cpn_service.py
Logique métier : CPN — produit cartésien Référence × Comptes.
Aucune dépendance Streamlit.
"""

from __future__ import annotations

import pandas as pd


def build_cpn(
    series_int: pd.Series,
    series_cli_prod: pd.Series,
    series_cli_acc: pd.Series,
) -> pd.DataFrame:
    """
    Calcule le produit cartésien (int_ref, cli_prod) × cli_acc.
    Le zip garantit l'alignement int_ref ↔ cli_prod (même ligne du fichier source).

    Paramètres
    ----------
    series_int      : Références internes (8 chiffres), index réinitialisé
    series_cli_prod : Codes produit client, alignés avec series_int
    series_cli_acc  : Numéros de compte du périmètre

    Retourne
    --------
    DataFrame(CustomerItemId, AccountNumber, InternalItemID)
    """
    rows = [
        {
            "CustomerItemId": cli_prod,
            "AccountNumber":  acc,
            "InternalItemID": int_ref,
        }
        for int_ref, cli_prod in zip(series_int, series_cli_prod)
        for acc in series_cli_acc
    ]
    return pd.DataFrame(rows)


def build_cpn_ack(dstr: str) -> str:
    """Génère la chaîne ACK AFRXHYBCPNA."""
    return (
        f"DFRXHYBCPNA{dstr}000148250201IT"
        f"DFRXHYBCPNA{dstr}CPNAHYBFRX                    OK000000"
    )
