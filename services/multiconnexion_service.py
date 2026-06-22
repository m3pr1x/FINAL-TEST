# -*- coding: utf-8 -*-
"""
services/multiconnexion_service.py
Logique métier : Multiconnexion — tables PF1→PF6.
Aucune dépendance Streamlit.
"""

from __future__ import annotations
import re

import pandas as pd

try:
    from postal.parser import parse_address  # type: ignore
    _USE_POSTAL = True
except ImportError:
    _USE_POSTAL = False


def split_address(addr: str) -> dict:
    """
    Décompose une adresse libre en champs structurés.
    Utilise libpostal si disponible, sinon fallback regex.
    Retourne toujours {'num', 'voie', 'cp', 'ville', 'pays'}.
    """
    parts = {"num": "", "voie": "", "cp": "", "ville": "", "pays": "FR"}
    if _USE_POSTAL:
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

    pattern = re.compile(
        r"^\s*(?P<num>[\d\w\-]*)\s*(?P<voie>[^,]+?)[,\s]+(?P<cp>\d{2}\s?\d{3})\s+(?P<ville>.+?)\s*$",
        re.IGNORECASE,
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
    integration_type: str = "OCI",
) -> list[pd.DataFrame]:
    """
    Construit PF1→PF5 (+ PF6 si cXML) en O(n) via accumulation de listes.

    Lève
    ----
    ValueError si des colonnes obligatoires sont absentes du DataFrame source.
    """
    REQUIRED = ["Numéro de compte", "Raison sociale", "Adresse", "Code agence"]
    missing = [c for c in REQUIRED if c not in df_src.columns]
    if missing:
        raise ValueError(f"Colonnes manquantes dans le fichier : {', '.join(missing)}")

    pf1, pf2, pf3, pf4, pf5, pf6 = [], [], [], [], [], []

    for _, row in df_src.iterrows():
        account = row["Numéro de compte"]
        company = row["Raison sociale"]
        branch  = row["Code agence"]
        addr    = split_address(str(row.get("Adresse") or "").strip())

        pf1.append({
            "uid": account, "name": company, "locName": company,
            "CXmIAssignedConfiguration": f"frx-variant-{entreprise}-configuration-set",
            "pcCompoundProfile": f"PC_{entreprise}",
            "ViewMasterCatalog": view_master_catalog,
        })
        pf2.append({
            "B2B Unit": account,
            "ADRESSE / Numéro de rue": addr["num"],
            "ADRESSE / rue": addr["voie"],
            "ADRESSE / Code postal": addr["cp"],
            "ADRESSE / Ville": addr["ville"],
            "ADRESSE / Pays/Région": addr["pays"],
            "INFORMATIONS D'ADRESSE SUPPLÉMENTAIRES / Téléphone 1": "",
        })
        pf3.append({
            "B2BUnitID": account,
            "itemtype": "PunchoutAccountAndBranchAssociation",
            "managingBranches": branch,
            "punchoutUserID": punchout_user_id,
            "sealed": "false",
        })
        pf4.append({
            "aliasName": branch, "branch": branch,
            "punchoutUserID": punchout_user_id, "sealed": "false",
        })
        pf5.append({"B2BUnitID": account, "punchoutUserID": punchout_user_id})

        if integration_type == "cXML":
            pf6.append({"number": account, "domain": domain, "identity": identity})

    tables = [pd.DataFrame(r) for r in [pf1, pf2, pf3, pf4, pf5]]
    if integration_type == "cXML":
        tables.append(pd.DataFrame(pf6))
    return tables
