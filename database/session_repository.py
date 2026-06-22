# -*- coding: utf-8 -*-
"""
database/session_repository.py
Implémentation actuelle : persistence en RAM via st.session_state.

Pour migrer vers PostgreSQL / Supabase / SQLite :
  → Créer database/postgres_repository.py (ou supabase_repository.py)
  → Implémenter IResultRepository
  → Remplacer SessionRepository() dans multi_tool_app.py
  Aucun autre fichier à modifier.
"""

from __future__ import annotations
from typing import Optional

import pandas as pd
import streamlit as st

from database.repository import IResultRepository


class SessionRepository(IResultRepository):
    """
    Stocke tout en st.session_state (mémoire de la session navigateur).
    Données perdues à la fermeture ou à la réinitialisation.
    """

    # ── Clés internes ─────────────────────────────────────────────────────────
    def _df_key(self, tool_id: str, key: str) -> str:
        return f"{tool_id}__{key}__df"

    def _files_key(self, tool_id: str) -> str:
        return f"{tool_id}__files"

    # ── DataFrames ────────────────────────────────────────────────────────────
    def save_dataframe(self, tool_id: str, key: str, df: pd.DataFrame) -> None:
        st.session_state[self._df_key(tool_id, key)] = df

    def load_dataframe(self, tool_id: str, key: str) -> Optional[pd.DataFrame]:
        return st.session_state.get(self._df_key(tool_id, key))

    # ── Fichiers ──────────────────────────────────────────────────────────────
    def save_file(self, tool_id: str, filename: str,
                  data: bytes, mime: str, label: str) -> None:
        fkey = self._files_key(tool_id)
        st.session_state.setdefault(fkey, {})
        st.session_state[fkey][filename] = {
            "label": label, "data": data, "filename": filename, "mime": mime,
        }

    def list_files(self, tool_id: str) -> dict[str, dict]:
        store = st.session_state.get(self._files_key(tool_id), {})
        # compatibilité si l'ancienne structure était une liste
        if isinstance(store, list):
            store = {item["filename"]: item for item in store}
            st.session_state[self._files_key(tool_id)] = store
        return store

    # ── Reset ─────────────────────────────────────────────────────────────────
    def clear(self, tool_id: str | None = None) -> None:
        if tool_id is None:
            st.session_state.clear()
        else:
            prefix = f"{tool_id}__"
            to_del = [k for k in st.session_state if k.startswith(prefix)]
            for k in to_del:
                del st.session_state[k]
