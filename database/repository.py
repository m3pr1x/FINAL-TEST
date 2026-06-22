# -*- coding: utf-8 -*-
"""
database/repository.py
Interface abstraite de persistence — IResultRepository.

Pour intégrer une base de données (PostgreSQL, Supabase, SQLite) :
1. Créer un nouveau fichier (ex: database/postgres_repository.py)
2. Implémenter IResultRepository
3. Dans multi_tool_app.py, remplacer :
       _repo = SessionRepository()
   par :
       _repo = PostgresRepository(connection_string=os.environ["DATABASE_URL"])

Aucune modification des services/ ni des pages UI n'est nécessaire.
"""

from __future__ import annotations
from abc import ABC, abstractmethod
from typing import Optional

import pandas as pd


class IResultRepository(ABC):
    """
    Contrat de persistence pour les résultats des outils.
    Implémentez cette interface pour connecter n'importe quel backend.
    """

    @abstractmethod
    def save_dataframe(self, tool_id: str, key: str, df: pd.DataFrame) -> None:
        """Persiste un DataFrame résultat pour l'outil `tool_id`."""
        ...

    @abstractmethod
    def load_dataframe(self, tool_id: str, key: str) -> Optional[pd.DataFrame]:
        """Charge un DataFrame précédemment sauvegardé."""
        ...

    @abstractmethod
    def save_file(
        self, tool_id: str, filename: str,
        data: bytes, mime: str, label: str,
    ) -> None:
        """Persiste un fichier généré (bytes) pour l'outil `tool_id`."""
        ...

    @abstractmethod
    def list_files(self, tool_id: str) -> dict[str, dict]:
        """
        Retourne tous les fichiers sauvegardés pour `tool_id`.
        Format : {filename: {"label", "data", "filename", "mime"}}
        """
        ...

    @abstractmethod
    def clear(self, tool_id: str | None = None) -> None:
        """
        Vide les résultats d'un outil (ou de tous si tool_id=None).
        Utilisé par le bouton « Réinitialiser ».
        """
        ...

    # ── Méthode concrète : bulk save depuis un dict de fichiers ──────────────
    def save_files_dict(self, tool_id: str, files: dict[str, dict]) -> None:
        """
        Sauvegarde un dict {filename: {label, data, filename, mime}}
        tel que retourné par pc_service.build_pc_files_data().
        """
        for info in files.values():
            self.save_file(
                tool_id,
                info["filename"],
                info["data"],
                info["mime"],
                info["label"],
            )
