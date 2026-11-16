"""Catalogue des données d'une facture (``champs.yaml``, extrait des fiches de spécification).

Chaque donnée a un ID (« BT-31 », « EXT-FR-FE-163 »…), une clé (« BT_31 ») qui sert de nom à sa
cellule Excel, un rôle (``entete`` ou le tableau qui la porte : ``notes``, ``lignes``, ``tva``,
``remises``, ``charges``), un type de valeur et son chemin XML. Les données ``fixe`` sont écrites par
le gabarit XML lui-même (constantes, attributs) : elles n'ont pas de cellule.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

import yaml

FICHIER = Path(__file__).with_name("champs.yaml")
PROFIL_ETENDU = "EXTENDED"
TABLEAUX = ("notes", "lignes", "tva", "remises", "charges")
_ID_EN_TETE = re.compile(r"\(([^()]+)\)\s*$")


@dataclass(frozen=True)
class Champ:
    id: str
    cle: str
    libelle: str
    bloc: str
    role: str
    type: str
    profil: str
    xpath: tuple
    fixe: str | None = None

    @property
    def etendu(self) -> bool:
        return self.profil == PROFIL_ETENDU


class Catalogue:
    def __init__(self, champs: list[Champ]):
        self.champs = champs
        self.par_cle = {c.cle: c for c in champs}
        self.par_id = {c.id: c for c in champs}

    def du_role(self, role: str) -> list[Champ]:
        """Données à saisir (hors données fixes) d'un rôle, dans l'ordre des fiches."""
        return [c for c in self.champs if c.role == role and c.fixe is None]

    def depuis_entete(self, entete: str) -> Champ | None:
        """« Quantité facturée (BT-129) » → donnée BT-129 (en-tête de colonne d'un tableau)."""
        m = _ID_EN_TETE.search(str(entete or ""))
        return self.par_id.get(m.group(1).strip()) if m else None

    @staticmethod
    def entete(champ: Champ, titre: str | None = None) -> str:
        return f"{titre or champ.libelle} ({champ.id})"


@lru_cache(maxsize=1)
def charger(fichier: Path = FICHIER) -> Catalogue:
    donnees = yaml.safe_load(Path(fichier).read_text(encoding="utf-8"))
    return Catalogue([Champ(**{**d, "xpath": tuple(d["xpath"])}) for d in donnees])
