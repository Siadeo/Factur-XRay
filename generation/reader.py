"""Lecture d'un onglet de facture d'exemple ouvert dans Excel (xlwings) → contexte de ``cii.generer_xml``.

Conventions du classeur (voir ``tools/exemples/build.py``) :
- chaque donnée d'en-tête est une cellule portant un nom local à l'onglet, égal à sa clé (« BT_31 ») ;
- chaque donnée répétable est une colonne d'un tableau Excel nommé ``T<n°>_<rôle>`` (``lignes``,
  ``notes``, ``tva``, ``remises``, ``charges``) dont l'en-tête se termine par l'ID : « Quantité (BT-129) ».
On lit les valeurs calculées par Excel, telles qu'affichées.
"""
from __future__ import annotations

from .catalog import TABLEAUX, Catalogue, charger


def role_tableau(nom_tableau: str) -> str | None:
    role = nom_tableau.rsplit("_", 1)[-1]
    return role if role in TABLEAUX else None


def lire(sheet, catalogue: Catalogue | None = None) -> dict:
    catalogue = catalogue or charger()
    contexte = {"entete": {}, **{t: [] for t in TABLEAUX}}
    for nom in sheet.names:
        cle = nom.name.split("!")[-1]
        if cle in catalogue.par_cle:
            contexte["entete"][cle] = nom.refers_to_range.value
    for tableau in sheet.tables:
        role = role_tableau(tableau.name)
        if role is None or tableau.data_body_range is None:
            continue
        champs = [catalogue.depuis_entete(e) for e in tableau.header_row_range.value]
        for ligne in tableau.data_body_range.options(ndim=2).value:
            contexte[role].append({c.cle: v for c, v in zip(champs, ligne) if c is not None})
    return contexte
