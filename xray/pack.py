"""Paquet « Rayons X » : ce qu'il faut savoir d'une cellule des factures d'exemple.

- le catalogue (``generation/champs.yaml``) : ID, libellé, bloc, profil et chemin XML de chaque donnée ;
- les fiches des cas (``exemples/fiches.yaml``) : descriptif, points d'attention et commentaires par donnée.
Python pur : aucun accès à Excel.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from functools import lru_cache

import yaml

from generation.catalog import Catalogue, Champ, charger
from tools.utils import resource_path

TRADE_PREFIX = "/rsm:CrossIndustryInvoice/rsm:SupplyChainTradeTransaction/"
COULEURS = {  # fond par bloc (mode couleurs) : une couleur par bloc, dans l'ordre de la légende
    "En-tête – contexte": "6A89CC",
    "En-tête – document et notes": "F8C291",
    "Vendeur": "82CCDD",
    "Acheteur": "B8E994",
    "Représentant fiscal du vendeur": "3C6382",
    "Tiers (agents, facturant, payeur)": "4A69BD",
    "Références et pièces jointes": "FA983A",
    "Livraison": "38ADA9",
    "Facture antérieure": "E55039",
    "Bénéficiaire et paiement": "78E08F",
    "Période de facturation": "079992",
    "Remises et charges du document": "E58E26",
    "Ventilation de la TVA": "60A3BC",
    "Totaux": "F6B93B",
    "Lignes de facture": "FAD390",
    "Spécifique": "EB2F06",
}
COULEUR_DEFAUT = "F2F2F2"


def _luminance(hexa: str) -> float:
    """Luminance relative (WCAG) d'une couleur « RRGGBB »."""
    def canal(c):
        c = int(c, 16) / 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (canal(hexa[i:i + 2]) for i in (0, 2, 4))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def texte_sur(fond: str) -> str:
    """Couleur de texte la plus lisible (« #000000 » ou « #FFFFFF ») sur un fond « #RRGGBB »."""
    lum = _luminance(fond.lstrip("#"))
    return "#FFFFFF" if 1.05 / (lum + 0.05) > (lum + 0.05) / 0.05 else "#000000"


@dataclass
class Cellule:
    champ: Champ

    @property
    def cle(self) -> str:
        return self.champ.cle

    @property
    def titre(self) -> str:
        return f"{self.champ.id} · {self.champ.libelle}"


class Paquet:
    def __init__(self, catalogue: Catalogue, fiches: dict):
        self.catalogue, self.fiches = catalogue, fiches

    def cellule(self, champ: Champ) -> Cellule:
        return Cellule(champ)

    def fiche(self, onglet: str) -> dict | None:
        """Fiche d'un onglet (« 30a_… » et « 30b_… » partagent la fiche du cas 30)."""
        if onglet in self.fiches:
            return self.fiches[onglet]
        numero = re.match(r"^(\d+)", onglet)
        return next((f for k, f in self.fiches.items() if numero and k.startswith(numero.group(1) + "_")), None)

    def commentaire(self, onglet: str, champ: Champ) -> str:
        return ((self.fiche(onglet) or {}).get("commentaires") or {}).get(champ.id, "")

    @staticmethod
    def couleur(champ: Champ) -> str:
        return "#" + COULEURS.get(champ.bloc, COULEUR_DEFAUT)

    @staticmethod
    def legende(blocs) -> list[tuple[str, str]]:
        """[(bloc, couleur)] des blocs donnés, dans l'ordre de COULEURS."""
        presents = set(blocs)
        return [(b, "#" + c) for b, c in COULEURS.items() if b in presents]


@lru_cache(maxsize=1)
def charger_paquet() -> Paquet:
    fiches = yaml.safe_load(resource_path("exemples/fiches.yaml").read_text(encoding="utf-8"))
    return Paquet(charger(), fiches)


def short_xpath(xpath: str) -> str:
    return "…/" + xpath[len(TRADE_PREFIX):] if xpath.startswith(TRADE_PREFIX) else xpath


def note_text(paquet: Paquet, cellule: Cellule, onglet: str = "") -> str:
    """Texte de la note Excel affichée au survol d'une cellule."""
    champ = cellule.champ
    lignes = [cellule.titre, f"{champ.bloc} · profil {champ.profil}"]
    lignes += [f"XML : {short_xpath(x)}" for x in champ.xpath]
    commentaire = paquet.commentaire(onglet, champ)
    if commentaire:
        lignes.append(f"Dans ce cas : {commentaire}")
    return "\n".join(lignes)
