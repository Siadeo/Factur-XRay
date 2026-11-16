"""Mise en forme des valeurs Excel telles qu'elles s'écrivent dans le XML CII.

Une seule règle, partagée par la génération de la facture et par l'aperçu des Rayons X :
- date → AAAAMMJJ (format 102) ;
- montant → 2 décimales ; prix unitaire et quantité → 4 décimales ; taux → 2 décimales ;
- texte → tel quel (un nombre entier saisi dans Excel perd son « .0 »).
"""
from __future__ import annotations

from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal

DECIMALES = {"montant": 2, "taux": 2, "prix": 4, "quantite": 4}


def vide(valeur) -> bool:
    return valeur is None or (isinstance(valeur, str) and not valeur.strip())


def _nombre(valeur, decimales: int) -> str:
    quantum = Decimal(1).scaleb(-decimales)
    return str(Decimal(str(valeur).replace(",", ".").replace(" ", "")).quantize(quantum, ROUND_HALF_UP))


def formater(type_valeur: str, valeur) -> str | None:
    """Valeur XML d'une cellule, ou None si elle est vide."""
    if vide(valeur):
        return None
    if type_valeur == "date":
        if isinstance(valeur, (datetime, date)):
            return valeur.strftime("%Y%m%d")
        return str(valeur).replace("-", "").strip()
    if type_valeur in DECIMALES:
        try:
            return _nombre(valeur, DECIMALES[type_valeur])
        except ArithmeticError:
            return str(valeur).strip()  # saisie non numérique : laissée telle quelle
    if isinstance(valeur, bool):
        return "true" if valeur else "false"
    if isinstance(valeur, float) and valeur.is_integer():
        return str(int(valeur))
    return str(valeur).strip()
