"""Classeurs de factures d'exemple dans Excel (xlwings) : ouverture en lecture seule, onglet actif, PDF."""
from __future__ import annotations

import re
from pathlib import Path

import xlwings as xw
import yaml

from tools.utils import resource_path

EXEMPLES = resource_path("exemples")
FACTURES = resource_path("factures")
_CLASSEUR = re.compile(r"^(?:\d\d_Index|[A-Z]_\w+)\.xlsx$", re.IGNORECASE)


def classeurs() -> dict[str, Path]:
    """Libellé du menu « Ouvrir » → chemin : « Index des cas », puis « A. Factures de base »…"""
    fiches = yaml.safe_load((EXEMPLES / "fiches.yaml").read_text(encoding="utf-8")) if (EXEMPLES / "fiches.yaml").exists() else {}
    familles = {f["classeur"]: f["famille"] for f in fiches.values()}
    menu = {}
    for chemin in sorted(EXEMPLES.glob("*.xlsx")):
        if chemin.stem.endswith("Index"):
            menu["Index des cas"] = chemin
        elif re.match(r"^[A-Z]_", chemin.stem):
            menu[familles.get(chemin.stem, chemin.stem.replace("_", " "))] = chemin
    return menu


def excel() -> xw.App:
    """Instance Excel active, ou une nouvelle instance visible."""
    app = xw.apps.active if xw.apps.count else xw.App(visible=True, add_book=False)
    app.visible = True
    return app


def ouvrir(chemin: Path) -> xw.Book:
    """Ouvre le classeur en lecture seule (les exemples ne sont jamais écrasés) ou l'active s'il est ouvert."""
    app = excel()
    for book in app.books:
        if book.name.lower() == chemin.name.lower():
            book.activate()
            return book
    book = app.books.open(str(chemin), read_only=True)
    book.activate()
    return book


def est_exemple(book: xw.Book) -> bool:
    return bool(_CLASSEUR.match(book.name)) and not book.name.lower().endswith("index.xlsx")


def onglet_actif() -> xw.Sheet:
    """Onglet actif d'Excel, s'il appartient à un classeur de factures d'exemple."""
    if not xw.apps.count or xw.apps.active.books.count == 0:
        raise ValueError("Ouvrez d'abord un classeur d'exemples (menu « Ouvrir »).")
    book = xw.apps.active.books.active
    if not est_exemple(book):
        raise ValueError(f"Le classeur actif ({book.name}) n'est pas un classeur de factures d'exemple.")
    return book.sheets.active


def exporter_pdf(sheet: xw.Sheet, chemin: Path) -> None:
    """Exporte la zone d'impression de l'onglet (la zone technique, à droite, n'est pas imprimée)."""
    chemin.parent.mkdir(parents=True, exist_ok=True)
    sheet.to_pdf(path=str(chemin), show=False)


def fermer_exemples() -> None:
    """Ferme, sans enregistrer, les classeurs d'exemples ouverts (les autres classeurs restent ouverts)."""
    if not xw.apps.count:
        return
    for app in xw.apps:
        for book in list(app.books):
            if _CLASSEUR.match(book.name):
                try:
                    book.close()
                except Exception:
                    pass
