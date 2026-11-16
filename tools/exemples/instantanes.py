"""Lit chaque onglet des classeurs d'exemple dans Excel et enregistre son instantané JSON et son XML.

Usage (développement, Excel requis) : ``python tools/exemples/instantanes.py [dossier_xml]``

- ``exemples/instantanes/<onglet>.json`` : les valeurs telles qu'Excel les calcule, lues comme le fait
  l'application (``generation.reader``). Les tests génèrent et valident les 50 XML à partir de ces
  fichiers, sans Excel.
- ``<dossier_xml>/<onglet>.xml`` (facultatif) : le XML CII de chaque cas, par exemple pour un contrôle
  par lot dans Fchecker.
"""
from __future__ import annotations

import json
import sys
from datetime import date, datetime
from pathlib import Path

import xlwings as xw

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from generation.cii import generer_xml  # noqa: E402
from generation.reader import lire  # noqa: E402

EXEMPLES = ROOT / "exemples"
INSTANTANES = EXEMPLES / "instantanes"


def _json(valeur):
    if isinstance(valeur, (datetime, date)):
        return valeur.strftime("%Y-%m-%d")
    return valeur


def main(dossier_xml: Path | None = None):
    INSTANTANES.mkdir(exist_ok=True)
    if dossier_xml:
        dossier_xml.mkdir(parents=True, exist_ok=True)
    app = xw.App(visible=False, add_book=False)
    try:
        for chemin in sorted(EXEMPLES.glob("[A-I]_*.xlsx")):
            book = app.books.open(str(chemin), read_only=True)
            try:
                for sheet in book.sheets:
                    contexte = lire(sheet)
                    contexte = {k: ({c: _json(v) for c, v in d.items()} if isinstance(d, dict)
                                    else [{c: _json(v) for c, v in l.items()} for l in d])
                                for k, d in contexte.items()}
                    texte = json.dumps(contexte, ensure_ascii=False, indent=1, sort_keys=True)
                    (INSTANTANES / f"{sheet.name}.json").write_text(texte + "\n", encoding="utf-8")
                    if dossier_xml:
                        xml, _ = generer_xml(contexte)
                        (dossier_xml / f"{sheet.name}.xml").write_bytes(xml)
                    print(f"{chemin.stem} / {sheet.name}")
            finally:
                book.close()
    finally:
        app.quit()


if __name__ == "__main__":
    main(Path(sys.argv[1]) if len(sys.argv) > 1 else None)
