"""Données d'un onglet → XML CII (Factur-X).

Le contexte vient de ``reader.lire`` (ou d'un instantané JSON) :
``{"entete": {clé: valeur}, "notes": [{clé: valeur}], "lignes": [...], "tva": [...], "remises": [...],
"charges": [...]}``. Les valeurs sont brutes (telles qu'Excel les affiche) : on les met en forme ici,
sans aucun recalcul. Ce qu'on voit dans le classeur est ce qui part dans le XML.
"""
from __future__ import annotations

from pathlib import Path

from jinja2 import Environment, FileSystemLoader, select_autoescape
from lxml import etree

from .catalog import TABLEAUX, Catalogue, charger
from .formats import formater

GUIDELINES = {
    "EN16931": "urn:cen.eu:en16931:2017",
    "EXTENDED": "urn:cen.eu:en16931:2017#conformant#urn:factur-x.eu:1p0:extended",
}
NIVEAUX_FACTURX = {"EN16931": "en16931", "EXTENDED": "extended"}

_env = Environment(
    loader=FileSystemLoader(Path(__file__).with_name("templates")),
    autoescape=select_autoescape(["xml"]),
)


def _mettre_en_forme(donnees: dict, catalogue: Catalogue) -> dict:
    """Ne garde que les cellules remplies, mises en forme selon le type de la donnée."""
    sortie = {}
    for cle, valeur in donnees.items():
        champ = catalogue.par_cle.get(cle)
        if champ is None:
            continue
        texte = formater(champ.type, valeur)
        if texte is not None:
            sortie[cle] = texte
    return sortie


def preparer(contexte: dict, catalogue: Catalogue | None = None) -> dict:
    catalogue = catalogue or charger()
    prepare = {"h": _mettre_en_forme(contexte.get("entete", {}), catalogue)}
    for tableau in TABLEAUX:
        lignes = (_mettre_en_forme(l, catalogue) for l in contexte.get(tableau, []))
        prepare[tableau] = [l for l in lignes if l]
    return prepare


def profil(contexte: dict, catalogue: Catalogue | None = None) -> str:
    """EXTENDED dès qu'une donnée propre au profil étendu est remplie, sinon EN16931."""
    catalogue = catalogue or charger()
    prepare = preparer(contexte, catalogue)
    blocs = [prepare["h"], *(l for t in TABLEAUX for l in prepare[t])]
    etendu = any(catalogue.par_cle[cle].etendu for bloc in blocs for cle in bloc)
    return "EXTENDED" if etendu else "EN16931"


def generer_xml(contexte: dict, catalogue: Catalogue | None = None) -> tuple[bytes, str]:
    """(XML CII indenté, profil EN16931 / EXTENDED)."""
    catalogue = catalogue or charger()
    niveau = profil(contexte, catalogue)
    texte = _env.get_template("cii.xml").render(guideline=GUIDELINES[niveau], **preparer(contexte, catalogue))
    racine = etree.fromstring(texte.encode("utf-8"), etree.XMLParser(remove_blank_text=True))
    return etree.tostring(racine, xml_declaration=True, encoding="UTF-8", pretty_print=True), niveau
