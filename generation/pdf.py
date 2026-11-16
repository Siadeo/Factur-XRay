"""PDF exporté par Excel → Factur-X (PDF/A-3 avec le XML CII embarqué).

Aucune validation : le XML est embarqué tel qu'il est produit (le contrôle se fait dans Fchecker).
"""
from __future__ import annotations

import logging
import re
from pathlib import Path

import pikepdf
from facturx import generate_from_file

from tools.utils import resource_path

from .cii import NIVEAUX_FACTURX

logger = logging.getLogger(__name__)


def ajouter_output_intent(pdf_path: Path) -> None:
    """Ajoute un OutputIntent sRGB (profil ICC embarqué) : Excel produit du DeviceRGB sans OutputIntent,
    ce qui est interdit en PDF/A."""
    icc = resource_path("assets/sRGB.icc").read_bytes()
    with pikepdf.open(str(pdf_path), allow_overwriting_input=True) as pdf:
        if len(pdf.Root.get("/OutputIntents", [])) > 0:
            return
        profil = pikepdf.Stream(pdf, icc)
        profil.N = 3
        intent = pdf.make_indirect(pikepdf.Dictionary(
            Type=pikepdf.Name.OutputIntent, S=pikepdf.Name("/GTS_PDFA1"),
            OutputConditionIdentifier=pikepdf.String("sRGB"), Info=pikepdf.String("sRGB IEC61966-2.1"),
            DestOutputProfile=profil))
        pdf.Root.OutputIntents = pikepdf.Array([intent])
        pdf.save(str(pdf_path))


def aligner_producteur(pdf_path: Path) -> None:
    """La lib factur-x écrit son producteur dans le XMP mais garde celui d'Excel dans le dictionnaire Info :
    les deux doivent être identiques en PDF/A."""
    with pikepdf.open(str(pdf_path), allow_overwriting_input=True) as pdf:
        xmp = pdf.Root.Metadata.read_bytes().decode("utf-8")
        match = re.search(r"<pdf:Producer>(.*?)</pdf:Producer>", xmp)
        if match:
            pdf.docinfo["/Producer"] = match.group(1)
            pdf.save(str(pdf_path))


def facturx(pdf_path: Path, xml: bytes, profil: str) -> None:
    """Transforme sur place le PDF d'Excel en Factur-X du profil donné (EN16931 ou EXTENDED)."""
    ajouter_output_intent(pdf_path)
    generate_from_file(str(pdf_path), xml, flavor="factur-x", level=NIVEAUX_FACTURX[profil],
                       check_xsd=False, check_schematron=False)
    aligner_producteur(pdf_path)
