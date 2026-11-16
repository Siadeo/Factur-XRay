"""Génération du XML CII sans Excel : catalogue, mise en forme, rendu et validation XSD des 51 instantanés."""
import json
import re
from datetime import datetime
from pathlib import Path

import facturx
import pytest
from lxml import etree

from generation.catalog import TABLEAUX, charger
from generation.cii import GUIDELINES, NIVEAUX_FACTURX, generer_xml, profil
from generation.formats import formater
from generation.reader import role_tableau

ROOT = Path(__file__).resolve().parents[2]
INSTANTANES = sorted((ROOT / "exemples" / "instantanes").glob("*.json"))
GABARIT = (ROOT / "generation" / "templates" / "cii.xml").read_text(encoding="utf-8")
NS = {"ram": "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100",
      "rsm": "urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100"}
CAT = charger()


def test_catalogue():
    assert len(CAT.par_cle) == len(CAT.champs) > 200
    assert all(re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", c.cle) for c in CAT.champs)  # noms Excel valides
    assert {c.role for c in CAT.champs} <= {"entete", *TABLEAUX}
    assert CAT.depuis_entete("Quantité (BT-129)").cle == "BT_129"
    assert CAT.depuis_entete("Réf. (Réf. commande ligne)").cle == "REF_COMMANDE_LIGNE"
    assert CAT.depuis_entete("sans identifiant") is None


def test_gabarit_n_utilise_que_des_cles_du_catalogue():
    cles = set(re.findall(r"\b[hlntrc]\.([A-Za-z0-9_]+)", GABARIT))
    assert cles and cles <= set(CAT.par_cle), cles - set(CAT.par_cle)


@pytest.mark.parametrize("type_valeur, valeur, attendu", [
    ("date", datetime(2026, 3, 16), "20260316"),
    ("date", "2026-03-16", "20260316"),
    ("montant", 2940.0, "2940.00"),
    ("montant", 0.125, "0.13"),
    ("prix", 0.489, "0.4890"),
    ("quantite", 2.5, "2.5000"),
    ("taux", 5.5, "5.50"),
    ("texte", 380.0, "380"),
    ("texte", "  B1 ", "B1"),
    ("texte", "", None),
    ("montant", None, None),
])
def test_mise_en_forme(type_valeur, valeur, attendu):
    assert formater(type_valeur, valeur) == attendu


def test_role_des_tableaux():
    assert role_tableau("T01_lignes") == "lignes"
    assert role_tableau("T30a_tva") == "tva"
    assert role_tableau("Tableau1") is None


def test_un_instantane_par_onglet():
    assert len(INSTANTANES) == 51  # 50 cas, dont le 30 en deux documents (30a, 30b)


@pytest.mark.parametrize("chemin", INSTANTANES, ids=lambda p: p.stem)
def test_xml_valide_au_schema(chemin):
    contexte = json.loads(chemin.read_text(encoding="utf-8"))
    xml, niveau = generer_xml(contexte)
    assert facturx.xml_check_xsd(xml, flavor="factur-x", level=NIVEAUX_FACTURX[niveau])
    racine = etree.fromstring(xml)
    assert racine.findtext(".//rsm:ExchangedDocumentContext/ram:GuidelineSpecifiedDocumentContextParameter/ram:ID",
                           namespaces=NS) == GUIDELINES[niveau]
    # ce qu'on voit est ce qui part dans le XML : le total TTC est recopié tel qu'Excel l'affiche
    ttc = racine.findtext(".//ram:SpecifiedTradeSettlementHeaderMonetarySummation/ram:GrandTotalAmount", namespaces=NS)
    assert ttc == formater("montant", contexte["entete"]["BT_112"])


def test_profil_deduit_des_donnees():
    base = json.loads((ROOT / "exemples" / "instantanes" / "01_Biens.json").read_text(encoding="utf-8"))
    assert profil(base) == "EN16931"
    base["lignes"][0]["EXT_FR_FE_163"] = "DETAIL"
    assert profil(base) == "EXTENDED"
    assert profil({"entete": {"BT_1": "F1", "BT_24": "x"}}) == "EN16931"


def test_cellule_vide_absente_du_xml():
    xml, _ = generer_xml({"entete": {"BT_1": "F1", "BT_3": "380", "BT_5": "EUR", "BT_13": "  "}})
    assert b"BuyerOrderReferencedDocument" not in xml
