"""Rayons X : parties sans Excel (paquet, notes, aperçu XML)."""
from datetime import datetime

import pytest
from lxml import etree

from generation.catalog import charger
from xray.fragment import fragment
from xray.pack import COULEURS, charger_paquet, note_text, short_xpath, texte_sur

PAQUET = charger_paquet()
CAT = charger()


def cellule(cle):
    return PAQUET.cellule(CAT.par_cle[cle])


def test_fiches_et_onglets_derives():
    assert len(PAQUET.fiches) == 50
    assert PAQUET.fiche("01_Biens")["profil"] == "EN16931"
    # 30a et 30b partagent la fiche du cas 30
    assert PAQUET.fiche("30a_Netting_autofacture") is PAQUET.fiche("30_Netting_facture_croisee")
    assert PAQUET.fiche("inconnu") is None


def test_commentaire_propre_au_cas():
    assert "débits" in PAQUET.commentaire("03_Services_debits", CAT.par_cle["BT_8"])
    assert PAQUET.commentaire("01_Biens", CAT.par_cle["BT_8"]) == ""


def test_note_au_survol():
    texte = note_text(PAQUET, cellule("BT_31"), "01_Biens")
    assert texte.startswith("BT-31 · Numéro de TVA du vendeur")
    assert "SpecifiedTaxRegistration" in texte
    assert "Dans ce cas : Égal au total TTC" in note_text(PAQUET, cellule("BT_113"), "05_Deja_payee")


def test_xpath_raccourci():
    assert short_xpath(CAT.par_cle["BT_27"].xpath[0]).startswith("…/ram:ApplicableHeaderTradeAgreement")


def test_couleur_par_bloc():
    assert PAQUET.couleur(CAT.par_cle["BT_27"]) == "#82CCDD"
    assert PAQUET.couleur(CAT.par_cle["BT_129"]) == "#FAD390"


def test_une_couleur_distincte_par_bloc_du_catalogue():
    blocs = {c.bloc for c in CAT.par_cle.values()}
    assert blocs <= set(COULEURS)  # aucun bloc sans couleur
    assert len(set(COULEURS.values())) == len(COULEURS)  # aucune couleur partagée


def test_legende_dans_l_ordre_sans_doublon():
    assert PAQUET.legende(["Totaux", "Vendeur", "Totaux"]) == [("Vendeur", "#82CCDD"), ("Totaux", "#F6B93B")]
    assert PAQUET.legende([]) == []


@pytest.mark.parametrize("fond, texte", [("#FAD390", "#000000"), ("#3C6382", "#FFFFFF"), ("#4A69BD", "#FFFFFF")])
def test_texte_lisible_sur_le_fond(fond, texte):
    assert texte_sur(fond) == texte


@pytest.mark.parametrize("cle, valeur, attendu", [
    ("BT_31", "FR88100000009", ">FR88100000009<"),
    ("BT_2", datetime(2026, 11, 18), ">20261118<"),
    ("BT_112", 2461.2, ">2461.20<"),
    ("BT_152", 5.5, ">5.50<"),
    ("BT_146", 0.489, ">0.4890<"),
])
def test_fragment_valeur_mise_en_forme(cle, valeur, attendu):
    champ = CAT.par_cle[cle]
    xml = fragment(champ.xpath[0], valeur, champ.type)
    assert attendu in xml
    etree.fromstring(_wrap(xml))  # fragment bien formé


def test_fragment_date_format_102():
    champ = CAT.par_cle["BT_2"]
    assert 'format="102"' in fragment(champ.xpath[0], datetime(2026, 3, 16), champ.type)


def test_fragment_attribut():
    champ = CAT.par_cle["BT_130"]  # …/ram:BilledQuantity/@unitCode
    xml = fragment(champ.xpath[0], "HUR", champ.type)
    assert 'unitCode="HUR"' in xml
    etree.fromstring(_wrap(xml))


def test_fragment_code_type_impose():
    champ = CAT.par_cle["BT_122"]  # AdditionalReferencedDocument[ram:TypeCode='916']
    xml = fragment(champ.xpath[0], "ATT-01", champ.type)
    assert "<ram:TypeCode>916</ram:TypeCode>" in xml and ">ATT-01<" in xml


def _wrap(xml):
    ns = 'xmlns:rsm="urn:rsm" xmlns:ram="urn:ram" xmlns:udt="urn:udt" xmlns:qdt="urn:qdt"'
    return f"<racine {ns}>{xml}</racine>"
