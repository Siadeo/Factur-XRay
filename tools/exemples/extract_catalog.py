"""Extrait le catalogue des données et les fiches des 50 cas depuis les classeurs de spécification.

Usage (développement) : ``python tools/exemples/extract_catalog.py ["D:\\__E-invoicing\\ez facture new"]``

Produit :
- ``generation/champs.yaml`` : une entrée par donnée (ID, clé de cellule Excel, libellé, bloc, XPath CII,
  profil minimal, type de valeur, rôle : en-tête ou tableau, valeur fixe éventuelle) ;
- ``exemples/fiches.yaml`` : l'en-tête de chaque cas (famille, descriptif, codes, profil, points
  d'attention) et ses commentaires par donnée.

Libellés : les libellés français de la norme EN 16931 sont protégés. Pour les BT, on utilise le glossaire
de Fchecker (traduit depuis l'anglais, EUPL-1.2) ou, à défaut, ``LIBELLES`` ci-dessous (traduction maison
depuis les noms anglais). Le libellé de la fiche n'est conservé que pour les données hors norme (EXT-FR,
codes techniques).
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

import openpyxl
import yaml

ROOT = Path(__file__).resolve().parents[2]
SPEC_DIR = Path(r"D:\__E-invoicing\ez facture new")
GLOSSAIRE = ROOT.parent / "fchecker" / "src" / "fchecker" / "core" / "catalog" / "en16931_fr.yaml"
TRADE = "/rsm:CrossIndustryInvoice/rsm:SupplyChainTradeTransaction/"
PREMIERE_LIGNE = 17

# ID non normalisés de la fiche → clé de cellule
CLES = {
    "Agent acheteur": "AGENT_ACHETEUR",
    "Facturant": "FACTURANT",
    "Facturé à": "FACTURE_A",
    "Réf. commande ligne": "REF_COMMANDE_LIGNE",
    "Réf. livraison ligne": "REF_LIVRAISON_LIGNE",
}

# Données que le gabarit XML écrit lui-même (constantes, attributs, éléments conteneurs) : pas de cellule.
FIXES = {
    "BT-2-1": "102", "BT-72-1": "102", "BT-26-1": "102", "BT-9-1": "102", "BT-73-1": "102",
    "BT-74-1": "102", "BT-134-1": "102", "BT-135-1": "102", "EXT-FR-FE-138-1": "102",
    "EXT-FR-FE-158-1": "102",
    "BT-30-1": "0002", "BT-47-1": "0002",
    "BT-31-0": "VA", "BT-48-0": "VA", "BT-63-0": "VA",
    "BT-95-0": "VAT", "BT-102-0": "VAT", "BT-118-0": "VAT", "BT-151-0": "VAT",
    "BG-20-1": "false", "BG-21-1": "true", "BG-27-1": "false", "BG-28-1": "true", "BT-147-02": "false",
    "BT-122-0": "916", "BT-17-0": "50", "BT-18-0": "130", "BT-128-0": "130",
    "BT-11-0": "Project reference",
    "BT-110-1": "devise de la facture (BT-5)", "BT-111-1": "EUR",
    "BT-127-00": "", "EXT-FR-FE-BG-06": "", "EXT-FR-FE-BG-10": "", "EXT-FR-FE-150": "",
    "EXT-FR-FE-BG-11": "", "EXT-FR-FE-158-0": "",
}

TABLES = {
    "notes": {"BT-21", "BT-22", "BT-21o", "BT-22o"},
    "remises": {"BT-92", "BT-93", "BT-94", "BT-95", "BT-96", "BT-97", "BT-98"},
    "charges": {"BT-99", "BT-100", "BT-101", "BT-102", "BT-103", "BT-104", "BT-105",
                "EXT-FR-FE-189", "EXT-FR-FE-190"},
    "lignes": {"Réf. commande ligne", "Réf. livraison ligne"},
}
BLOC_TABLE = {"Ventilation de la TVA": "tva", "Lignes de facture": "lignes"}

# Traduction maison (depuis les noms anglais des BT) pour les BT absents du glossaire Fchecker.
LIBELLES = {
    "BT-9": "Date d'échéance", "BT-10": "Référence de l'acheteur", "BT-11": "Référence du projet",
    "BT-12": "Référence du contrat", "BT-13": "Référence du bon de commande",
    "BT-14": "Référence de la commande client (vendeur)", "BT-15": "Référence de l'avis de réception",
    "BT-16": "Référence de l'avis d'expédition", "BT-18": "Identifiant de l'objet facturé",
    "BT-19": "Référence comptable de l'acheteur", "BT-20": "Conditions de paiement",
    "BT-21": "Code du sujet de la note", "BT-22": "Note de facture", "BT-23": "Cadre de facturation",
    "BT-26": "Date de la facture antérieure", "BT-28": "Nom commercial du vendeur",
    "BT-33": "Informations juridiques complémentaires du vendeur", "BT-35": "Adresse du vendeur, ligne 1",
    "BT-36": "Adresse du vendeur, ligne 2", "BT-37": "Ville du vendeur", "BT-38": "Code postal du vendeur",
    "BT-41": "Contact chez le vendeur", "BT-42": "Téléphone du contact vendeur",
    "BT-43": "Courriel du contact vendeur", "BT-46": "Identifiant de l'acheteur",
    "BT-50": "Adresse de l'acheteur, ligne 1", "BT-51": "Adresse de l'acheteur, ligne 2",
    "BT-52": "Ville de l'acheteur", "BT-53": "Code postal de l'acheteur",
    "BT-56": "Contact chez l'acheteur", "BT-57": "Téléphone du contact acheteur",
    "BT-58": "Courriel du contact acheteur", "BT-60": "Identifiant du bénéficiaire du paiement",
    "BT-61": "Identifiant légal du bénéficiaire du paiement",
    "BT-64": "Adresse du représentant fiscal, ligne 1", "BT-65": "Adresse du représentant fiscal, ligne 2",
    "BT-66": "Ville du représentant fiscal", "BT-67": "Code postal du représentant fiscal",
    "BT-70": "Nom du destinataire de la livraison", "BT-71": "Identifiant du lieu de livraison",
    "BT-75": "Adresse de livraison, ligne 1", "BT-76": "Adresse de livraison, ligne 2",
    "BT-77": "Ville de livraison", "BT-78": "Code postal de livraison",
    "BT-79": "Subdivision du pays de livraison", "BT-85": "Nom du compte de paiement",
    "BT-86": "Identifiant du prestataire de paiement (BIC)", "BT-87": "Numéro de carte (partiel)",
    "BT-88": "Titulaire de la carte", "BT-89": "Référence du mandat de prélèvement",
    "BT-91": "Compte débité", "BT-94": "Pourcentage de la remise sur le document",
    "BT-101": "Pourcentage des frais sur le document", "BT-123": "Description de la pièce justificative",
    "BT-124": "Adresse de la pièce justificative externe", "BT-127": "Note de ligne",
    "BT-128": "Identifiant de l'objet facturé à la ligne", "BT-132": "Ligne de commande référencée",
    "BT-133": "Référence comptable de l'acheteur (ligne)", "BT-138": "Pourcentage de la remise sur la ligne",
    "BT-147": "Rabais sur le prix de l'article", "BT-149": "Quantité de base du prix",
    "BT-154": "Description de l'article", "BT-155": "Référence vendeur de l'article",
    "BT-156": "Référence acheteur de l'article", "BT-159": "Pays d'origine de l'article",
    "BT-163": "Adresse de l'acheteur, ligne 3", "BT-165": "Adresse de livraison, ligne 3",
}


# Données hors fiches, utiles à un exemple (identifiant fiscal du vendeur en franchise en base).
COMPLEMENTS = [
    {"id": "BT-32", "cle": "BT_32", "libelle": "Identifiant fiscal du vendeur", "bloc": "Vendeur", "role": "entete",
     "type": "texte", "profil": "EN16931",
     "xpath": [TRADE + "ram:ApplicableHeaderTradeAgreement/ram:SellerTradeParty/ram:SpecifiedTaxRegistration/ram:ID[@schemeID='FC']"]},
]


def cle(ident: str) -> str:
    if ident in CLES:
        return CLES[ident]
    return re.sub(r"[^0-9A-Za-z]+", "_", ident).strip("_")


def nettoie_xpath(xpath: str) -> str:
    """« …/ram:ID  (@schemeID) » → « …/ram:ID » ; « …@format) » → « …/@format »."""
    xpath = re.sub(r"\s*\(@\w+\)\s*$", "", xpath.strip())
    xpath = re.sub(r"\s*=\s*\"[^\"]*\"\s*$", "", xpath)
    xpath = xpath.rstrip(")")
    return re.sub(r"(?<=[\w\]])@", "/@", xpath)


def type_valeur(ident: str, xpath: str) -> str:
    feuille = xpath.rsplit("/", 1)[-1]
    if "DateTimeString" in feuille:
        return "date"
    if ident in ("BT-146", "BT-148", "BT-147"):
        return "prix"
    if "Quantity" in feuille:
        return "quantite"
    if "Amount" in feuille:
        return "montant"
    if "Percent" in feuille:
        return "taux"
    return "texte"


def role(ident: str, bloc: str) -> str:
    for table, ids in TABLES.items():
        if ident in ids:
            return table
    return BLOC_TABLE.get(bloc, "entete")


def libelle(ident: str, libelle_fiche: str, glossaire: dict) -> str:
    base = re.match(r"^(BT-\d+)$", ident)
    if base:
        return glossaire.get(ident) or LIBELLES.get(ident) or libelle_fiche
    libelle_fiche = re.sub(r"\s*\(ID EXT-FR-FE à confirmer\)", "", libelle_fiche)
    return re.sub(r"\s+", " ", libelle_fiche).strip()


def lire_glossaire() -> dict:
    if not GLOSSAIRE.exists():
        print(f"Attention : glossaire Fchecker absent ({GLOSSAIRE}), libellés maison uniquement.")
        return {}
    return yaml.safe_load(GLOSSAIRE.read_text(encoding="utf-8")).get("glossary", {})


def extraire(spec_dir: Path):
    glossaire = lire_glossaire()
    champs: dict[str, dict] = {}
    fiches: dict[str, dict] = {}
    for classeur in sorted(spec_dir.glob("[A-Z]_*.xlsx")):
        wb = openpyxl.load_workbook(classeur, data_only=True)
        for ws in wb.worksheets:
            entete = {str(ws.cell(r, 1).value).strip(): ws.cell(r, 3).value for r in range(3, 13)
                      if ws.cell(r, 1).value}
            fiche = {
                "classeur": classeur.stem,
                "titre": str(ws["A1"].value).strip(),
                "famille": entete.get("Famille"),
                "descriptif": entete.get("Descriptif"),
                "type_document": str(entete.get("Type de document (BT-3)") or ""),
                "cadre": str(entete.get("Cadre de facturation (BT-23)") or ""),
                "codes": entete.get("Codes"),
                "cas_usage": entete.get("Cas d'usage AFNOR"),
                "profil": entete.get("Profil Factur-X minimal"),
                "attention": entete.get("Points d'attention"),
                "commentaires": {},
            }
            bloc = None
            for r in range(PREMIERE_LIGNE, ws.max_row + 1):
                a, ident, lib, xpath, profil = (ws.cell(r, c).value for c in range(1, 6))
                if not ident:
                    bloc = a or bloc
                    continue
                ident = str(ident).strip()
                bloc = a or bloc
                commentaire = ws.cell(r, 9).value
                if commentaire:
                    fiche["commentaires"][ident] = str(commentaire).strip()
                xpath = nettoie_xpath(str(xpath or ""))
                if ident in champs:  # même ID à deux endroits (BT-149 : prix brut et prix net)
                    if xpath not in champs[ident]["xpath"]:
                        champs[ident]["xpath"].append(xpath)
                    continue
                entree = {
                    "id": ident,
                    "cle": cle(ident),
                    "libelle": libelle(ident, str(lib or ""), glossaire),
                    "bloc": bloc,
                    "role": role(ident, bloc),
                    "type": type_valeur(ident, xpath),
                    # Les données EXT-FR n'existent que dans le profil étendu (la fiche classe EXT-FR-FE-02 en EN16931).
                    "profil": "EXTENDED" if ident.startswith("EXT-") else profil,
                    "xpath": [xpath],
                }
                if ident in FIXES:
                    entree["fixe"] = FIXES[ident]
                champs[ident] = entree
            fiches[ws.title] = fiche
    return list(champs.values()) + COMPLEMENTS, fiches


def main():
    spec_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else SPEC_DIR
    champs, fiches = extraire(spec_dir)
    cles = [c["cle"] for c in champs]
    assert len(cles) == len(set(cles)), "clés en double"
    entete = "# Généré par tools/exemples/extract_catalog.py : ne pas éditer à la main.\n"
    dump = dict(allow_unicode=True, sort_keys=False, width=120)
    (ROOT / "generation" / "champs.yaml").write_text(entete + yaml.safe_dump(champs, **dump), encoding="utf-8")
    (ROOT / "exemples").mkdir(exist_ok=True)
    (ROOT / "exemples" / "fiches.yaml").write_text(entete + yaml.safe_dump(fiches, **dump), encoding="utf-8")
    print(f"{len(champs)} données, {len(fiches)} cas.")


if __name__ == "__main__":
    main()
