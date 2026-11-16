"""Construit les classeurs de factures d'exemple (``exemples/*.xlsx``) à partir des données fictives.

Usage (développement) : ``python tools/exemples/build.py``

Entrées : ``generation/champs.yaml`` (catalogue), ``exemples/fiches.yaml`` (fiches des cas),
``tools/exemples/cas/commun.yaml`` (parties et valeurs par défaut) et ``tools/exemples/cas/[A-I].yaml``
(un cas par onglet). Sortie : un classeur par famille, un onglet par cas, et ``00_Index.xlsx``.

Mise en page d'un onglet :
- blocs d'en-tête (libellé | valeur), sur deux colonnes de blocs ; chaque valeur porte un nom Excel
  local à l'onglet égal à la clé de la donnée (« BT_31 ») ;
- tableaux Excel des lignes, remises, frais, TVA et notes, dont les en-têtes finissent par l'ID ;
- montants calculés par des formules Excel (lignes, ventilation de TVA, totaux) ;
- à droite de la zone d'impression : les données présentes dans le XML mais pas sur le papier, et la
  fiche du cas.
Seuls les blocs et les colonnes utilisés par le cas sont créés.
"""
from __future__ import annotations

import copy
import re
import textwrap
import sys
from datetime import date, timedelta
from pathlib import Path

import yaml
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from generation.catalog import charger  # noqa: E402
from generation.cii import GUIDELINES, profil  # noqa: E402

CAS_DIR = Path(__file__).with_name("cas")
SORTIE = ROOT / "exemples"

TYPES = {"380": "Facture", "381": "Avoir", "384": "Facture rectificative", "386": "Facture d'acompte",
         "389": "Autofacture", "393": "Facture affacturée", "396": "Avoir affacturé",
         "472": "Facture rectificative affacturée", "471": "Autofacture rectificative",
         "500": "Autofacture d'acompte", "501": "Autofacture affacturée", "503": "Avoir sur acompte",
         "261": "Avoir autofacturé", "262": "Avoir de remise globale", "473": "Autofacture rectificative affacturée",
         "502": "Avoir autofacturé affacturé"}

BLOCS = [
    ("Document", ["BT_1", "BT_2", "BT_3", "BT_23", "BT_5", "BT_9"]),
    ("Vendeur", ["BT_27", "BT_28", "BT_33", "BT_30", "BT_31", "BT_32", "BT_29", "BT_29_0", "BT_29d", "BT_29d_1",
                 "BT_35", "BT_36", "BT_38", "BT_37", "BT_40", "BT_41", "BT_41_0", "BT_42", "BT_43"]),
    ("Acheteur", ["BT_44", "BT_47", "BT_48", "BT_46_0", "BT_50", "BT_51", "BT_163", "BT_53", "BT_52", "BT_55",
                  "BT_56_0", "BT_57", "BT_58"]),
    ("Représentant fiscal du vendeur", ["BT_62", "BT_63", "BT_64", "BT_65", "BT_67", "BT_66", "BT_69"]),
    ("Tiers", ["FACTURANT", "EXT_FR_FE_BG_03", "AGENT_ACHETEUR", "FACTURE_A", "EXT_FR_FE_BG_02"]),
    ("Bénéficiaire du paiement", ["BT_59", "BT_61", "BT_60_0"]),
    ("Références", ["BT_10", "BT_13", "BT_14", "BT_12", "BT_11", "BT_17", "BT_18", "BT_18_1", "BT_122",
                    "BT_123", "BT_124", "BT_19", "BT_15", "BT_16"]),
    ("Facture antérieure", ["BT_25", "BT_26", "EXT_FR_FE_02"]),
    ("Livraison", ["BT_70", "BT_71", "BT_75", "BT_76", "BT_165", "BT_78", "BT_77", "BT_79", "BT_80", "BT_72"]),
    ("Période de facturation", ["BT_73", "BT_74"]),
    ("Paiement", ["BT_81", "BT_82", "BT_84", "BT_85", "BT_86", "BT_91", "BT_87", "BT_88", "BT_89", "BT_20", "BT_83"]),
]
OBLIGATOIRES = {"Document", "Vendeur", "Acheteur"}
TOTAUX = ["BT_106", "BT_107", "BT_108", "BT_109", "BT_110", "BT_111", "BT_112", "BT_113", "BT_114", "BT_115"]
TECHNIQUE = ["BT_24", "BT_34", "BT_49"]  # dans le XML, absents du papier
ORDRE_LIGNES = ["BT_126", "EXT_FR_FE_162", "EXT_FR_FE_163", "BT_155", "BT_153", "BT_154", "BT_127", "EXT_FR_FE_183",
                "BT_129", "BT_130", "BT_148", "BT_147", "BT_146", "BT_149", "BT_151", "BT_152", "EXT_FR_FE_178",
                "EXT_FR_FE_179", "EXT_FR_FE_180", "BT_136", "BT_138", "BT_137", "BT_140", "BT_139", "BT_141",
                "BT_144", "BT_134", "BT_135", "BT_131", "EXT_FR_FE_181", "EXT_FR_FE_184"]
ESSENTIELLES = ["BT_126", "BT_153", "BT_129", "BT_130", "BT_146", "BT_151", "BT_152", "BT_131"]
TITRES = {  # en-têtes de colonnes courts (le libellé complet est dans les Rayons X)
    "BT_126": "N°", "BT_153": "Désignation", "BT_129": "Quantité", "BT_130": "Unité", "BT_146": "Prix net",
    "BT_148": "Prix brut", "BT_147": "Rabais", "BT_151": "Cat. TVA", "BT_152": "Taux TVA", "BT_131": "Montant HT",
    "BT_155": "Réf.", "BT_154": "Description", "BT_127": "Note", "BT_136": "Remise", "BT_141": "Frais",
    "BT_134": "Début", "BT_135": "Fin", "BT_149": "Base prix", "EXT_FR_FE_162": "Parent", "EXT_FR_FE_163": "Sous-type",
    "EXT_FR_FE_181": "TVA ligne", "EXT_FR_FE_184": "TTC ligne", "BT_21": "Code", "BT_22": "Note",
    "BT_118": "Catégorie", "BT_119": "Taux", "BT_116": "Base HT", "BT_117": "Montant TVA", "BT_120": "Motif",
    "BT_121": "Code motif", "BT_8": "Exigibilité",
}

# Mise en forme
GRIS, BLEU = "595959", "1F4E79"
F_TITRE = Font(name="Calibri", size=18, bold=True, color=BLEU)
F_BLOC = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
F_LIB = Font(name="Calibri", size=9, color=GRIS)
F_VAL = Font(name="Calibri", size=10)
F_TOTAL = Font(name="Calibri", size=11, bold=True)
FOND_BLOC = PatternFill("solid", fgColor=BLEU)
FOND_TECH = PatternFill("solid", fgColor="F2F2F2")
TRAIT = Border(bottom=Side(style="thin", color="BFBFBF"))
FORMATS = {"date": "dd/mm/yyyy", "montant": "#,##0.00", "prix": "#,##0.00##", "quantite": "General",
           "taux": "0.00", "texte": "@"}
# Grille : les colonnes B à I portent le tableau des lignes (N°, désignation, quantité, unité, prix, catégorie,
# taux, montant) ; les blocs d'en-tête utilisent des cellules fusionnées (libellé | valeur) sur deux volets.
LARGEURS = {"A": 2, "B": 6, "C": 26, "D": 11, "E": 11, "F": 11, "G": 11, "H": 11, "I": 20}
VOLETS = [(3, 3, 4, 6), (7, 8, 9, 9)]  # (libellé début, libellé fin, valeur début, valeur fin)
DERNIERE_COLONNE_GRILLE = 9


def charger_yaml(chemin: Path):
    return yaml.safe_load(chemin.read_text(encoding="utf-8"))


def fusion(*dicts):
    out = {}
    for d in dicts:
        out.update({k: v for k, v in (d or {}).items()})
    return out


class Cas:
    """Données d'un cas, complétées par les valeurs par défaut de commun.yaml."""

    def __init__(self, onglet: str, donnees: dict, fiche: dict, commun: dict, numero: int):
        self.onglet, self.fiche = onglet, fiche
        d = donnees or {}
        parties = commun["parties"]
        vendeur = parties[d.get("vendeur", "vendeur")]
        acheteur = parties[d.get("acheteur", "acheteur")]
        emission = date.fromisoformat(str(d.get("date", commun["defaut"]["date"])))
        type_doc = str(d.get("type", fiche["type_document"].split()[0]))
        cadre = d.get("cadre") or re.split(r"[/ ]", fiche["cadre"] or "")[0] or None
        entete = {"BT_1": f"F2026{numero:05d}", "BT_2": emission, "BT_3": type_doc, "BT_5": "EUR"}
        if cadre:
            entete["BT_23"] = cadre
        if type_doc not in ("381", "396", "261", "262", "503", "502"):
            entete["BT_9"] = emission + timedelta(days=30)
        # date de livraison ou d'exécution (mention obligatoire), sauf acompte ou période de facturation
        if type_doc not in ("386", "500") and "BT_73" not in (d.get("entete") or {}):
            entete["BT_72"] = emission - timedelta(days=4)
        groupes = [commun[g] for g in d.get("avec", [])]  # ex. avec: [affactureur]
        self.entete = fusion(entete, vendeur, acheteur, commun["paiement"], *groupes, d.get("entete"))
        self.entete = {k: v for k, v in self.entete.items() if v is not None}
        notes = d.get("notes", commun["notes"])
        self.notes = [*notes, *d.get("notes_plus", [])]
        self.lignes = [fusion(commun["ligne"], l) for l in d.get("lignes", [])]
        for i, l in enumerate(self.lignes, 1):
            l.setdefault("BT_126", i)
            if l.get("BT_151") in ("O", "E", "AE", "K", "G", "Z") and "BT_152" not in d.get("lignes", [{}])[i - 1]:
                l["BT_152"] = None if l["BT_151"] == "O" else 0
            if l.get("EXT_FR_FE_163") in ("GROUP", "INFORMATION"):
                for cle in ("BT_151", "BT_152", "BT_130") if l["EXT_FR_FE_163"] == "GROUP" else ("BT_151", "BT_152"):
                    if cle not in d["lignes"][i - 1]:
                        l[cle] = None
        self.sous_factures = str(cadre or "")[-1:] in ("8", "9")  # multi-vendeurs, bidirectionnelle
        if self.sous_factures:
            for l in self.lignes:
                if l.get("EXT_FR_FE_163") == "DETAIL" and l.get("BT_128"):
                    l.setdefault("EXT_FR_FE_178", f"#{l['BT_128']}#")  # BR-FR-MV-08
                if l.get("EXT_FR_FE_163") == "GROUP":
                    l.setdefault("EXT_FR_FE_181", "auto")  # BR-FR-MV-03
                    l.setdefault("EXT_FR_FE_184", "auto")
        self.remises = d.get("remises", [])
        self.charges = d.get("charges", [])
        self.motifs_tva = d.get("tva", {})
        self.exigibilite = d.get("exigibilite")  # BT-8, porté par chaque ventilation de TVA
        self.technique = {"BT_34": vendeur.get("BT_34"), "BT_49": acheteur.get("BT_49")}
        self.technique.update({k: v for k, v in (d.get("entete") or {}).items() if k in TECHNIQUE})
        for k in TECHNIQUE:
            self.entete.pop(k, None)
        # En profil étendu, chaque ligne exonérée reprend le motif de sa ventilation (BR-FXEXT-xx-08)
        if profil(self.contexte_brut()) == "EXTENDED":
            for l in self.lignes:
                motif = self.motifs_tva.get(l.get("BT_151"), {})
                if "BT_120" in motif:
                    l.setdefault("EXT_FR_FE_178", motif["BT_120"])
                if "BT_121" in motif:
                    l.setdefault("EXT_FR_FE_179", motif["BT_121"])

    def contexte_brut(self) -> dict:
        """Données remplies (formules comprises), pour déduire le profil Factur-X."""
        return {"entete": {**self.entete, **self.technique}, "notes": self.notes, "lignes": self.lignes,
                "remises": self.remises, "charges": self.charges, "tva": []}


class Feuille:
    def __init__(self, ws, cas: Cas, catalogue, code: str):
        self.ws, self.cas, self.cat, self.code = ws, cas, catalogue, code
        self.cellules = {}  # clé → adresse absolue ($C$5)
        self.colonnes = {}  # (rôle, clé) → (lettre, première ligne, dernière ligne)
        self.derniere_colonne = 6

    # --- primitives ------------------------------------------------------------------------
    def nommer(self, cle, cellule):
        adresse = f"${cellule.column_letter}${cellule.row}"
        self.cellules[cle] = adresse
        self.ws.defined_names[cle] = DefinedName(cle, attr_text=f"'{self.ws.title}'!{adresse}")

    def ecrire(self, cellule, valeur, type_valeur="texte"):
        if isinstance(valeur, str) and valeur.startswith("="):
            valeur = self.formule(valeur)
        elif type_valeur == "date" and isinstance(valeur, str):
            valeur = date.fromisoformat(valeur)
        elif type_valeur == "texte" and isinstance(valeur, (int, float)) and not isinstance(valeur, bool):
            valeur = str(valeur)  # codes et identifiants : texte (zéros de tête, « 380 »)
        cellule.value = valeur
        cellule.number_format = FORMATS[type_valeur]

    def formule(self, texte):
        """« ={BT_112} » → « =$C$40 » (références aux cellules nommées de l'onglet)."""
        return re.sub(r"\{(\w+)\}", lambda m: self.cellules[m.group(1)], texte)

    def bandeau(self, ligne, colonne, texte, largeur=2):
        """Bandeau de titre de bloc sur ``largeur`` colonnes."""
        for c in range(colonne, colonne + largeur):
            self.ws.cell(ligne, c).fill = FOND_BLOC
        cell = self.ws.cell(ligne, colonne, texte)
        cell.font = F_BLOC

    def paire(self, ligne, volet, cle, valeur, police=F_VAL, fond=None):
        """Libellé | valeur sur un volet (l1, l2, v1, v2) ; la cellule de la valeur porte le nom de la donnée."""
        l1, l2, v1, v2 = volet
        champ = self.cat.par_cle[cle]
        lib = self.ws.cell(ligne, l1, champ.libelle)
        lib.font, lib.alignment = F_LIB, Alignment(vertical="top", wrap_text=True)
        val = self.ws.cell(ligne, v1)
        for c1, c2 in ((l1, l2), (v1, v2)):
            if c2 > c1:
                self.ws.merge_cells(start_row=ligne, start_column=c1, end_row=ligne, end_column=c2)
        for c in range(v1, v2 + 1):
            self.ws.cell(ligne, c).border = TRAIT
        self.ecrire(val, valeur, champ.type)
        val.font, val.border = police, TRAIT
        val.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        if fond:
            lib.fill = val.fill = fond
        self.nommer(cle, val)

    # --- blocs d'en-tête ---------------------------------------------------------------------
    def blocs(self, ligne):
        visibles = []
        for titre, cles in BLOCS:
            presentes = [c for c in cles if c in self.cas.entete]
            if presentes or titre in OBLIGATOIRES:
                visibles.append((titre, presentes))
        for i in range(0, len(visibles), 2):
            paire_blocs = visibles[i:i + 2]
            hauteur = max(len(c) for _, c in paire_blocs) + 1
            for (titre, cles), volet in zip(paire_blocs, VOLETS):
                self.bandeau(ligne, volet[0], titre, volet[3] - volet[0] + 1)
                for j, cle in enumerate(cles, 1):
                    self.paire(ligne + j, volet, cle, self.cas.entete[cle])
            ligne += hauteur + 1
        return ligne

    # --- tableaux ----------------------------------------------------------------------------
    def tableau(self, ligne, role, cles, lignes, titre, formules=None, colonne=2):
        """Tableau Excel ``T<code>_<role>`` ; renvoie la ligne qui suit."""
        self.bandeau(ligne, colonne, titre, max(len(cles), 2))
        ligne += 1
        for j, cle in enumerate(cles):
            champ = self.cat.par_cle[cle]
            cell = self.ws.cell(ligne, colonne + j, self.cat.entete(champ, TITRES.get(cle)))
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            cell.font = Font(name="Calibri", size=9, bold=True)
            self.colonnes[(role, cle)] = (get_column_letter(colonne + j), ligne + 1, ligne + len(lignes))
        self.ws.row_dimensions[ligne].height = 30
        for i, donnees in enumerate(lignes, 1):
            for j, cle in enumerate(cles):
                cell = self.ws.cell(ligne + i, colonne + j)
                valeur = donnees.get(cle)
                if formules and cle in formules and (cle not in donnees or donnees[cle] == "auto"):
                    valeur = formules[cle](ligne + i, donnees)
                self.ecrire(cell, valeur, self.cat.par_cle[cle].type)
                cell.font = F_VAL
                cell.alignment = Alignment(vertical="top", wrap_text=self.cat.par_cle[cle].type == "texte")
        fin = ligne + max(len(lignes), 1)
        ref = f"{get_column_letter(colonne)}{ligne}:{get_column_letter(colonne - 1 + len(cles))}{fin}"
        table = Table(displayName=f"T{self.code}_{role}", ref=ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
        self.ws.add_table(table)
        self.derniere_colonne = max(self.derniere_colonne, colonne - 1 + len(cles))
        return fin + 2

    def plage(self, role, cle):
        if (role, cle) not in self.colonnes:
            return None
        lettre, debut, fin = self.colonnes[(role, cle)]
        return f"${lettre}${debut}:${lettre}${fin}"

    def ref(self, role, cle, ligne):
        lettre = self.colonnes[(role, cle)][0]
        return f"{lettre}{ligne}"

    def lignes_facture(self, ligne):
        cas = self.cas
        utilisees = {k for l in cas.lignes for k, v in l.items() if v is not None}
        cles = ESSENTIELLES + [k for k in ORDRE_LIGNES if k in utilisees and k not in ESSENTIELLES]
        cles += [c.cle for c in self.cat.du_role("lignes") if c.cle in utilisees and c.cle not in cles]
        role = "lignes"
        # les colonnes doivent être connues avant d'écrire les formules
        for j, cle in enumerate(cles):
            self.colonnes[(role, cle)] = (get_column_letter(2 + j), ligne + 2, ligne + 1 + len(cas.lignes))
        r = lambda cle, n: self.ref(role, cle, n) if (role, cle) in self.colonnes else None  # noqa: E731

        def montant(n, l):
            if l.get("EXT_FR_FE_163") == "GROUP":
                return f"=SUMIFS({self.plage(role, 'BT_131')},{self.plage(role, 'EXT_FR_FE_162')},{r('BT_126', n)})"
            if l.get("EXT_FR_FE_163") == "INFORMATION":
                return 0
            base = f"IF(N({r('BT_149', n)})=0,1,{r('BT_149', n)})" if r("BT_149", n) else "1"
            calc = f"N({r('BT_129', n)})*N({r('BT_146', n)})/{base}"
            if r("BT_136", n):
                calc += f"-N({r('BT_136', n)})"
            if r("BT_141", n):
                calc += f"+N({r('BT_141', n)})"
            return f"=ROUND({calc},2)"

        def prix_net(n, l):
            if l.get("BT_148") is not None:
                return f"={r('BT_148', n)}-N({r('BT_147', n)})" if r("BT_147", n) else f"={r('BT_148', n)}"
            return None if l.get("BT_146") in (None, "auto") else l["BT_146"]

        def tva_ligne(n, l):
            if l.get("EXT_FR_FE_163") == "GROUP":
                p, m, t = (self.plage(role, k) for k in ("EXT_FR_FE_162", "BT_131", "BT_152"))
                return f"=ROUND(SUMPRODUCT(({p}={r('BT_126', n)})*{m}*{t})/100,2)"
            return f"=ROUND({r('BT_131', n)}*N({r('BT_152', n)})/100,2)"

        def ttc_ligne(n, l):
            return f"={r('BT_131', n)}+{r('EXT_FR_FE_181', n)}"

        formules = {"BT_131": montant, "BT_146": prix_net, "EXT_FR_FE_181": tva_ligne, "EXT_FR_FE_184": ttc_ligne}
        for l in cas.lignes:  # colonnes calculées : remplies même si le cas ne les donne pas
            for cle in formules:
                # TVA et TTC de ligne (EXT-FR-FE-181 / 184) : exigés sur les lignes GROUP seulement
                if cle.startswith("EXT_") and l.get("EXT_FR_FE_163") != "GROUP":
                    continue
                if (role, cle) in self.colonnes:
                    l.setdefault(cle, "auto")
        return self.tableau(ligne, role, cles, cas.lignes, "Lignes de la facture", formules)

    def remises_charges(self, ligne):
        for role, titre, donnees in (("remises", "Remises sur le document", self.cas.remises),
                                     ("charges", "Frais sur le document", self.cas.charges)):
            if donnees:
                utilisees = {k for d in donnees for k in d}
                cles = [c.cle for c in self.cat.du_role(role) if c.cle in utilisees]
                ligne = self.tableau(ligne, role, cles, donnees, titre, colonne=3)
        return ligne

    def ventilation_tva(self, ligne):
        cas, role = self.cas, "tva"
        groupes = []
        for source in (cas.lignes, [{"BT_151": d.get("BT_95"), "BT_152": d.get("BT_96")} for d in cas.remises],
                       [{"BT_151": d.get("BT_102"), "BT_152": d.get("BT_103")} for d in cas.charges]):
            for l in source:
                # multi-vendeurs : une ventilation par sous-facture, repérée par « #n°# » (BR-FR-MV-09)
                cle = (l.get("BT_151"), l.get("BT_152"), l.get("EXT_FR_FE_178") if cas.sous_factures else None)
                if cle[0] and cle not in groupes:
                    groupes.append(cle)
        donnees = []
        for categorie, taux, sous_facture in groupes:
            d = {"BT_118": categorie, "BT_119": taux, **cas.motifs_tva.get(categorie, {})}
            if sous_facture:
                d["BT_120"] = sous_facture
            if cas.exigibilite:
                d["BT_8"] = cas.exigibilite
            donnees.append(d)
        cles = ["BT_118", "BT_119", "BT_116", "BT_117"]
        cles += [k for k in ("BT_120", "BT_121", "BT_8") if any(k in d for d in donnees)]
        for j, cle in enumerate(cles):
            self.colonnes[(role, cle)] = (get_column_letter(3 + j), ligne + 2, ligne + 1 + len(donnees))

        def critere(n, cle):
            return self.ref(role, cle, n)

        def base(n, d):
            termes = []
            for src, c_m, c_cat, c_taux in (("lignes", "BT_131", "BT_151", "BT_152"), ("remises", "BT_92", "BT_95", "BT_96"),
                                            ("charges", "BT_99", "BT_102", "BT_103")):
                if self.plage(src, c_m) is None:
                    continue
                taux = f'IF({critere(n, "BT_119")}="","",{critere(n, "BT_119")})'
                somme = (f"SUMIFS({self.plage(src, c_m)},{self.plage(src, c_cat)},{critere(n, 'BT_118')},"
                         f"{self.plage(src, c_taux)},{taux}")
                if src == "lignes" and cas.sous_factures:
                    somme += f",{self.plage(src, 'EXT_FR_FE_178')},{critere(n, 'BT_120')}"
                somme += ")"
                termes.append(("-" if src == "remises" else "+") + somme)
            return "=ROUND(" + "".join(termes).lstrip("+") + ",2)"

        def montant(n, d):
            return f"=ROUND({critere(n, 'BT_116')}*N({critere(n, 'BT_119')})/100,2)"

        for d in donnees:
            d["BT_116"], d["BT_117"] = "auto", d.get("BT_117", "auto")
        return self.tableau(ligne, role, cles, donnees, "Récapitulatif de TVA", {"BT_116": base, "BT_117": montant},
                            colonne=3)

    def totaux(self, ligne):
        cas, e = self.cas, self.cas.entete
        lignes131 = self.plage("lignes", "BT_131")
        sous_type = self.plage("lignes", "EXT_FR_FE_163")
        formules = {
            "BT_106": f'=SUMIFS({lignes131},{sous_type},"<>DETAIL")' if sous_type else f"=SUM({lignes131})",
            "BT_107": f"=SUM({self.plage('remises', 'BT_92')})" if cas.remises else None,
            "BT_108": f"=SUM({self.plage('charges', 'BT_99')})" if cas.charges else None,
            "BT_110": f"=SUM({self.plage('tva', 'BT_117')})",
            "BT_113": 0,
        }
        valeurs = {k: v for k, v in formules.items() if v is not None}
        valeurs.update({k: e.pop(k) for k in TOTAUX if k in e})
        hors_taxes = "={BT_106}" + ("-{BT_107}" if "BT_107" in valeurs else "") + ("+{BT_108}" if "BT_108" in valeurs else "")
        valeurs.setdefault("BT_109", hors_taxes)
        valeurs.setdefault("BT_112", "={BT_109}+{BT_110}")
        valeurs.setdefault("BT_115", "={BT_112}-{BT_113}" + ("+{BT_114}" if "BT_114" in valeurs else ""))
        volet = VOLETS[1]
        self.bandeau(ligne, volet[0], "Totaux", volet[3] - volet[0] + 1)
        for j, cle in enumerate(k for k in TOTAUX if k in valeurs):
            gras = cle in ("BT_112", "BT_115")
            self.paire(ligne + 1 + j, volet, cle, valeurs[cle], F_TOTAL if gras else F_VAL)
        return ligne + len(valeurs) + 2

    def notes(self, ligne):
        if not self.cas.notes:
            return ligne
        fin = self.tableau(ligne, "notes", ["BT_21", "BT_22"], self.cas.notes, "Notes de la facture")
        for n in range(ligne + 2, fin - 1):
            texte = str(self.ws.cell(n, 3).value or "")
            self.ws.row_dimensions[n].height = 13 * max(1, len(texte) // 26 + 1)
        return fin

    def zone_technique(self):
        """À droite de la zone d'impression : données du XML absentes du papier, puis fiche du cas.
        Les textes longs sont découpés en lignes courtes pour ne pas agrandir les lignes de la facture."""
        col = self.derniere_colonne + 2
        self.ws.column_dimensions[get_column_letter(col)].width = 26
        self.ws.column_dimensions[get_column_letter(col + 1)].width = 75
        volet = (col, col, col + 1, col + 1)
        self.bandeau(1, col, "Dans le XML, pas sur le papier")
        niveau = profil(self.cas.contexte_brut(), self.cat)
        technique = {"BT_24": GUIDELINES[niveau], **{k: v for k, v in self.cas.technique.items() if v}}
        for j, (cle, valeur) in enumerate(technique.items(), 2):
            self.paire(j, volet, cle, valeur, fond=FOND_TECH)
        ligne = len(technique) + 3
        self.bandeau(ligne, col, "Fiche du cas")
        fiche = self.cas.fiche
        for lib, texte in [("Cas", fiche["titre"]), ("Famille", fiche["famille"]), ("Codes", fiche["codes"]),
                           ("Profil Factur-X", niveau), ("Descriptif", fiche["descriptif"]),
                           ("Points d'attention", fiche["attention"])]:
            morceaux = [m for p in str(texte or "").splitlines() for m in (textwrap.wrap(p, 95) or [""])]
            for k, morceau in enumerate(morceaux):
                ligne += 1
                a = self.ws.cell(ligne, col, lib if k == 0 else None)
                b = self.ws.cell(ligne, col + 1, morceau)
                a.font, b.font = F_LIB, F_VAL
                a.fill = b.fill = FOND_TECH

    def construire(self):
        ws = self.ws
        for lettre, largeur in LARGEURS.items():
            ws.column_dimensions[lettre].width = largeur
        titre = TYPES.get(str(self.cas.entete["BT_3"]), "Facture").upper()
        ws.cell(2, 2, titre).font = F_TITRE
        ws.cell(3, 2, self.cas.fiche["titre"]).font = Font(name="Calibri", size=9, italic=True, color=GRIS)
        ligne = self.blocs(5)
        ligne = self.lignes_facture(ligne)
        ligne = self.remises_charges(ligne)
        ligne = self.ventilation_tva(ligne)
        ligne = self.totaux(ligne)
        ligne = self.notes(ligne)
        for lettre in (get_column_letter(c) for c in range(DERNIERE_COLONNE_GRILLE + 1, self.derniere_colonne + 1)):
            ws.column_dimensions[lettre].width = 14
        self.derniere_colonne = max(self.derniere_colonne, DERNIERE_COLONNE_GRILLE)
        self.zone_technique()
        ws.print_area = f"B1:{get_column_letter(self.derniere_colonne)}{ligne}"
        ws.page_setup.orientation = "landscape" if self.derniere_colonne > 12 else "portrait"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_options.horizontalCentered = True
        ws.page_margins.left = ws.page_margins.right = 0.4
        ws.sheet_view.showGridLines = False
        ws.sheet_view.zoomScale = 90


def code_onglet(onglet: str) -> str:
    return re.match(r"^(\d+[a-z]?)", onglet).group(1)


def construire(sortie: Path = SORTIE):
    catalogue = charger()
    fiches = charger_yaml(SORTIE / "fiches.yaml")
    commun = charger_yaml(CAS_DIR / "commun.yaml")
    sortie.mkdir(parents=True, exist_ok=True)
    index = []
    for fichier in sorted(CAS_DIR.glob("[A-I].yaml")):
        famille = charger_yaml(fichier)
        classeur = famille["classeur"]
        wb = Workbook()
        wb.remove(wb.active)
        for onglet, donnees in famille["cas"].items():
            fiche = fiches[donnees.get("fiche", onglet) if donnees else onglet]
            numero = int(re.match(r"\d+", onglet).group())
            cas = Cas(onglet, copy.deepcopy(donnees), fiche, commun, numero * 10 + (1 if onglet[2:3] == "b" else 0))
            Feuille(wb.create_sheet(onglet), cas, catalogue, code_onglet(onglet)).construire()
            index.append((classeur, onglet, fiche))
        wb.calculation.fullCalcOnLoad = True
        wb.save(sortie / f"{classeur}.xlsx")
        print(f"{classeur}.xlsx : {len(famille['cas'])} onglets")
    construire_index(index, sortie)


def construire_index(index, sortie: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Index"
    ws["B2"] = "Factures d'exemple Factur-X : index des cas"
    ws["B2"].font = F_TITRE
    ws["B3"] = "Cliquer sur un cas pour ouvrir son onglet. Données entièrement fictives."
    ws["B3"].font = F_LIB
    entetes = ["N°", "Cas", "Codes", "Profil", "Famille", "Descriptif"]
    for j, t in enumerate(entetes):
        c = ws.cell(5, 2 + j, t)
        c.font, c.fill = F_BLOC, FOND_BLOC
    for i, (classeur, onglet, fiche) in enumerate(index, 6):
        valeurs = [code_onglet(onglet), fiche["titre"], fiche["codes"], fiche["profil"], fiche["famille"],
                   fiche["descriptif"]]
        for j, v in enumerate(valeurs):
            c = ws.cell(i, 2 + j, v)
            c.font = F_VAL
            c.alignment = Alignment(vertical="top", wrap_text=j == 5)
        lien = ws.cell(i, 3)
        lien.hyperlink = f"{classeur}.xlsx#'{onglet}'!A1"
        lien.font = Font(name="Calibri", size=10, color="0563C1", underline="single")
    for lettre, largeur in zip("ABCDEFG", (2, 6, 52, 16, 12, 30, 90)):
        ws.column_dimensions[lettre].width = largeur
    ws.sheet_view.showGridLines = False
    wb.save(sortie / "00_Index.xlsx")


if __name__ == "__main__":
    construire()
