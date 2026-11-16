# Factur-Xray (EzFacture)

Application Windows qui permet de générer des factures de test à partir de 50 factures d'exemple
de la réforme française (Flux 2). Chaque exemple est un onglet Excel : on peut l'explorer aux **Rayons X** (le BT
EN 16931 et la balise XML derrière chaque cellule), puis on **génère le Factur-X** correspondant (PDF/A-3 avec
le XML CII embarqué, profil EN16931 ou EXTENDED).

Toutes les données des exemples sont **fictives**.

## Les exemples

`exemples/` contient un index et 9 classeurs, un onglet par cas :

| Classeur | Cas |
|---|---|
| A. Factures de base | 01 à 07 : biens, services (encaissement, débits), mixte, déjà payée, multi-commandes, abonnement |
| B. Acomptes | 08 à 11 : acompte, acompte déjà payé, facture définitive, avoir sur acompte |
| C. Corrections | 12 à 17 : rectificative, avoirs total, partiel, de remise globale, net de taxe, note de débit |
| D. Affacturage | 18 à 21 |
| E. Mandat et autofacturation | 22 à 30 (le cas 30, compensation, comporte deux documents : 30a et 30b) |
| F. Tiers payeurs et intermédiaires | 31 à 36 |
| G. BTP, sous-traitance, cotraitance | 37 à 41 |
| H. TVA et régimes particuliers | 42 à 48 |
| I. Structure avancée | 49 sous-lignes, 50 multi-vendeurs |

Dans chaque onglet :

- la facture (zone imprimée) : blocs d'en-tête, tableau des lignes, récapitulatif de TVA, totaux et notes. Les
  montants sont calculés par des formules Excel ;
- à droite, hors impression : les données présentes dans le XML mais pas sur le papier (identifiant de
  spécification BT-24, adresses électroniques BT-34 / BT-49) et la **fiche du cas** (descriptif, codes, points
  d'attention).

**Ce qu'on voit est ce qui part dans le XML** : chaque cellule de donnée porte le nom de sa donnée (« BT_31 »),
chaque colonne de tableau finit par son identifiant (« Quantité (BT-129) »), et le XML reprend les valeurs
affichées, sans recalcul.

## Utilisation

1. **Ouvrir** : choisir l'index ou une famille. Le classeur s'ouvre en lecture seule : on peut modifier des
   valeurs pour expérimenter, les exemples d'origine ne sont jamais écrasés.
2. Choisir un cas avec les onglets Excel.
3. **Rayons X** : un panneau suit la cellule active et affiche son BT, son libellé, l'aperçu de la balise CII,
   le chemin XML et le commentaire propre au cas. L'interrupteur colore les données par bloc (une couleur par
   bloc, légende dans le panneau) et pose une note au survol de chaque cellule.
4. **Générer la facture** : produit `factures/<classeur>/<onglet>.pdf` (Factur-X) et `.xml`, puis ouvre le
   PDF. Aucune validation n'est faite : pour contrôler une facture, l'ouvrir dans Fchecker.

## Prérequis

- **Windows** et **Microsoft Excel** desktop (pilotage par COM/xlwings).
- **Python 3.13** pour travailler à partir des sources.

## Développeurs

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
python main.py
```

Tests (sans Excel : catalogue, rendu et validation XSD des 51 factures à partir de leurs instantanés) :

```powershell
pytest generation/tests xray/tests
```

### Régénérer les exemples

Les exemples sont construits à partir des fiches de spécification et de données fictives :

```powershell
python tools/exemples/extract_catalog.py "D:\__E-invoicing\ez facture new"   # fiches → generation/champs.yaml + exemples/fiches.yaml
python tools/exemples/build.py                                               # cas/*.yaml → exemples/*.xlsx
python tools/exemples/instantanes.py [dossier_xml]                           # Excel requis : exemples/instantanes/*.json (+ XML)
```

- `tools/exemples/cas/commun.yaml` : parties fictives (SIREN, TVA et IBAN aux clés valides) et valeurs par défaut
  (notes PMD, PMT, AAB, BAR ; coordonnées de paiement) ;
- `tools/exemples/cas/[A-I].yaml` : un cas par onglet, limité à ce qui le distingue ;
- `generation/templates/cii.xml` : le gabarit du XML CII, dans l'ordre du schéma.

Les 51 XML passent les contrôles de Fchecker (XSD, schematrons EN 16931, BR-FR, Factur-X EXTENDED), sauf le
cas 28 (achat à un particulier sans SIREN), dont le codage reste à confirmer dans la fiche.

## Compilation en exécutable

```powershell
python build.py
```

Le script lance PyInstaller ([main.spec](main.spec)) puis copie `exemples/`, `assets/` et les licences dans `dist/`, et crée `factures/`.

## Licence

Distribué sous licence Apache, version 2.0 : voir [LICENSE](LICENSE). Composants tiers : [LICENSE.md](LICENSE.md).
