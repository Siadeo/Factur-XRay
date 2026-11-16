# Licence

Copyright (c) 2026 Siadeo

Distribué sous licence Apache, version 2.0 (la « Licence ») ; vous ne pouvez utiliser ce fichier que dans le respect des conditions de la Licence. Le texte intégral de la Licence se trouve dans le fichier [LICENSE](LICENSE) et à l'adresse

   https://www.apache.org/licenses/LICENSE-2.0

Sauf disposition légale applicable ou accord écrit, le logiciel distribué sous la Licence est fourni « EN L'ÉTAT », SANS GARANTIE NI CONDITION D'AUCUNE SORTE, expresse ou implicite. Consultez la Licence pour connaître les termes régissant les autorisations et les limitations qui en découlent.

---

## Licences tierces

Ce projet utilise les composants tiers suivants, chacun distribué sous sa propre licence.

Les textes de licence des composants inclus dans l'exécutable sont réunis dans [THIRD_PARTY_LICENSES.txt](THIRD_PARTY_LICENSES.txt), fourni avec le dépôt et avec l'exécutable. Ce fichier est généré à partir de l'exécutable : le régénérer après chaque build.

Les libellés des données affichés par l'application ne reprennent pas la version française de la norme EN 16931 : ils viennent du glossaire de Fchecker (traduit depuis l'anglais, EUPL-1.2) ou d'une traduction propre au projet. Les données des factures d'exemple sont fictives.

| Composant | Licence |
|---|---|
| [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) | MIT |
| [packaging](https://github.com/pypa/packaging) | Apache-2.0 OU BSD-2-Clause |
| [pywin32](https://github.com/mhammond/pywin32) | Licence Python Software Foundation (PSF) |
| [xlwings](https://github.com/xlwings/xlwings) | BSD-3-Clause (hors sous-paquet `xlwings.pro`, exclu de l'exécutable) |
| [lxml](https://github.com/lxml/lxml) | BSD-3-Clause |
| [Jinja2](https://github.com/pallets/jinja) | BSD-3-Clause |
| [PyYAML](https://github.com/yaml/pyyaml) | MIT |
| [factur-x](https://github.com/akretion/factur-x) | BSD-3-Clause |
| [pikepdf](https://github.com/pikepdf/pikepdf) | MPL-2.0 |
| [pypdf](https://github.com/py-pdf/pypdf) | BSD-3-Clause |
| [python-stdnum](https://github.com/arthurdejong/python-stdnum) | LGPL-2.1-or-later |
| [Pillow](https://github.com/python-pillow/Pillow) | MIT-CMU |
| [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl) | MIT |
| [Babel](https://github.com/python-babel/babel) | BSD-3-Clause |
| [Python](https://www.python.org/) et Tcl/Tk | Licence PSF et licences de ses composants |
| Polices Roboto (fournies avec CustomTkinter) | Apache-2.0 |
| Schémas et règles fournis avec factur-x | Selon les fichiers : notices UN/CEFACT, OASIS et W3C, EUPL-1.2 |
| [PyInstaller](https://github.com/pyinstaller/pyinstaller) | GPL-2.0-or-later avec exception pour l'exécutable généré (outil de compilation) |

## Code source des composants sous licence à réciprocité

pikepdf (MPL-2.0) et python-stdnum (LGPL) sont inclus dans l'exécutable sans modification. Leur code source est disponible aux adresses indiquées ci-dessus.

Le code source de Factur-Xray étant public, vous pouvez recompiler l'exécutable avec une version modifiée de ces composants en suivant les instructions de compilation du [readme](readme.md).
