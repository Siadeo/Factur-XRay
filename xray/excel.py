"""Rayons X dans Excel (xlwings) : repérer la donnée de la cellule active, annoter, restaurer.

Une donnée d'en-tête est une cellule nommée (« BT_31 ») ; une donnée répétable est une colonne d'un
tableau dont l'en-tête se termine par l'ID (« Quantité (BT-129) »). Le mode couleurs et notes modifie
l'onglet : ``XRayMode.off`` remet tout en l'état, avant tout export PDF.
"""
from __future__ import annotations

import logging

from .pack import Cellule, Paquet, note_text, texte_sur

logger = logging.getLogger(__name__)


def _bornes(api_range):
    """(ligne1, colonne1, ligne2, colonne2) d'une plage COM, zone fusionnée comprise."""
    zone = api_range.MergeArea if api_range.Cells.Count == 1 else api_range
    return zone.Row, zone.Column, zone.Row + zone.Rows.Count - 1, zone.Column + zone.Columns.Count - 1


def zones(paquet: Paquet, sheet):
    """[(bornes, Cellule, plage COM de la valeur ou de la colonne, cellule COM qui porte la note)]."""
    catalogue, resultat = paquet.catalogue, []
    for nom in sheet.names:
        champ = catalogue.par_cle.get(nom.name.split("!")[-1])
        if champ is None:
            continue
        try:
            rng = nom.refers_to_range.api
            resultat.append((_bornes(rng), paquet.cellule(champ), rng, rng))
        except Exception as e:  # nom cassé : ignoré
            logger.debug("Rayons X : %s ignoré (%s)", nom.name, e)
    for tableau in sheet.tables:
        corps = tableau.api.DataBodyRange
        for j, entete in enumerate(tableau.header_row_range.value or []):
            champ = catalogue.depuis_entete(entete)
            if champ is None:
                continue
            colonne = tableau.api.ListColumns(j + 1)
            resultat.append((_bornes(colonne.Range), paquet.cellule(champ),
                             corps.Columns(j + 1) if corps is not None else colonne.Range, colonne.Range.Cells(1, 1)))
    return resultat


class Locator:
    """Associe la cellule active d'Excel à sa donnée (zones recalculées quand l'onglet change)."""

    def __init__(self, paquet: Paquet, app):
        self.paquet, self.app = paquet, app
        self._onglet, self._zones = None, []

    def active(self):
        """(Cellule, valeur, nom de l'onglet) de la cellule active ; Cellule = None hors référentiel."""
        sel = self.app.selection
        if sel is None:
            return None, None, None
        sheet = sel.sheet
        cle = (sheet.book.name, sheet.name)
        if cle != self._onglet:
            self._onglet, self._zones = cle, zones(self.paquet, sheet)
        for (r1, c1, r2, c2), cellule, _, _ in self._zones:
            if r1 <= sel.row <= r2 and c1 <= sel.column <= c2:
                return cellule, sheet.cells(sel.row, sel.column).value, sheet.name
        return None, None, sheet.name


class XRayMode:
    """Couleur par bloc et note au survol sur chaque donnée de l'onglet actif."""

    def __init__(self, paquet: Paquet, app):
        self.paquet, self.app = paquet, app
        # (cellule de la note, plage colorée, couleur d'origine, note d'origine, couleur du texte d'origine)
        self.saved = []
        self.blocs = []  # blocs colorés dans l'onglet, pour la légende
        self.active = False

    def on(self):
        if self.active:
            return
        try:
            self._appliquer(self.app.books.active.sheets.active)
        except Exception:  # Excel occupé en cours de route : on défait ce qui a été posé
            self.active = True
            self.off()
            raise
        self.active = True

    def _appliquer(self, sheet):
        for _, cellule, plage, porte_note in zones(self.paquet, sheet):
            commentaire = porte_note.Comment
            note = commentaire.Text() if commentaire is not None else None
            couleur = plage.Interior.Color if plage.Interior.ColorIndex != -4142 else None
            police = plage.Font.Color  # None si la plage mélange plusieurs couleurs : laissée telle quelle
            self.saved.append((porte_note, plage, couleur, note, police))
            if commentaire is not None:
                porte_note.ClearComments()
            texte = note_text(self.paquet, cellule, sheet.name)
            bulle = porte_note.AddComment(texte)
            bulle.Shape.Width = 340
            bulle.Shape.Height = 15 * (texte.count("\n") + 2)
            fond = self.paquet.couleur(cellule.champ)
            plage.Interior.Color = _bgr(fond)
            if police is not None:
                plage.Font.Color = _bgr(texte_sur(fond))
            self.blocs.append(cellule.champ.bloc)

    def off(self):
        """Restaure couleurs et notes d'origine (sans effet si le mode est inactif)."""
        if not self.active:
            return
        for porte_note, plage, couleur, note, police in reversed(self.saved):
            try:
                porte_note.ClearComments()
                if note is not None:
                    porte_note.AddComment(note)
                if couleur is None:
                    plage.Interior.ColorIndex = -4142  # xlColorIndexNone
                else:
                    plage.Interior.Color = couleur
                if police is not None:
                    plage.Font.Color = police
            except Exception as e:  # classeur fermé entre-temps
                logger.debug("Rayons X : restauration impossible (%s)", e)
        self.saved = []
        self.blocs = []
        self.active = False


def _bgr(hexa: str) -> int:
    """« #RRGGBB » → entier de couleur COM d'Excel (ordre BGR)."""
    r, g, b = (int(hexa[i:i + 2], 16) for i in (1, 3, 5))
    return r + g * 256 + b * 65536


__all__ = ["Cellule", "Locator", "XRayMode", "zones"]
