"""Contrôleur : ouvrir un classeur d'exemples, générer le Factur-X de l'onglet actif, Rayons X."""
import logging
import os
import webbrowser

import pywintypes

from . import workbook
from generation.cii import generer_xml
from generation.pdf import facturx
from generation.reader import lire
from tools.utils import close_all_protected_sheets, disable_win_auto_printer_manage, use_pdf_printer

logger = logging.getLogger(__name__)
AIDE = "https://github.com/Siadeo/Factur-XRay#readme"


class Controller:

    def __init__(self, view):
        self.view = view
        self.xray_panel = None  # panneau « Rayons X » (xray/panel.py), ouvert à la demande
        self.menu = workbook.classeurs()
        self.view.set_actions({"generer": self.generer, "aide": self.aide, "xray": self.open_xray})
        self.view.set_menu_ouvrir(self.ouvrir, list(self.menu))
        self.view.on_quit_confirmed = self.fermer
        close_all_protected_sheets()
        # « Laisser Windows gérer mon imprimante par défaut » empêche Excel d'exporter le PDF
        disable_win_auto_printer_manage()
        if not self.menu:
            self.view.show_feedback(f"Aucun classeur d'exemples dans {workbook.EXEMPLES}", "error")

    def ouvrir(self, choix):
        self.view.reset_menu_ouvrir()
        try:
            workbook.ouvrir(self.menu[choix])
        except Exception as e:
            logger.exception(e)
            self.view.show_feedback(f"Impossible d'ouvrir {choix} : {e}", "error")

    def generer(self):
        """Génère le Factur-X (PDF + XML) de l'onglet actif, sans validation, puis ouvre le PDF."""
        self._xray_off()  # couleurs et notes des Rayons X : jamais dans le PDF
        try:
            sheet = workbook.onglet_actif()
            self.view.show_feedback(f"Génération de « {sheet.name} »…")
            xml, profil = generer_xml(lire(sheet))
            dossier = workbook.FACTURES / sheet.book.name.rsplit(".", 1)[0]
            pdf = dossier / f"{sheet.name}.pdf"
            use_pdf_printer()
            workbook.exporter_pdf(sheet, pdf)
            facturx(pdf, xml, profil)
            pdf.with_suffix(".xml").write_bytes(xml)
        except ValueError as e:
            self.view.show_feedback(str(e), "error")
            return
        except (pywintypes.com_error, OSError) as e:
            logger.exception(e)
            self.view.show_feedback("Erreur lors de la création du PDF : si la facture est ouverte dans un "
                                    f"lecteur PDF, fermez-la puis réessayez. ({e})", "error")
            return
        except Exception as e:
            logger.exception(e)
            self.view.show_feedback(f"Erreur lors de la génération : {e}", "error")
            return
        self.view.show_feedback(f"Factur-X {profil} généré :\n{pdf}\n{pdf.with_suffix('.xml').name}", "success")
        try:
            os.startfile(pdf)
        except OSError as e:
            logger.error(f"Impossible d'ouvrir le PDF : {e}")

    def aide(self):
        webbrowser.open(AIDE)

    def fermer(self):
        self._xray_close()
        workbook.fermer_exemples()

    # ------------------------------------------------------------------ Rayons X
    def open_xray(self):
        """Ouvre le panneau « Rayons X » : BT, XML et commentaires du cas pour la cellule active."""
        if self.xray_panel is not None and self.xray_panel.winfo_exists():
            self.xray_panel.focus()
            return
        try:
            from xray import charger_paquet
            from xray.panel import XRayPanel
            self.xray_panel = XRayPanel(self.view, charger_paquet(), workbook.excel(), on_close=self._xray_forget)
        except Exception as e:
            logger.exception(e)
            self.view.show_feedback(f"Rayons X indisponible : {e}", "error")

    def _xray_forget(self):
        self.xray_panel = None

    def _xray_off(self):
        """Retire couleurs et notes des Rayons X avant d'exporter le document."""
        if self.xray_panel is not None and self.xray_panel.winfo_exists():
            self.xray_panel.suspend()

    def _xray_close(self):
        if self.xray_panel is not None and self.xray_panel.winfo_exists():
            self.xray_panel.close()
        self.xray_panel = None
