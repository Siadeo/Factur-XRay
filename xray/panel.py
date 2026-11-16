"""Panneau « Rayons X » : le BT et le XML de la cellule active d'Excel, mis à jour en direct."""
from __future__ import annotations

import logging

import customtkinter

from .excel import Locator, XRayMode
from .fragment import fragment
from .pack import Paquet, short_xpath, texte_sur

logger = logging.getLogger(__name__)
POLL_MS = 600
MONO = ("Consolas", 12)


class XRayPanel(customtkinter.CTkToplevel):
    def __init__(self, parent, paquet: Paquet, app, on_close=None):
        super().__init__(parent)
        self.paquet, self.on_close = paquet, on_close
        self.locator = Locator(paquet, app)
        self.mode = XRayMode(paquet, app)
        self.current = None
        self._job = None

        self.title("Rayons X — le BT derrière la cellule")
        self.geometry(self._placement(parent, 560, 680))
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.attributes("-topmost", True)  # reste devant Excel pendant qu'on clique dans les cellules
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=3)
        self.grid_rowconfigure(4, weight=2)

        self.switch = customtkinter.CTkSwitch(self, text="Colorer les données par bloc et ajouter une note au survol",
                                              command=self._toggle)
        self.switch.grid(row=0, column=0, padx=15, pady=(12, 4), sticky="w")
        self.heading = customtkinter.CTkLabel(self, text="Sélectionnez une cellule de la facture",
                                              font=("", 15, "bold"), anchor="w", justify="left", wraplength=520)
        self.heading.grid(row=1, column=0, padx=15, pady=(6, 0), sticky="ew")
        self.meta = customtkinter.CTkLabel(self, text="", anchor="w", justify="left", wraplength=520,
                                           text_color="#bdc3c7")
        self.meta.grid(row=2, column=0, padx=15, pady=(0, 6), sticky="ew")
        self.xml = customtkinter.CTkTextbox(self, font=MONO, wrap="none", height=200)
        self.xml.grid(row=3, column=0, padx=15, pady=4, sticky="nsew")
        self.details = customtkinter.CTkTextbox(self, wrap="word", height=150)
        self.details.grid(row=4, column=0, padx=15, pady=(4, 6), sticky="nsew")
        self.legend = customtkinter.CTkFrame(self, fg_color="transparent")  # affichée avec les couleurs
        self.legend.grid_columnconfigure((0, 1), weight=1, uniform="legende")
        self.status = customtkinter.CTkLabel(self, text="", anchor="w", text_color="#e67e22")
        self.status.grid(row=6, column=0, padx=15, pady=(0, 10), sticky="ew")
        self._poll()

    @staticmethod
    def _placement(parent, width, height):
        x = max(parent.winfo_x() - width - 10, 0)
        y = max(parent.winfo_y() - height // 3, 0)
        return f"{width}x{height}+{x}+{y}"

    # ------------------------------------------------------------ mode couleurs et notes
    def _toggle(self):
        try:
            self.mode.on() if self.switch.get() else self.mode.off()
            self.status.configure(text="")
        except Exception as e:
            logger.exception(e)
            self.switch.deselect()
            self.status.configure(text="Excel est occupé : réessayez après la saisie en cours.")
        self._show_legend()

    def suspend(self):
        """Retire couleurs et notes (avant la génération de la facture)."""
        self.mode.off()
        self.switch.deselect()
        self._show_legend()

    def _show_legend(self):
        """Légende des blocs colorés dans l'onglet, deux colonnes ; masquée quand le mode est éteint."""
        for child in self.legend.winfo_children():
            child.destroy()
        entrees = self.paquet.legende(self.mode.blocs) if self.mode.active else []
        if not entrees:
            self.legend.grid_remove()
            return
        for i, (bloc, couleur) in enumerate(entrees):
            customtkinter.CTkLabel(self.legend, text=bloc, fg_color=couleur, text_color=texte_sur(couleur),
                                   corner_radius=6, height=22, anchor="w").grid(
                row=i // 2, column=i % 2, padx=2, pady=2, sticky="ew", ipadx=6)
        self.legend.grid(row=5, column=0, padx=13, pady=(0, 4), sticky="ew")

    # ------------------------------------------------------------ suivi de la cellule active
    def _poll(self):
        try:
            cellule, valeur, onglet = self.locator.active()
            cle = (onglet, cellule.cle if cellule else None, repr(valeur))
            if cle != self.current:
                self.current = cle
                self.show(cellule, valeur, onglet)
        except Exception as e:  # Excel en cours d'édition (appel COM rejeté) : on réessaie
            logger.debug("Rayons X : sélection illisible (%s)", e)
        self._job = self.after(POLL_MS, self._poll)

    def show(self, cellule, valeur, onglet):
        fiche = self.paquet.fiche(onglet or "") or {}
        if cellule is None:  # hors donnée : la fiche du cas
            self.heading.configure(text=fiche.get("titre", "Cellule hors référentiel"))
            self.meta.configure(text=" · ".join(str(x) for x in (fiche.get("codes"), fiche.get("profil")) if x))
            self._set_text(self.xml, "(Sélectionnez une valeur de la facture pour voir sa balise XML.)")
            self._set_text(self.details, "\n\n".join(str(x) for x in (fiche.get("descriptif"), fiche.get("attention")) if x))
            return
        champ = cellule.champ
        self.heading.configure(text=cellule.titre)
        self.meta.configure(text=f"{champ.bloc} · profil {champ.profil} · cellule {champ.cle}")
        self._set_text(self.xml, "\n\n".join(fragment(x, valeur, champ.type) for x in champ.xpath))
        details = []
        commentaire = self.paquet.commentaire(onglet, champ)
        if commentaire:
            details.append(f"Dans ce cas : {commentaire}")
        details += [f"Chemin : {short_xpath(x)}" for x in champ.xpath]
        self._set_text(self.details, "\n".join(details))

    @staticmethod
    def _set_text(box, text):
        box.configure(state="normal")
        box.delete("0.0", "end")
        box.insert("end", text)
        box.configure(state="disabled")

    def close(self):
        if self._job:
            self.after_cancel(self._job)
            self._job = None
        try:
            self.mode.off()
        finally:
            if self.on_close:
                self.on_close()
            self.destroy()
