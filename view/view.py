import customtkinter

BLEU, BLEU_SURVOL, GRIS = "#3498db", "#2980b9", "#34495e"


class AppView(customtkinter.CTk):
    """Fenêtre principale : ouvrir un classeur d'exemples, générer la facture de l'onglet actif, Rayons X."""

    def __init__(self):
        super().__init__()
        self.buttons = {}
        self.alert_window = None
        self.on_quit_confirmed = None  # rappel du contrôleur (fermeture des classeurs d'exemples)
        self.protocol("WM_DELETE_WINDOW", self.on_quit)

        customtkinter.set_appearance_mode("dark")
        self.title("Factur-Xray")
        self.geometry(self.window_to_right_bottom("320", "340"))
        self.grid_columnconfigure(0, weight=1)

        menu = customtkinter.CTkFrame(self, fg_color="#242424")
        menu.grid(row=0, column=0, padx=15, pady=(10, 0), sticky="w")
        self.menu_ouvrir = customtkinter.CTkOptionMenu(
            menu, values=[], command=None, corner_radius=10, width=210, fg_color=GRIS, button_color=GRIS,
            button_hover_color="#2c3e50", dropdown_fg_color=("#ffffff", "#ffffff"),
            dropdown_text_color="#333333", dropdown_hover_color="#bdc3c7")
        self.menu_ouvrir.grid(row=0, column=0, padx=5, pady=10, sticky="w")
        self.menu_ouvrir.set("Ouvrir")
        self.buttons["aide"] = customtkinter.CTkButton(menu, text="Aide", width=60, fg_color=GRIS)
        self.buttons["aide"].grid(row=0, column=1, padx=5, pady=10, sticky="w")

        actions = customtkinter.CTkFrame(self, fg_color="#242424")
        actions.grid(row=1, column=0, padx=15, pady=(0, 5), sticky="ew")
        self.buttons["generer"] = customtkinter.CTkButton(
            actions, text="Générer la facture", width=290, fg_color="#38ADA9", hover_color="#4FC3BF")
        self.buttons["generer"].grid(row=0, column=0, padx=5, pady=10, sticky="ew")
        # Rayons X : le BT EN 16931 et le XML derrière chaque cellule
        self.buttons["xray"] = customtkinter.CTkButton(
            actions, text="Rayons X (BT et XML)", width=290, fg_color="#4A69BD", hover_color="#6A89CC")
        self.buttons["xray"].grid(row=1, column=0, padx=5, pady=(0, 10), sticky="ew")

        self.textbox_feedback = customtkinter.CTkTextbox(
            master=self, width=280, height=150, corner_radius=10, wrap="word", fg_color="#242424", state="disabled")
        self.textbox_feedback.grid(row=2, column=0, padx=10, pady=5, sticky="nsew")

    def set_actions(self, actions):
        for name, action in actions.items():
            if name in self.buttons:
                self.buttons[name].configure(command=action)

    def set_menu_ouvrir(self, action, values):
        self.menu_ouvrir.configure(command=action, values=values)

    def reset_menu_ouvrir(self):
        self.menu_ouvrir.set("Ouvrir")

    def window_to_right_bottom(self, width, height):
        x = self.winfo_screenwidth() - int(width) - 20
        y = self.winfo_screenheight() - int(height) - 40
        return f"{width}x{height}+{x}+{y}"

    def show_feedback(self, txt="", message_type="feedback", stack=False):
        """Affiche un message : message_type = "feedback", "error" ou "success" ; stack empile les messages."""
        couleurs = {"feedback": BLEU, "error": "#e67e22", "success": "#27ae60"}
        self.textbox_feedback.configure(state="normal", text_color=couleurs[message_type])
        if not stack:
            self.textbox_feedback.delete("0.0", "end")
        self.textbox_feedback.insert("end", txt + "\n")
        self.textbox_feedback.configure(state="disabled")
        self.update_idletasks()

    def close(self):
        self.quit()
        self.destroy()

    def on_quit(self):
        if self.alert_window is None or not self.alert_window.winfo_exists():
            self.alert_window = AlertWindowQuit(self, txt="Fermer l'application ?\n"
                                                          "Les classeurs d'exemples ouverts seront fermés\n"
                                                          "sans enregistrer les modifications.")
        else:
            self.alert_window.focus()


class AlertWindowQuit(customtkinter.CTkToplevel):
    """Confirmation avant de quitter l'application."""

    def __init__(self, parent, txt=""):
        super().__init__()
        self.parent = parent
        self.title("Fermeture de l'application")
        largeur, hauteur = 340, 160
        x = self.winfo_screenwidth() // 2 - largeur // 2
        y = self.winfo_screenheight() // 2 - hauteur // 2
        self.geometry(f"{largeur}x{hauteur}+{x}+{y}")
        customtkinter.CTkLabel(self, text=txt).pack(padx=20, pady=20)
        customtkinter.CTkButton(self, text="OUI", command=self.on_yes).pack(side="left", padx=20, pady=20)
        customtkinter.CTkButton(self, text="NON", command=self.destroy).pack(side="right", padx=20, pady=20)

    def on_yes(self):
        if self.parent.on_quit_confirmed:
            self.parent.on_quit_confirmed()
        self.destroy()
        self.parent.close()
