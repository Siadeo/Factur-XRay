"""Utilitaires Windows : chemins des ressources, vues protégées d'Excel, imprimante PDF."""
import logging
import sys
from pathlib import Path

logger = logging.getLogger(__name__)


def resource_path(relative_path: str) -> Path:
    """Chemin d'une ressource livrée à côté de l'exécutable (ou à la racine du projet en développement)."""
    if getattr(sys, "frozen", False):
        base_path = Path(sys.executable).resolve().parent
    else:
        base_path = Path(__file__).resolve().parent.parent
    return (base_path / relative_path).resolve()


def close_all_protected_sheets():
    """Ferme les classeurs ouverts en mode protégé au lancement : xlwings renvoie alors un objet
    ProtectedViewWindow au lieu d'un Book, ce qui provoque des erreurs."""
    import win32com.client
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except BaseException:
        return
    try:
        pv = excel.ProtectedViewWindows
        for i in range(pv.Count, 0, -1):  # en sens inverse : les index changent à chaque fermeture
            try:
                pv.Item(i).Close()
            except BaseException as e:
                logger.error(f"Impossible de fermer le fichier en mode protégé : {e}")
    except BaseException as e:
        logger.error(f"Erreur lors de la gestion des ProtectedViewWindows : {e}")


def disable_win_auto_printer_manage():
    """Désactive « Laisser Windows gérer mon imprimante par défaut » (sinon Excel n'exporte pas le PDF)."""
    import winreg
    key_path = r"Software\Microsoft\Windows NT\CurrentVersion\Windows"
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
        winreg.SetValueEx(key, "LegacyDefaultPrinterMode", 0, winreg.REG_DWORD, 1)


def use_pdf_printer() -> bool:
    """Bascule l'imprimante par défaut sur « Microsoft Print To PDF » (mise en page de l'export PDF)."""
    import win32print
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    for _, _, name, _ in printers:
        try:
            handle = win32print.OpenPrinter(name)
            driver = win32print.GetPrinter(handle, 2)["pDriverName"]
            win32print.ClosePrinter(handle)
            if driver == "Microsoft Print To PDF":
                win32print.SetDefaultPrinter(name)
                return True
        except Exception:
            continue
    return False
