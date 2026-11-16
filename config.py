"""Configuration publique de l'application."""
import logging

# Niveaux : DEBUG, INFO, WARNING, ERROR, CRITICAL
LEVEL = logging.WARNING
HANDLERS = [logging.FileHandler("log.log", encoding="utf-8")]

if LEVEL == logging.DEBUG:
    HANDLERS.append(logging.StreamHandler())


def setup_logger():
    logging.basicConfig(
        level=LEVEL,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        handlers=HANDLERS,
        force=True,  # réinitialise la configuration si logging.basicConfig() a déjà été appelé
    )
