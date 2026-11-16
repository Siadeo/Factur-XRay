from config import setup_logger
from controller import Controller
from view import AppView

setup_logger()


if __name__ == "__main__":
    app_view = AppView()
    app_view.update()
    Controller(view=app_view)
    app_view.mainloop()
