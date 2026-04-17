import tkinter as tk

from core.logger_setup import setup_logger
from ui.main_window import MainWindow


def main() -> None:
    logger = setup_logger()
    logger.info("Uygulama açıldı.")

    root = tk.Tk()
    MainWindow(root=root, logger=logger)
    root.mainloop()


if __name__ == "__main__":
    main()
