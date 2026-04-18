import logging
import sys
from pathlib import Path
from typing import Union


def create_app_log_formatter() -> logging.Formatter:
    """Konsol, app.log ve run log dosyası için ortak biçim."""
    return logging.Formatter(
        fmt="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def attach_run_file_log(logger: logging.Logger, log_path: Union[str, Path]) -> logging.FileHandler:
    """
    Tek bir reconciliation çalışması için UTF-8 metin dosyasına yazar.
    Dosya her seferinde sıfırdan oluşturulur (mode=w).
    """
    path = Path(log_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    handler = logging.FileHandler(path, mode="w", encoding="utf-8")
    handler.setLevel(logging.INFO)
    handler.setFormatter(create_app_log_formatter())
    logger.addHandler(handler)
    return handler


def detach_run_file_log(logger: logging.Logger, handler: logging.FileHandler) -> None:
    if handler in logger.handlers:
        logger.removeHandler(handler)
    handler.close()


def setup_logger() -> logging.Logger:
    """Uygulama için terminal + logs/app.log loglamasını hazırlar."""
    log_dir = Path("logs")
    log_dir.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger("yevmiye_kontrol")
    logger.setLevel(logging.INFO)
    logger.propagate = False

    if logger.handlers:
        return logger

    formatter = create_app_log_formatter()

    file_handler = logging.FileHandler(log_dir / "app.log", encoding="utf-8")
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setLevel(logging.INFO)
    stream_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)
    return logger
