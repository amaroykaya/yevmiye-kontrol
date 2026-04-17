import logging
from pathlib import Path

import pandas as pd

COLUMN_NAMES = [
    "hesap_kodu",
    "hesap_adi",
    "aciklama",
    "detay",
    "borc",
    "alacak",
]


def _normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Veriyi fiş ayırmaya uygun hale getirir."""
    normalized = df.iloc[:, :6].copy()
    normalized.columns = COLUMN_NAMES
    normalized = normalized.fillna("")

    # String alanlarda tip güvenliği sağlanır.
    normalized["hesap_kodu"] = normalized["hesap_kodu"].astype(str)
    normalized["aciklama"] = normalized["aciklama"].astype(str)

    return normalized


def _row_text_for_matching(row: pd.Series) -> str:
    """Satır içindeki metinleri birleştirip karşılaştırma metni üretir."""
    text_parts = []
    for column in ["hesap_kodu", "hesap_adi", "aciklama", "detay"]:
        value = row.get(column, "")
        text_parts.append(str(value))
    return " ".join(text_parts).upper()


def _row_contains_keyword(row: pd.Series, keyword: str) -> bool:
    target = keyword.strip().upper()
    for column in ["hesap_kodu", "hesap_adi", "aciklama", "detay"]:
        value = str(row.get(column, "")).strip().upper()
        if target in value:
            return True
    return False


def parse_fis_blocks(excel_path: str, logger: logging.Logger) -> list[pd.DataFrame]:
    """
    Yevmiye dosyasını okuyup MAHSUP-TOPLAM bloklarına göre fişleri ayırır.
    """
    path_obj = Path(excel_path)
    logger.info("Excel okuma başladı: %s", path_obj)

    # İlk MAHSUP satırının header tarafından yutulmaması için dosya başlıksız okunur.
    raw_df = pd.read_excel(path_obj, header=None)
    df = _normalize_dataframe(raw_df.iloc[:, :6].reset_index(drop=True))

    logger.info("Excel okuma tamamlandı.")
    logger.info("Fiş ayırma başladı.")

    fisler: list[pd.DataFrame] = []
    current_rows: list[dict] = []
    inside_fis = False
    first_mahsup_captured = False
    current_start_idx: int | None = None

    def close_current_fis(end_idx: int) -> None:
        nonlocal current_rows, current_start_idx
        if not current_rows:
            return
        fis_df = pd.DataFrame(current_rows, columns=COLUMN_NAMES)
        fisler.append(fis_df)
        fis_no = len(fisler)
        start_idx = current_start_idx if current_start_idx is not None else end_idx
        mahsup_count = sum(1 for _, r in fis_df.iterrows() if _row_contains_keyword(r, "MAHSUP"))
        toplam_count = sum(1 for _, r in fis_df.iterrows() if _row_contains_keyword(r, "TOPLAM"))
        logger.info(
            "Fiş %s | start=%s end=%s satır=%s mahsup_sayisi=%s toplam_sayisi=%s",
            fis_no,
            start_idx + 1,
            end_idx + 1,
            len(fis_df),
            mahsup_count,
            toplam_count,
        )
        if mahsup_count > 1:
            logger.warning("Fiş %s: Birleşmiş fiş olabilir", fis_no)
        if toplam_count == 0:
            logger.warning("Fiş %s: TOPLAM bulunamadı", fis_no)
        if len(fis_df) > 25:
            logger.warning("Fiş %s: Anormal uzun fiş", fis_no)
        current_rows = []
        current_start_idx = None

    for row_idx, row in df.iterrows():
        row_dict = row.to_dict()

        if _row_contains_keyword(row, "MAHSUP"):
            if inside_fis and current_rows:
                close_current_fis(row_idx - 1)
            inside_fis = True
            current_start_idx = row_idx
            if not first_mahsup_captured:
                first_mahsup_captured = True
                logger.info("İlk MAHSUP satırı yakalandı: Evet")
                logger.info("İlk fiş başlangıç satırı: %s", row_idx + 1)

        if inside_fis:
            current_rows.append(row_dict)

        if inside_fis and _row_contains_keyword(row, "TOPLAM"):
            close_current_fis(row_idx)
            inside_fis = False

    # Dosya TOPLAM ile bitmezse son açık fişi de koru.
    if inside_fis and current_rows:
        close_current_fis(len(df) - 1)

    if not first_mahsup_captured:
        logger.warning("İlk MAHSUP satırı yakalandı: Hayır")

    logger.info("Fiş ayırma tamamlandı.")
    logger.info("Toplam fiş sayısı: %s", len(fisler))

    for idx, fis in enumerate(fisler[:3], start=1):
        logger.info("Fiş %s satır sayısı: %s", idx, len(fis))

    return fisler
