import logging
import re
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

# 00001-----00003-----MAHSUP-----01/01/2025
FIS_HEADER_REGEX = re.compile(
    r"^(\d+)\-+(\d+)\-+([A-ZÇĞİÖŞÜ]+)\-+(\d{2}[./-]\d{2}[./-]\d{4})\s*$",
    re.IGNORECASE,
)


def _normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Veriyi fiş ayırmaya uygun hale getirir."""
    normalized = df.iloc[:, :6].copy()
    normalized.columns = COLUMN_NAMES
    normalized = normalized.fillna("")

    normalized["hesap_kodu"] = normalized["hesap_kodu"].astype(str)
    normalized["aciklama"] = normalized["aciklama"].astype(str)

    return normalized


def _row_contains_keyword(row: pd.Series, keyword: str) -> bool:
    target = keyword.strip().upper()
    for column in ["hesap_kodu", "hesap_adi", "aciklama", "detay"]:
        value = str(row.get(column, "")).strip().upper()
        if target in value:
            return True
    return False


def _row_combined_upper(row: pd.Series) -> str:
    parts = [str(row.get(col, "")) for col in COLUMN_NAMES]
    return " ".join(parts).upper()


def parse_fis_header(text: object) -> dict | None:
    """Fiş giriş satırını parse eder; yoksa None."""
    if text is None:
        return None
    raw = str(text).strip()
    if not raw or raw.lower() == "nan":
        return None
    match = FIS_HEADER_REGEX.match(raw)
    if not match:
        return None
    fis_no_bas, fis_no_bit, fis_tipi, tarih_raw = match.groups()
    tarih = tarih_raw.replace("-", ".").replace("/", ".")
    return {
        "fis_no_bas": fis_no_bas,
        "fis_no_bit": fis_no_bit,
        "fis_tipi": fis_tipi.upper(),
        "tarih": tarih,
        "raw": raw,
    }


def _is_fis_start(row: pd.Series) -> tuple[bool, dict | None]:
    """Başlık regex veya MAHSUP yedek ile fiş başlangıcı."""
    for column in ["hesap_kodu", "hesap_adi", "aciklama"]:
        header = parse_fis_header(row.get(column, ""))
        if header:
            return True, header
    if _row_contains_keyword(row, "MAHSUP"):
        return True, None
    return False, None


def parse_fis_blocks(excel_path: str, logger: logging.Logger) -> list[pd.DataFrame]:
    """
    Yevmiye dosyasını okuyup fiş bloklarına ayırır.
    Birincil sınır: NNNNN-----NNNNN-----TIP-----tarih
    Yedek: MAHSUP ... TOPLAM
    """
    path_obj = Path(excel_path)
    logger.info("Excel okuma başladı: %s", path_obj)

    raw_df = pd.read_excel(path_obj, header=None)
    df = _normalize_dataframe(raw_df.iloc[:, :6].reset_index(drop=True))

    logger.info("Excel okuma tamamlandı.")
    logger.info("Fiş ayırma başladı.")

    fisler: list[pd.DataFrame] = []
    current_rows: list[dict] = []
    inside_fis = False
    first_header_captured = False
    current_start_idx: int | None = None
    current_header: dict | None = None
    header_parsed_count = 0
    mahsup_fallback_count = 0

    def close_current_fis(end_idx: int) -> None:
        nonlocal current_rows, current_start_idx, current_header
        if not current_rows:
            return
        fis_df = pd.DataFrame(current_rows, columns=COLUMN_NAMES)
        if current_header:
            fis_df.attrs["fis_header"] = dict(current_header)
        fisler.append(fis_df)
        fis_no = len(fisler)
        start_idx = current_start_idx if current_start_idx is not None else end_idx
        mahsup_count = sum(1 for _, r in fis_df.iterrows() if _row_contains_keyword(r, "MAHSUP"))
        toplam_count = sum(1 for _, r in fis_df.iterrows() if _row_contains_keyword(r, "TOPLAM"))
        logger.info(
            "Fiş %s | start=%s end=%s satır=%s mahsup_sayisi=%s toplam_sayisi=%s header=%s",
            fis_no,
            start_idx + 1,
            end_idx + 1,
            len(fis_df),
            mahsup_count,
            toplam_count,
            current_header.get("raw") if current_header else "-",
        )
        if mahsup_count > 1:
            logger.warning("Fiş %s: Birleşmiş fiş olabilir", fis_no)
        if toplam_count == 0:
            logger.warning("Fiş %s: TOPLAM bulunamadı", fis_no)
        if len(fis_df) > 25:
            logger.warning("Fiş %s: Anormal uzun fiş", fis_no)
        current_rows = []
        current_start_idx = None
        current_header = None

    for row_idx, row in df.iterrows():
        row_dict = row.to_dict()
        is_start, header = _is_fis_start(row)

        if is_start:
            if inside_fis and current_rows:
                close_current_fis(row_idx - 1)
            inside_fis = True
            current_start_idx = row_idx
            current_header = header
            if header:
                header_parsed_count += 1
            else:
                mahsup_fallback_count += 1
            if not first_header_captured:
                first_header_captured = True
                logger.info("İlk fiş başlangıç satırı: %s", row_idx + 1)
                if header:
                    logger.info("İlk fiş başlığı parse edildi: %s", header.get("raw"))

        if inside_fis:
            current_rows.append(row_dict)

        if inside_fis and _row_contains_keyword(row, "TOPLAM"):
            close_current_fis(row_idx)
            inside_fis = False

    if inside_fis and current_rows:
        close_current_fis(len(df) - 1)

    if not first_header_captured:
        logger.warning("İlk fiş başlangıcı yakalanamadı")

    logger.info("Fiş ayırma tamamlandı.")
    logger.info("Toplam fiş sayısı: %s", len(fisler))
    logger.info("Başlık regex ile parse edilen: %s", header_parsed_count)
    logger.info("MAHSUP yedek ile açılan: %s", mahsup_fallback_count)

    for idx, fis in enumerate(fisler[:3], start=1):
        logger.info("Fiş %s satır sayısı: %s", idx, len(fis))

    return fisler
