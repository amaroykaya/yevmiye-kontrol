import logging
import os
from datetime import datetime

import pandas as pd
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


OUTPUT_COLUMNS = [
    "Sıra No",
    "Boş",
    "Tarih",
    "Fatura No",
    "Firma",
    "KDV",
    "KDV'siz",
    "Toplam",
    "Etiket",
    "Ödeme Kanalı",
]

# Beklenen başlıklar (normalize edilmiş karşılaştırma için).
GIDER_EXPECTED_HEADERS = {
    1: ("no",),
    3: ("tarih",),
    4: ("firma", "firma/kisi", "firma/kişi"),
    8: ("toplam",),
    10: ("fatura no", "faturano"),
    11: ("kdv",),
    12: ("kdvsiz", "kdv'siz", "kdv siz"),
    13: ("etiket",),
}

GELIR_EXPECTED_HEADERS = {
    1: ("no",),
    2: ("tarih",),
    3: ("firma",),
    6: ("fatura no", "faturano"),
    8: ("kdv tl", "kdv"),
    9: ("kdv siz tl", "kdvsiz tl", "kdv'siz tl", "kdvsiz"),
    10: ("toplam tl", "toplam"),
}


def _to_float(value: object) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return 0.0
    text = text.replace(" ", "")
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def _format_date(value: object) -> str:
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.strftime("%d.%m.%Y")
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""
    return text.replace("/", ".")


def _safe_cell(df: pd.DataFrame, row_idx: int, col_idx: int) -> object:
    if col_idx >= df.shape[1]:
        return ""
    return df.iat[row_idx, col_idx]


def _norm_header(value: object) -> str:
    text = str(value or "").strip().lower()
    text = (
        text.replace("ı", "i")
        .replace("İ", "i")
        .replace("ş", "s")
        .replace("ğ", "g")
        .replace("ü", "u")
        .replace("ö", "o")
        .replace("ç", "c")
    )
    return " ".join(text.split())


def _validate_headers(
    df: pd.DataFrame,
    header_row_idx: int,
    expected: dict[int, tuple[str, ...]],
    sheet_label: str,
    logger: logging.Logger,
) -> None:
    if header_row_idx >= len(df):
        raise ValueError(f"{sheet_label}: başlık satırı bulunamadı (satır {header_row_idx + 1}).")

    missing: list[str] = []
    for col_idx, aliases in expected.items():
        actual = _norm_header(_safe_cell(df, header_row_idx, col_idx))
        if not any(alias in actual for alias in aliases):
            missing.append(f"kolon {col_idx + 1} (beklenen: {aliases[0]!r}, bulunan: {actual!r})")

    if missing:
        detail = "; ".join(missing)
        raise ValueError(
            f"{sheet_label} şablon uyuşmuyor. Beklenen kolonlar değişmiş olabilir: {detail}"
        )
    logger.info("%s şablon doğrulaması başarılı (başlık satırı %s).", sheet_label, header_row_idx + 1)


def read_gider_rows(file_path: str, logger: logging.Logger) -> list[dict]:
    logger.info("Gider dosyası okunuyor: %s", file_path)
    df = pd.read_excel(file_path, dtype=object, header=None)
    _validate_headers(df, 3, GIDER_EXPECTED_HEADERS, "Gider Excel", logger)
    rows: list[dict] = []

    # 5. satırdan (1-indexed) başlanır.
    for i in range(4, len(df)):
        try:
            b_val = str(_safe_cell(df, i, 1)).strip()  # B
            if not b_val or b_val.lower() == "nan" or b_val == "-":
                continue

            fatura_no = str(_safe_cell(df, i, 10)).strip()  # K
            if not fatura_no or fatura_no.lower() == "nan":
                continue

            kdv = _to_float(_safe_cell(df, i, 11))  # L
            kdvsiz = _to_float(_safe_cell(df, i, 12))  # M
            toplam = _to_float(_safe_cell(df, i, 8))  # I
            # Excel'de KDV 0 ve Toplam≈KDVsiz → tevkifat/istisna adayı (yevmiye ile doğrulanır).
            tevkifat_aday = abs(kdv) <= 0.01 and abs(toplam - kdvsiz) <= 0.05

            rows.append(
                {
                    "Tarih": _format_date(_safe_cell(df, i, 3)),  # D
                    "Fatura No": fatura_no,
                    "Firma": str(_safe_cell(df, i, 4)).strip(),  # E
                    "KDV": kdv,
                    "KDV'siz": kdvsiz,
                    "Toplam": toplam,
                    "Etiket": str(_safe_cell(df, i, 13)).strip(),  # N
                    "Ödeme Kanalı": str(_safe_cell(df, i, 9)).strip(),  # J
                    "_kaynak": "gider",
                    "_tevkifat_aday": tevkifat_aday,
                }
            )
        except Exception as exc:
            logger.warning("Gider satırı atlandı (index=%s): %s", i, exc)

    logger.info("Gider satır sayısı: %s", len(rows))
    logger.info("Gider tevkifat adayı (KDV=0) satır: %s", sum(1 for r in rows if r.get("_tevkifat_aday")))
    return rows


def read_gelir_rows(file_path: str, logger: logging.Logger) -> list[dict]:
    logger.info("Gelir dosyası okunuyor: %s", file_path)
    df = pd.read_excel(file_path, dtype=object, header=None)
    _validate_headers(df, 0, GELIR_EXPECTED_HEADERS, "Gelir Excel", logger)
    rows: list[dict] = []

    # 2. satırdan (1-indexed) başlanır.
    for i in range(1, len(df)):
        try:
            b_val = str(_safe_cell(df, i, 1)).strip()  # B
            if not b_val or b_val.lower() == "nan" or b_val == "-":
                continue

            fatura_no = str(_safe_cell(df, i, 6)).strip()  # G
            if not fatura_no or fatura_no.lower() == "nan":
                continue

            kdv = _to_float(_safe_cell(df, i, 8))  # I
            kdvsiz = _to_float(_safe_cell(df, i, 9))  # J
            toplam = _to_float(_safe_cell(df, i, 10))  # K
            istisna_aday = abs(kdv) <= 0.01

            rows.append(
                {
                    "Tarih": _format_date(_safe_cell(df, i, 2)),  # C
                    "Fatura No": fatura_no,
                    "Firma": str(_safe_cell(df, i, 3)).strip(),  # D
                    "KDV": kdv,
                    "KDV'siz": kdvsiz,
                    "Toplam": toplam,
                    "Etiket": "",
                    "Ödeme Kanalı": "",
                    "_kaynak": "gelir",
                    "_istisna_aday": istisna_aday,
                }
            )
        except Exception as exc:
            logger.warning("Gelir satırı atlandı (index=%s): %s", i, exc)

    logger.info("Gelir satır sayısı: %s", len(rows))
    logger.info("Gelir istisna adayı (KDV=0) satır: %s", sum(1 for r in rows if r.get("_istisna_aday")))
    return rows


def build_combined_rows(gider_rows: list[dict], gelir_rows: list[dict]) -> list[dict]:
    combined: list[dict] = []
    ordered_rows = list(gider_rows) + list(gelir_rows)

    for idx, row in enumerate(ordered_rows, start=1):
        combined.append(
            {
                "Sıra No": idx,
                "Boş": "",
                "Tarih": row.get("Tarih", ""),
                "Fatura No": row.get("Fatura No", ""),
                "Firma": row.get("Firma", ""),
                "KDV": float(row.get("KDV", 0.0)),
                "KDV'siz": float(row.get("KDV'siz", 0.0)),
                "Toplam": float(row.get("Toplam", 0.0)),
                "Etiket": row.get("Etiket", ""),
                "Ödeme Kanalı": row.get("Ödeme Kanalı", ""),
                "_kaynak": row.get("_kaynak", ""),
                "_tevkifat_aday": bool(row.get("_tevkifat_aday", False)),
                "_istisna_aday": bool(row.get("_istisna_aday", False)),
            }
        )

    return combined
