import logging
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

# 320/120/300/100 birincil; 102/770 banka ödemeli gider fişleri için yedek.
MAIN_ACCOUNT_PRIORITY = [
    "320",
    "120",
    "300",
    "100",
    "102",
    "770",
    "760",
    "750",
    "740",
    "730",
    "255",
]
LABEL_MAPPING = {
    "730.01": "ÜRETİM",
    "750.30": "ANTEN",
    "750.10": "ENDÜSTRİYEL",
    "770.10": "GENEL",
    "750.20": "SAVUNMA",
}


def _default_year() -> str:
    return str(datetime.now().year)


SUPPORTED_DOC_TYPES = ["EF", "FT", "EA", "PO", "FŞ", "FS", "SM", "DK"]
DOCUMENT_LINE_REGEX = re.compile(
    r"^\s*(EF|FT|EA|PO|FŞ|FS|SM|DK)\s+(\S+)(?:\s+(\d{2}[./-]\d{2}))?(?:\s+(.*))?$",
    re.IGNORECASE,
)
SHORT_DATE_REGEX = re.compile(r"(\d{2}[./-]\d{2})(?![./-]\d{2,4})")
FIRMA_EXCLUDE_WORDS = ["KDV", "GİDER", "GIDER", "HESAPLANAN", "İNDİRİLECEK", "INDIRILECEK"]
FIRMA_EXCLUDE_EXTRA = [
    "KREDİ KARTI ÖDEME",
    "KREDİ KARTI ODEME",
    "BANKA",
    "SATICILAR",
    "ALICILAR",
    "GENEL YÖNETİM",
    "ÜRETİM",
    "ENDÜSTRİYEL",
    "SAVUNMA",
    "ANTEN",
    "GENEL",
]
LABEL_KEYWORDS = {
    "ÜRETİM": "ÜRETİM",
    "ENDÜSTRİYEL": "ENDÜSTRİYEL",
    "ANTEN": "ANTEN",
    "SAVUNMA": "SAVUNMA",
    "GENEL": "GENEL",
}


def to_decimal(value: object) -> float:
    """Excel sayilarini guvenli sekilde float'a cevirir."""
    if value is None:
        return 0.0

    if isinstance(value, (int, float)) and not pd.isna(value):
        return float(value)

    if pd.isna(value):
        return 0.0

    text = str(value).strip()
    if not text:
        return 0.0

    # String temizligi sadece stringlerde uygulanir.
    text = text.replace(" ", "")
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")

    try:
        return float(text)
    except (TypeError, ValueError):
        return 0.0


def _row_combined_text(row: pd.Series) -> str:
    return " ".join(str(row.get(col, "")) for col in ["hesap_kodu", "hesap_adi", "aciklama", "detay"])


def _extract_year_from_source(fis_df: pd.DataFrame, file_path: str) -> str:
    first_row_text = _row_combined_text(fis_df.iloc[0]) if not fis_df.empty else ""
    match = re.search(r"(20\d{2})", first_row_text)
    if match:
        return match.group(1)

    filename = Path(file_path).name
    match = re.search(r"(20\d{2})", filename)
    if match:
        return match.group(1)

    return _default_year()


def parse_document_line(text: str | None) -> dict | None:
    if not text:
        return None
    normalized = " ".join(str(text).strip().split())
    if not normalized:
        return None

    match = DOCUMENT_LINE_REGEX.match(normalized)
    if not match:
        return None

    belge_tipi, fatura_no, short_date, firma = match.groups()
    return {
        "belge_tipi": belge_tipi.upper(),
        "fatura_no": (fatura_no or "").strip(),
        "kisa_tarih": (short_date or "").replace("-", "/").replace(".", "/"),
        "firma": (firma or "").strip(),
    }


def _collect_document_candidates(fis_df: pd.DataFrame) -> list[dict]:
    candidates: list[dict] = []
    for _, row in fis_df.iterrows():
        text = str(row.get("aciklama", "")).strip()
        parsed = parse_document_line(text)
        if parsed:
            candidates.append(parsed)
    return candidates


def find_best_document_info(fis_df: pd.DataFrame) -> dict:
    candidates = _collect_document_candidates(fis_df)
    if not candidates:
        return {
            "belge_tipi": "YOK",
            "fatura_no": "",
            "kisa_tarih": "",
            "firma": "",
            "_belge_aday_sayisi": 0,
            "_cok_belgeli": False,
        }

    unique_faturas = {
        (c.get("belge_tipi", ""), str(c.get("fatura_no", "")).strip().upper())
        for c in candidates
        if str(c.get("fatura_no", "")).strip()
    }
    cok_belgeli = len(unique_faturas) > 1

    def candidate_key(item: dict) -> tuple[int, int]:
        has_firma = 1 if item.get("firma", "").strip() else 0
        firma_len = len(item.get("firma", "").strip())
        return has_firma, firma_len

    best = max(candidates, key=candidate_key)
    best = dict(best)
    best["_belge_aday_sayisi"] = len(candidates)
    best["_cok_belgeli"] = cok_belgeli
    best["_benzersiz_fatura_sayisi"] = len(unique_faturas)
    return best


def _header_from_fis(fis_df: pd.DataFrame) -> dict | None:
    header = getattr(fis_df, "attrs", {}).get("fis_header")
    if isinstance(header, dict) and header.get("tarih"):
        return header
    if fis_df.empty:
        return None
    from core.yevmiye_parser import parse_fis_header

    for column in ["hesap_kodu", "hesap_adi", "aciklama"]:
        parsed = parse_fis_header(fis_df.iloc[0].get(column, ""))
        if parsed:
            return parsed
    return None


def _extract_date(fis_df: pd.DataFrame, file_path: str, doc_info: dict) -> str:
    if fis_df.empty:
        return ""

    header = _header_from_fis(fis_df)
    if header and header.get("tarih"):
        return str(header["tarih"])

    first_row_text = _row_combined_text(fis_df.iloc[0])
    full_date_match = re.search(r"(\d{2})[./-](\d{2})[./-](\d{4})", first_row_text)
    if full_date_match:
        day, month, year = full_date_match.groups()
        return f"{day}.{month}.{year}"

    year = _extract_year_from_source(fis_df, file_path)
    if doc_info.get("kisa_tarih"):
        day, month = doc_info["kisa_tarih"].split("/")
        return f"{day}.{month}.{year}"

    return ""


def _fis_has_account_prefix(fis_df: pd.DataFrame, prefix: str) -> bool:
    for value in fis_df["hesap_kodu"].tolist():
        code = str(value).strip()
        if code.startswith(prefix):
            return True
    return False


def _extract_main_account(fis_df: pd.DataFrame) -> str:
    codes = [str(v).strip() for v in fis_df["hesap_kodu"].tolist()]
    for main_code in MAIN_ACCOUNT_PRIORITY:
        if any(code.startswith(main_code) for code in codes):
            return main_code
    return ""


def _is_detail_row(hesap_kodu: str, aciklama: str) -> bool:
    code = hesap_kodu.strip()
    if code.count(".") >= 1:
        return True
    if parse_document_line(aciklama or ""):
        return True
    return False


def _is_kdv_detail_row(hesap_kodu: str, aciklama: str) -> bool:
    code = hesap_kodu.strip()
    if code.count(".") >= 2:
        return True
    if parse_document_line(aciklama or ""):
        return True
    return False


def _is_valid_firma_text(text: str) -> bool:
    if not text:
        return False
    upper_text = text.upper()
    excluded = FIRMA_EXCLUDE_WORDS + FIRMA_EXCLUDE_EXTRA
    if any(word in upper_text for word in excluded):
        return False
    if "%" in upper_text:
        return False
    return True


def _extract_company_name_from_accounts(fis_df: pd.DataFrame) -> str:
    candidates: list[str] = []
    for _, row in fis_df.iterrows():
        code = str(row.get("hesap_kodu", "")).strip()
        hesap_adi = str(row.get("hesap_adi", "")).strip()
        if not code or not hesap_adi:
            continue
        if not (code.startswith("320") or code.startswith("120") or code.startswith("300") or code.startswith("100") or code.startswith("102")):
            continue
        if not _is_detail_row(code, str(row.get("aciklama", ""))):
            continue
        if not _is_valid_firma_text(hesap_adi):
            continue
        candidates.append(hesap_adi)

    if not candidates:
        return ""
    return max(candidates, key=lambda x: len(x.strip()))


def _extract_company_name_from_aciklama_fallback(fis_df: pd.DataFrame) -> str:
    candidates: list[str] = []
    for _, row in fis_df.iterrows():
        text = str(row.get("aciklama", "")).strip()
        if not text:
            continue
        if parse_document_line(text):
            continue
        date_match = SHORT_DATE_REGEX.search(text)
        if not date_match:
            continue
        after_date = text[date_match.end():].strip()
        if not _is_valid_firma_text(after_date):
            continue
        candidates.append(after_date)

    if not candidates:
        return ""
    return max(candidates, key=lambda x: len(x.strip()))


def _extract_company_name(fis_df: pd.DataFrame, doc_info: dict) -> tuple[str, bool]:
    if doc_info.get("firma"):
        return doc_info["firma"].strip(), False

    from_accounts = _extract_company_name_from_accounts(fis_df)
    if from_accounts:
        return from_accounts, True

    from_aciklama = _extract_company_name_from_aciklama_fallback(fis_df)
    if from_aciklama:
        return from_aciklama, True

    return "", False


def _amount_from_row(row: pd.Series) -> float:
    detail = to_decimal(row.get("detay", ""))
    if detail != 0:
        return detail

    borc = to_decimal(row.get("borc", ""))
    if borc != 0:
        return borc

    return to_decimal(row.get("alacak", ""))


def _sum_kdv(fis_df: pd.DataFrame, main_account: str) -> tuple[float, int, list[str]]:
    if main_account == "120":
        target_prefix = "391"
    else:
        target_prefix = "191"

    total = 0.0
    source_count = 0
    source_codes: list[str] = []
    for _, row in fis_df.iterrows():
        code = str(row.get("hesap_kodu", "")).strip()
        aciklama = str(row.get("aciklama", ""))
        if not code.startswith(target_prefix):
            continue
        if not _is_kdv_detail_row(code, aciklama):
            continue
        amount = _amount_from_row(row)
        if amount != 0:
            source_count += 1
            total += amount
            source_codes.append(code)

    return round(total, 2), source_count, source_codes


def _extract_toplam_from_total_row(fis_df: pd.DataFrame) -> tuple[float, str]:
    for _, row in fis_df.iterrows():
        row_text = _row_combined_text(row).upper()
        if "TOPLAM" not in row_text:
            continue

        borc = to_decimal(row.get("borc", ""))
        alacak = to_decimal(row.get("alacak", ""))
        if borc != 0 and alacak != 0:
            if abs(borc - alacak) < 0.01:
                return round(borc, 2), "TOPLAM"
            return round(max(borc, alacak), 2), "TOPLAM"
        if borc != 0:
            return round(borc, 2), "TOPLAM"
        if alacak != 0:
            return round(alacak, 2), "TOPLAM"
    return 0.0, ""


def _extract_toplam_fallback(fis_df: pd.DataFrame, main_account: str) -> tuple[float, str]:
    if main_account == "120":
        fallback_codes = ["120"]
    elif main_account in {"102", "100"}:
        fallback_codes = [main_account, "770", "760", "750"]
    else:
        fallback_codes = ["320", "300", "100", "102"]
    for code_prefix in fallback_codes:
        for _, row in fis_df.iterrows():
            code = str(row.get("hesap_kodu", "")).strip()
            aciklama = str(row.get("aciklama", ""))
            if not code.startswith(code_prefix):
                continue
            if not _is_detail_row(code, aciklama):
                continue
            amount = _amount_from_row(row)
            if amount != 0:
                return round(amount, 2), "YEDEK"

    return 0.0, ""


def _extract_etiket(fis_df: pd.DataFrame, main_account: str) -> str:
    if main_account == "120":
        return ""

    for _, row in fis_df.iterrows():
        code = str(row.get("hesap_kodu", "")).strip()
        for target_code, label in LABEL_MAPPING.items():
            if code.startswith(target_code):
                return label

    for _, row in fis_df.iterrows():
        hesap_adi = str(row.get("hesap_adi", "")).upper()
        for keyword, label in LABEL_KEYWORDS.items():
            if keyword in hesap_adi:
                return label
    if main_account in {"300", "320", "102", "770"}:
        return "GENEL"
    return ""


def extract_fis_summary(fis_df: pd.DataFrame, index: int, file_path: str) -> dict:
    main_account = _extract_main_account(fis_df)
    doc_info = find_best_document_info(fis_df)
    header = _header_from_fis(fis_df)
    tarih = _extract_date(fis_df, file_path, doc_info)
    firma, fallback_firma = _extract_company_name(fis_df, doc_info)

    kdv, kdv_source_count, kdv_source_codes = (
        _sum_kdv(fis_df, main_account) if main_account else (0.0, 0, [])
    )
    toplam, toplam_source = _extract_toplam_from_total_row(fis_df)
    if toplam == 0:
        toplam, toplam_source = _extract_toplam_fallback(fis_df, main_account)

    mal_hizmet = toplam - kdv
    if -0.01 < mal_hizmet < 0.01:
        mal_hizmet = 0.0
    mal_hizmet = round(mal_hizmet, 2)
    etiket = _extract_etiket(fis_df, main_account)
    has_360 = _fis_has_account_prefix(fis_df, "360")
    # Tevkifat: 360 (sorumlu sifariş KDV) + indirilecek KDV satırı birlikte.
    tevkifatli = has_360 and kdv > 0

    return {
        "sira_no": index,
        "hesap_kodu": main_account,
        "tarih": tarih,
        "fatura_no": doc_info.get("fatura_no", ""),
        "firma": firma,
        "kdv": round(kdv, 2),
        "mal_hizmet": mal_hizmet,
        "toplam": round(toplam, 2),
        "etiket": etiket,
        "_kdv_source_count": kdv_source_count,
        "_kdv_detay_kodlari": kdv_source_codes,
        "_toplam_source": toplam_source or "YOK",
        "_belge_tipi": doc_info.get("belge_tipi", "YOK"),
        "_belge_satiri_bulundu": doc_info.get("belge_tipi", "YOK") != "YOK",
        "_fallback_firma": fallback_firma,
        "_kisa_tarih": doc_info.get("kisa_tarih", ""),
        "_fis_no_bas": (header or {}).get("fis_no_bas", ""),
        "_fis_no_bit": (header or {}).get("fis_no_bit", ""),
        "_fis_tipi": (header or {}).get("fis_tipi", ""),
        "_has_360": has_360,
        "_tevkifatli": tevkifatli,
        "_cok_belgeli": bool(doc_info.get("_cok_belgeli", False)),
        "_benzersiz_fatura_sayisi": int(doc_info.get("_benzersiz_fatura_sayisi", 0) or 0),
    }


def build_fis_summary_list(
    fisler: list[pd.DataFrame], file_path: str, logger: logging.Logger
) -> list[dict]:
    logger.info("Fiş özet çıkarma başladı.")
    summaries: list[dict] = []

    for idx, fis_df in enumerate(fisler, start=1):
        try:
            summaries.append(extract_fis_summary(fis_df, idx, file_path))
        except Exception as exc:
            logger.exception("Fiş %s özetlenirken hata oluştu: %s", idx, exc)

    logger.info("Fiş özet çıkarma tamamlandı. Toplam özet sayısı: %s", len(summaries))

    belge_bulunan = sum(1 for s in summaries if s.get("_belge_satiri_bulundu", False))
    belge_bulunamayan = len(summaries) - belge_bulunan
    fallback_firma_sayisi = sum(1 for s in summaries if s.get("_fallback_firma", False))
    fatura_bos_sayisi = sum(1 for s in summaries if not s.get("fatura_no", "").strip())
    hesap_bos_sayisi = sum(1 for s in summaries if not str(s.get("hesap_kodu", "")).strip())
    tevkifat_sayisi = sum(1 for s in summaries if s.get("_tevkifatli"))
    cok_belgeli = [s for s in summaries if s.get("_cok_belgeli")]

    logger.info("Belge satırı bulunan fiş sayısı: %s", belge_bulunan)
    logger.info("Belge satırı bulunamayan fiş sayısı: %s", belge_bulunamayan)
    logger.info("Firma fallback ile doldurulan fiş sayısı: %s", fallback_firma_sayisi)
    logger.info("Fatura_no boş kalan fiş sayısı: %s", fatura_bos_sayisi)
    logger.info("Hesap_kodu boş kalan fiş sayısı: %s", hesap_bos_sayisi)
    logger.info("Tevkifatlı (360+KDV) fiş sayısı: %s", tevkifat_sayisi)
    logger.info("Çok belgeli fiş sayısı: %s", len(cok_belgeli))
    for s in cok_belgeli[:20]:
        logger.warning(
            "Çok belgeli fiş uyarısı: Fiş %s | seçilen_fatura=%s | benzersiz_fatura=%s",
            s["sira_no"],
            s.get("fatura_no", ""),
            s.get("_benzersiz_fatura_sayisi", 0),
        )

    if summaries:
        first = summaries[0]
        logger.info("Fiş 1 KDV kaynak satır sayısı: %s", first.get("_kdv_source_count", 0))
        logger.info("Fiş 1 toplam kaynağı = %s", first.get("_toplam_source", "YOK"))
        logger.info("Fiş 1 belge tipi = %s", first.get("_belge_tipi", "YOK"))
        logger.info("Fiş 1 kısa tarih = %s", first.get("_kisa_tarih", ""))

    for summary in summaries[:10]:
        logger.info(
            "Fiş %s -> belge_tipi=%s, fatura_no=%s, firma=%s, belge_satiri_bulundu=%s, fallback_firma=%s, hesap_kodu=%s, tarih=%s, kdv=%.2f, mal_hizmet=%.2f, toplam=%.2f, etiket=%s, kdv_detay_kodlari=%s",
            summary["sira_no"],
            summary.get("_belge_tipi", "YOK"),
            summary.get("fatura_no", ""),
            summary["firma"],
            summary.get("_belge_satiri_bulundu", False),
            summary.get("_fallback_firma", False),
            summary["hesap_kodu"],
            summary["tarih"],
            summary["kdv"],
            summary["mal_hizmet"],
            summary["toplam"],
            summary["etiket"],
            summary.get("_kdv_detay_kodlari", []),
        )

    return summaries
