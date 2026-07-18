"""Eşleşme / parse yardımcılarının birim testleri."""

from core.reconciliation import (
    STATUS_FARK,
    STATUS_TAM,
    STATUS_TEVKIFAT,
    _evaluate_match,
    _is_prefix_match,
    _normalize_fatura_no,
)
from core.yevmiye_parser import parse_fis_header


def test_normalize_fatura_strips_zeros_and_separators() -> None:
    assert _normalize_fatura_no("ABC-00123") == "ABC123"
    assert _normalize_fatura_no("00045") == "45"
    assert _normalize_fatura_no("") is None
    assert _normalize_fatura_no(None) is None


def test_prefix_match_rejects_short_codes() -> None:
    assert _is_prefix_match("123456", "1234") is True
    assert _is_prefix_match("12", "123") is False
    assert _is_prefix_match("ABCD", "ABCD") is True


def test_fis_header_parse() -> None:
    h = parse_fis_header("03750-----03752-----MAHSUP-----01/11/2025")
    assert h is not None
    assert h["fis_no_bas"] == "03750"
    assert h["fis_no_bit"] == "03752"
    assert h["fis_tipi"] == "MAHSUP"
    assert h["tarih"] == "01.11.2025"
    assert parse_fis_header("HESAP KODU") is None


def test_tam_requires_exact_fatura() -> None:
    y_row = {
        "Fatura No": "ABC123",
        "Firma": "ACME",
        "Toplam": 100.0,
        "KDV": 18.0,
        "KDV'siz": 82.0,
        "Etiket": "GENEL",
        "_mal_hizmet": 82.0,
    }
    m_exact = {
        "Fatura No": "ABC-0123",
        "Firma": "ACME",
        "Toplam": 100.0,
        "KDV": 18.0,
        "KDV'siz": 82.0,
        "Etiket": "GENEL",
        "_kaynak": "gider",
    }
    m_prefix = {
        "Fatura No": "ABC12399",
        "Firma": "ACME",
        "Toplam": 100.0,
        "KDV": 18.0,
        "KDV'siz": 82.0,
        "Etiket": "GENEL",
        "_kaynak": "gider",
    }

    status_ok, note_ok, *_ = _evaluate_match(y_row, m_exact, is_last_chance=False)
    assert status_ok == STATUS_TAM
    assert "uyumlu" in note_ok.lower()

    status_soft, note_soft, *_ = _evaluate_match(y_row, m_prefix, is_last_chance=False)
    assert status_soft == STATUS_FARK
    assert "fatura" in note_soft.lower()


def test_tevkifat_match() -> None:
    y_row = {
        "Fatura No": "31814323",
        "Firma": "Akamai",
        "Toplam": 6850.15,
        "KDV": 1141.69,
        "KDV'siz": 5708.46,
        "Etiket": "ENDÜSTRİYEL",
        "_mal_hizmet": 5708.46,
        "_tevkifatli": True,
        "_has_360": True,
    }
    m_row = {
        "Fatura No": "31814323",
        "Firma": "Akamai Technologies",
        "Toplam": 5708.46,
        "KDV": 0.0,
        "KDV'siz": 5708.46,
        "Etiket": "ENDÜSTRİYEL",
        "_kaynak": "gider",
        "_tevkifat_aday": True,
    }
    status, note, *_ = _evaluate_match(y_row, m_row, is_last_chance=False)
    assert status == STATUS_TEVKIFAT
    assert "tevkifat" in note.lower()


def test_last_chance_never_tam() -> None:
    y_row = {
        "Fatura No": "99991234",
        "Firma": "X",
        "Toplam": 50.0,
        "KDV": 0.0,
        "Etiket": "",
        "_mal_hizmet": 50.0,
    }
    m_row = {
        "Fatura No": "88881234",
        "Firma": "X",
        "Toplam": 50.0,
        "KDV": 0.0,
        "Etiket": "",
        "_kaynak": "gider",
    }
    status, note, *_ = _evaluate_match(y_row, m_row, is_last_chance=True)
    assert status == STATUS_FARK
    assert "son hane" in note.lower()
