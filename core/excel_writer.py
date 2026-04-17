import os

import pandas as pd
from openpyxl.utils import get_column_letter


def write_summary_to_excel(summary_list: list[dict], output_dir: str, logger) -> str:
    logger.info("Excel yazma başladı")

    file_path = os.path.join(output_dir, "yevmiye_ozet.xlsx")

    df = pd.DataFrame(summary_list).copy()
    if df.empty:
        df = pd.DataFrame(
            columns=[
                "sira_no",
                "hesap_kodu",
                "tarih",
                "fatura_no",
                "firma",
                "kdv",
                "mal_hizmet",
                "toplam",
                "etiket",
            ]
        )

    selected_columns = [
        "sira_no",
        "hesap_kodu",
        "tarih",
        "fatura_no",
        "firma",
        "kdv",
        "mal_hizmet",
        "toplam",
        "etiket",
        "kdv_oran",
    ]
    df = df.reindex(columns=selected_columns)
    if "kdv" in df.columns and "mal_hizmet" in df.columns:
        df["kdv_oran"] = (
            pd.to_numeric(df["kdv"], errors="coerce")
            .div(pd.to_numeric(df["mal_hizmet"], errors="coerce").replace(0, pd.NA))
            .fillna(0)
            .round(2)
        )
    if not df.empty:
        # Hesap kodu 120 olan kayıtlar en altta yer alır.
        df["_hesap_120_last"] = (df["hesap_kodu"].astype(str).str.strip() == "120").astype(int)
        df = df.sort_values(by="_hesap_120_last", ascending=True, kind="mergesort").drop(
            columns="_hesap_120_last"
        )
        # Sıralama sonrası sıra numarası 1..N olarak yeniden yazılır.
        df = df.reset_index(drop=True)
        df["sira_no"] = df.index + 1

    rename_map = {
        "sira_no": "Sıra No",
        "hesap_kodu": "Hesap Kodu",
        "tarih": "Tarih",
        "fatura_no": "Fatura No",
        "firma": "Firma",
        "kdv": "KDV",
        "mal_hizmet": "Mal/Hizmet",
        "toplam": "Toplam",
        "etiket": "Etiket",
        "kdv_oran": "KDV Oran",
    }
    export_df = df.rename(columns=rename_map)

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Ozet")
        worksheet = writer.sheets["Ozet"]

        for col_idx, column_name in enumerate(export_df.columns, start=1):
            values = export_df[column_name].astype(str).tolist()
            max_length = max([len(column_name), *[len(v) for v in values]] if values else [len(column_name)])
            worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max_length + 2, 80)

    logger.info("Excel yazma tamamlandı: %s", file_path)
    return file_path
