import logging
import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk

from core.reconciliation import ReconciliationResult, run_reconciliation

EXCEL_FILETYPES = [
    ("Excel dosyaları", "*.xlsx *.xlsm *.xls"),
    ("Tüm dosyalar", "*.*"),
]


class MainWindow:
    def __init__(self, root: tk.Tk, logger: logging.Logger) -> None:
        self.root = root
        self.logger = logger
        self._busy = False

        self.selected_yevmiye_file = tk.StringVar(value="Henüz yevmiye dosyası seçilmedi.")
        self.selected_gider_file = tk.StringVar(value="Henüz gider dosyası seçilmedi.")
        self.selected_gelir_file = tk.StringVar(value="Henüz gelir dosyası seçilmedi.")
        self.selected_output_dir = tk.StringVar(value="Henüz klasör seçilmedi.")
        self.status_message = tk.StringVar(value="Durum: Hazır")

        self._configure_window()
        self._create_widgets()

    def _configure_window(self) -> None:
        self.root.title("Yevmiye Kontrol")
        self.root.geometry("760x450")
        self.root.minsize(760, 450)
        self.root.columnconfigure(0, weight=1)

    def _create_widgets(self) -> None:
        container = ttk.Frame(self.root, padding=16)
        container.grid(row=0, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)

        yevmiye_button = ttk.Button(
            container,
            text="Yevmiye Excel Seç",
            command=self.select_yevmiye_file,
        )
        yevmiye_button.grid(row=0, column=0, sticky="w")

        yevmiye_label = ttk.Label(
            container,
            textvariable=self.selected_yevmiye_file,
            wraplength=700,
        )
        yevmiye_label.grid(row=1, column=0, sticky="w", pady=(8, 12))

        gider_button = ttk.Button(
            container,
            text="Giderler Excel Seç",
            command=self.select_gider_file,
        )
        gider_button.grid(row=2, column=0, sticky="w")

        gider_label = ttk.Label(
            container,
            textvariable=self.selected_gider_file,
            wraplength=700,
        )
        gider_label.grid(row=3, column=0, sticky="w", pady=(8, 12))

        gelir_button = ttk.Button(
            container,
            text="Gelirler Excel Seç (opsiyonel)",
            command=self.select_gelir_file,
        )
        gelir_button.grid(row=4, column=0, sticky="w")

        gelir_label = ttk.Label(
            container,
            textvariable=self.selected_gelir_file,
            wraplength=700,
        )
        gelir_label.grid(row=5, column=0, sticky="w", pady=(8, 16))

        output_button = ttk.Button(
            container,
            text="Çıktı Klasörü Seç",
            command=self.select_output_directory,
        )
        output_button.grid(row=6, column=0, sticky="w")

        output_label = ttk.Label(
            container,
            textvariable=self.selected_output_dir,
            wraplength=700,
        )
        output_label.grid(row=7, column=0, sticky="w", pady=(8, 16))

        self.start_button = ttk.Button(container, text="Başlat", command=self.start_process)
        self.start_button.grid(row=8, column=0, sticky="w")

        self.status_label = ttk.Label(container, textvariable=self.status_message, foreground="#222222")
        self.status_label.grid(row=9, column=0, sticky="w", pady=(24, 0))

    def select_yevmiye_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Yevmiye Excel Seç",
            filetypes=EXCEL_FILETYPES,
        )
        if not file_path:
            self.status_message.set("Durum: Yevmiye dosya seçimi iptal edildi.")
            self.logger.info("Yevmiye dosya seçimi iptal edildi.")
            return

        normalized_path = str(Path(file_path))
        self.selected_yevmiye_file.set(normalized_path)
        self.status_message.set("Durum: Yevmiye dosyası seçildi.")
        self.status_label.configure(foreground="#222222")
        self.logger.info("Yevmiye dosyası seçildi: %s", normalized_path)

    def select_gider_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Giderler Excel Seç",
            filetypes=EXCEL_FILETYPES,
        )
        if not file_path:
            self.status_message.set("Durum: Gider dosya seçimi iptal edildi.")
            self.logger.info("Gider dosya seçimi iptal edildi.")
            return

        normalized_path = str(Path(file_path))
        self.selected_gider_file.set(normalized_path)
        self.status_message.set("Durum: Gider dosyası seçildi.")
        self.status_label.configure(foreground="#222222")
        self.logger.info("Gider dosyası seçildi: %s", normalized_path)

    def select_gelir_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Gelirler Excel Seç",
            filetypes=EXCEL_FILETYPES,
        )
        if not file_path:
            self.status_message.set("Durum: Gelir dosya seçimi iptal edildi.")
            self.logger.info("Gelir dosya seçimi iptal edildi.")
            return

        normalized_path = str(Path(file_path))
        self.selected_gelir_file.set(normalized_path)
        self.status_message.set("Durum: Gelir dosyası seçildi.")
        self.status_label.configure(foreground="#222222")
        self.logger.info("Gelir dosyası seçildi: %s", normalized_path)

    def select_output_directory(self) -> None:
        dir_path = filedialog.askdirectory(title="Çıktı Klasörü Seç")
        if not dir_path:
            self.status_message.set("Durum: Çıktı klasörü seçimi iptal edildi.")
            self.logger.info("Çıktı klasörü seçimi iptal edildi.")
            return

        normalized_path = str(Path(dir_path))
        self.selected_output_dir.set(normalized_path)
        self.status_message.set("Durum: Çıktı klasörü seçildi.")
        self.status_label.configure(foreground="#222222")
        self.logger.info("Çıktı klasörü seçildi: %s", normalized_path)

    def start_process(self) -> None:
        self.logger.info("Başlat butonuna tıklandı.")
        if self._busy:
            return

        selected_yevmiye = self.selected_yevmiye_file.get().strip()
        if not selected_yevmiye or selected_yevmiye == "Henüz yevmiye dosyası seçilmedi.":
            self.status_message.set("Durum: Lütfen önce yevmiye dosyasını seçin.")
            self.status_label.configure(foreground="#B00020")
            self.logger.warning("Başlat iptal: Yevmiye dosyası seçilmedi.")
            return

        selected_gider = self.selected_gider_file.get().strip()
        if not selected_gider or selected_gider == "Henüz gider dosyası seçilmedi.":
            self.status_message.set("Durum: Lütfen önce gider dosyasını seçin.")
            self.status_label.configure(foreground="#B00020")
            self.logger.warning("Başlat iptal: Gider dosyası seçilmedi.")
            return

        selected_gelir = self.selected_gelir_file.get().strip()
        if not selected_gelir or selected_gelir == "Henüz gelir dosyası seçilmedi.":
            selected_gelir = ""
        selected_output_dir = self.selected_output_dir.get().strip()
        if not selected_output_dir or selected_output_dir == "Henüz klasör seçilmedi.":
            self.status_message.set("Durum: Lütfen önce çıktı klasörü seçin.")
            self.status_label.configure(foreground="#B00020")
            self.logger.warning("Başlat iptal: Çıktı klasörü seçilmedi.")
            return

        self._busy = True
        self.start_button.configure(state="disabled")
        self.status_message.set("Durum: İşleniyor…")
        self.status_label.configure(foreground="#222222")

        thread = threading.Thread(
            target=self._run_reconciliation_worker,
            args=(
                selected_yevmiye,
                selected_gider,
                selected_gelir or None,
                selected_output_dir,
            ),
            daemon=True,
        )
        thread.start()

    def _run_reconciliation_worker(
        self,
        yevmiye: str,
        gider: str,
        gelir: str | None,
        output_dir: str,
    ) -> None:
        try:
            result = run_reconciliation(
                yevmiye_file_path=yevmiye,
                gider_file_path=gider,
                gelir_file_path=gelir,
                output_dir=output_dir,
                logger=self.logger,
            )
            self.root.after(0, lambda: self._on_success(result))
        except Exception as exc:
            self.logger.exception("Karşılaştırma sırasında hata oluştu: %s", exc)
            self.root.after(0, lambda: self._on_error(str(exc)))

    def _on_success(self, result: ReconciliationResult) -> None:
        self._busy = False
        self.start_button.configure(state="normal")
        self.status_message.set(
            f"Durum: ✓ TAM={result.tam_uyumlu} | TEVKIFAT={result.tevkifat_uyumlu} | "
            f"ISTISNA={result.istisna_uyumlu} | FARK={result.fark_var} | "
            f"YOK(y)={result.eslesme_yok_yevmiye} YOK(x)={result.eslesme_yok_muhasebe} | "
            f"uyum %{result.birebir_orani}"
        )
        self.status_label.configure(foreground="#1E7A1E")
        self.logger.info("Karşılaştırma tamamlandı: %s", result.output_dir)
        try:
            os.startfile(result.output_dir)
        except OSError as exc:
            self.logger.warning("Çıktı klasörü açılamadı: %s", exc)

    def _on_error(self, _message: str) -> None:
        self._busy = False
        self.start_button.configure(state="normal")
        self.status_message.set("Durum: Hata oluştu. Logları kontrol edin.")
        self.status_label.configure(foreground="#B00020")
