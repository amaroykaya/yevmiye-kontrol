"""
Birleşik kabuk için Yevmiye kontrol sayfası (PySide6).
İş mantığı: `core.reconciliation.run_reconciliation` (Tk arayüzü kullanılmaz).
"""

from __future__ import annotations

import logging
import os

from PySide6.QtCore import QObject, QThread, QUrl, Qt, Signal, Slot
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from core.logger_setup import setup_logger
from core.reconciliation import ReconciliationResult, run_reconciliation


class _ReconciliationWorker(QObject):
    finished_ok = Signal(object)
    finished_err = Signal(str)

    def __init__(
        self,
        yevmiye: str,
        gider: str,
        gelir: str | None,
        output_dir: str,
        logger: logging.Logger,
    ) -> None:
        super().__init__()
        self._yevmiye = yevmiye
        self._gider = gider
        self._gelir = gelir
        self._output_dir = output_dir
        self._logger = logger

    @Slot()
    def run(self) -> None:
        try:
            result = run_reconciliation(
                yevmiye_file_path=self._yevmiye,
                gider_file_path=self._gider,
                gelir_file_path=self._gelir,
                output_dir=self._output_dir,
                logger=self._logger,
            )
            self.finished_ok.emit(result)
        except Exception as exc:  # noqa: BLE001
            self._logger.exception("Karşılaştırma sırasında hata: %s", exc)
            self.finished_err.emit(str(exc))


class YevmiyePage(QWidget):
    """Dosya seçimleri, Başlat ve durum; uzun iş `QThread` üzerinde."""

    def __init__(self) -> None:
        super().__init__()
        self.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Expanding,
        )
        self._logger = setup_logger()
        self._logger.info("Yevmiye kontrol (PySide6) sayfası açıldı.")

        self._yevmiye_path = ""
        self._gider_path = ""
        self._gelir_path = ""
        self._output_dir = ""

        self._thread: QThread | None = None
        self._worker: _ReconciliationWorker | None = None

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(12)

        paths = QGroupBox("Dosyalar")
        form = QFormLayout(paths)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(10)

        self._yevmiye_edit = QLineEdit()
        self._yevmiye_edit.setReadOnly(True)
        self._yevmiye_edit.setPlaceholderText("Henüz seçilmedi")
        form.addRow("Yevmiye Excel:", self._file_row(self._yevmiye_edit, self._pick_yevmiye))

        self._gider_edit = QLineEdit()
        self._gider_edit.setReadOnly(True)
        self._gider_edit.setPlaceholderText("Henüz seçilmedi")
        form.addRow("Giderler Excel:", self._file_row(self._gider_edit, self._pick_gider))

        self._gelir_edit = QLineEdit()
        self._gelir_edit.setReadOnly(True)
        self._gelir_edit.setPlaceholderText("Opsiyonel — seçilmezse yalnız gider")
        form.addRow("Gelirler Excel:", self._file_row(self._gelir_edit, self._pick_gelir))

        self._output_edit = QLineEdit()
        self._output_edit.setReadOnly(True)
        self._output_edit.setPlaceholderText("Henüz seçilmedi")
        form.addRow("Çıktı klasörü:", self._file_row(self._output_edit, self._pick_output))

        root.addWidget(paths)

        self._start_btn = QPushButton("Başlat")
        self._start_btn.clicked.connect(self._on_start)
        root.addWidget(self._start_btn, alignment=Qt.AlignmentFlag.AlignLeft)

        self._status = QLabel("Durum: Hazır")
        self._status.setWordWrap(True)
        self._status.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        root.addWidget(self._status)
        root.addStretch(1)

    def _file_row(self, edit: QLineEdit, pick_slot) -> QWidget:
        w = QWidget()
        h = QHBoxLayout(w)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(8)
        btn = QPushButton("Seç…")
        btn.clicked.connect(pick_slot)
        h.addWidget(btn)
        h.addWidget(edit, stretch=1)
        return w

    def _pick_yevmiye(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Yevmiye Excel Seç", "", "Excel (*.xlsx *.xlsm *.xls)"
        )
        if path:
            self._yevmiye_path = path
            self._yevmiye_edit.setText(path)
            self._set_status("Durum: Yevmiye dosyası seçildi.", neutral=True)
            self._logger.info("Yevmiye dosyası seçildi: %s", path)

    def _pick_gider(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Giderler Excel Seç", "", "Excel (*.xlsx *.xlsm *.xls)"
        )
        if path:
            self._gider_path = path
            self._gider_edit.setText(path)
            self._set_status("Durum: Gider dosyası seçildi.", neutral=True)
            self._logger.info("Gider dosyası seçildi: %s", path)

    def _pick_gelir(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Gelirler Excel Seç", "", "Excel (*.xlsx *.xlsm *.xls)"
        )
        if path:
            self._gelir_path = path
            self._gelir_edit.setText(path)
            self._set_status("Durum: Gelir dosyası seçildi.", neutral=True)
            self._logger.info("Gelir dosyası seçildi: %s", path)

    def _pick_output(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Çıktı Klasörü Seç")
        if path:
            self._output_dir = path
            self._output_edit.setText(path)
            self._set_status("Durum: Çıktı klasörü seçildi.", neutral=True)
            self._logger.info("Çıktı klasörü seçildi: %s", path)

    def _set_status(
        self,
        text: str,
        *,
        neutral: bool = False,
        ok: bool = False,
        err: bool = False,
    ) -> None:
        self._status.setText(text)
        if err:
            self._status.setStyleSheet("color: #B00020;")
        elif ok:
            self._status.setStyleSheet("color: #1E7A1E;")
        elif neutral:
            self._status.setStyleSheet("color: #222222;")
        else:
            self._status.setStyleSheet("color: #222222;")

    def _on_start(self) -> None:
        self._logger.info("Başlat butonuna tıklandı.")
        if not self._yevmiye_path:
            self._set_status("Durum: Lütfen önce yevmiye dosyasını seçin.", err=True)
            self._logger.warning("Başlat iptal: Yevmiye dosyası seçilmedi.")
            return
        if not self._gider_path:
            self._set_status("Durum: Lütfen önce gider dosyasını seçin.", err=True)
            self._logger.warning("Başlat iptal: Gider dosyası seçilmedi.")
            return
        if not self._output_dir:
            self._set_status("Durum: Lütfen önce çıktı klasörünü seçin.", err=True)
            self._logger.warning("Başlat iptal: Çıktı klasörü seçilmedi.")
            return

        if self._thread is not None and self._thread.isRunning():
            return

        self._start_btn.setEnabled(False)
        self._set_status("Durum: İşleniyor…", neutral=True)

        self._thread = QThread(self)
        self._worker = _ReconciliationWorker(
            self._yevmiye_path,
            self._gider_path,
            self._gelir_path or None,
            self._output_dir,
            self._logger,
        )
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished_ok.connect(self._on_worker_ok)
        self._worker.finished_err.connect(self._on_worker_err)
        self._worker.finished_ok.connect(self._thread.quit)
        self._worker.finished_err.connect(self._thread.quit)
        self._thread.finished.connect(self._cleanup_thread)
        self._thread.start()

    @Slot(object)
    def _on_worker_ok(self, result: object) -> None:
        if not isinstance(result, ReconciliationResult):
            self._on_worker_err("Beklenmeyen sonuç tipi")
            return
        self._set_status(
            f"Durum: ✓ TAM={result.tam_uyumlu} | TEVKIFAT={result.tevkifat_uyumlu} | "
            f"ISTISNA={result.istisna_uyumlu} | FARK={result.fark_var} | "
            f"YOK(y)={result.eslesme_yok_yevmiye} YOK(x)={result.eslesme_yok_muhasebe} | "
            f"uyum %{result.birebir_orani}",
            ok=True,
        )
        self._logger.info("Karşılaştırma tamamlandı: %s", result.output_dir)
        self._start_btn.setEnabled(True)
        url = QUrl.fromLocalFile(result.output_dir)
        if url.isValid():
            QDesktopServices.openUrl(url)
        elif os.name == "nt":
            os.startfile(result.output_dir)  # type: ignore[attr-defined]

    @Slot(str)
    def _on_worker_err(self, msg: str) -> None:
        self._set_status("Durum: Hata oluştu. Logları kontrol edin.", err=True)
        self._start_btn.setEnabled(True)
        QMessageBox.critical(self, "Yevmiye kontrol", msg)

    @Slot()
    def _cleanup_thread(self) -> None:
        if self._worker is not None:
            self._worker.deleteLater()
            self._worker = None
        if self._thread is not None:
            self._thread.deleteLater()
            self._thread = None
