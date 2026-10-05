"""Application composition root for the KML/KMZ parcel review tool."""

from __future__ import annotations

from pathlib import Path
import sys

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QIntValidator
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QComboBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from kml_engine import KmlDocument, Parcel, STATUS_LABELS
from parcel_editor import ParcelEditor
from ui_theme import APP_STYLESHEET, STATUS_COLORS


class SummaryCard(QFrame):
    def __init__(self, label: str, value: int, color: str):
        super().__init__()
        self.setObjectName("card")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        title = QLabel(label)
        title.setObjectName("muted")
        number = QLabel(str(value))
        number.setStyleSheet(
            f"color:{color};font-size:26px;font-weight:700;background:transparent;"
        )
        layout.addWidget(title)
        layout.addWidget(number)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.document: KmlDocument | None = None
        self.current: Parcel | None = None
        self.setWindowTitle("مراجع بيانات الحيازة")
        self.resize(1480, 900)
        self._build_ui()

    def _build_ui(self) -> None:
        root = QWidget()
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(22, 18, 22, 16)
        root_layout.setSpacing(14)
        self.setCentralWidget(root)
        root_layout.addWidget(self._build_header())
        self.cards = QHBoxLayout()
        self.cards.setSpacing(10)
        root_layout.addLayout(self.cards)
        root_layout.addWidget(self._build_toolbar())
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self._build_results_panel())
        self.editor = ParcelEditor()
        self.editor.save_requested.connect(self.save_manual)
        self.editor.skip_requested.connect(self.skip_missing_fields)
        splitter.addWidget(self.editor)
        splitter.setSizes([720, 760])
        root_layout.addWidget(splitter, 1)
        self.status = self.statusBar()
        self.status.showMessage("ارفع ملف KML أو KMZ للبدء.")
        self._set_controls_enabled(False)

    def _build_header(self) -> QFrame:
        header = QFrame()
        header.setObjectName("header")
        layout = QHBoxLayout(header)
        layout.setContentsMargins(20, 16, 20, 16)
        words = QVBoxLayout()
        title = QLabel("مراجع بيانات الحيازة")
        title.setObjectName("headerTitle")
        subtitle = QLabel("فحص وإصلاح KML / KMZ قبل بناء ملف Excel")
        subtitle.setObjectName("headerSubtitle")
        words.addWidget(title)
        words.addWidget(subtitle)
        layout.addLayout(words)
        layout.addStretch()
        self.file_label = QLabel("لم يتم اختيار ملف")
        self.file_label.setObjectName("fileChip")
        load = QPushButton("رفع ملف")
        load.setObjectName("primary")
        load.clicked.connect(self.load_file)
        layout.addWidget(self.file_label)
        layout.addWidget(load)
        return header

    def _build_toolbar(self) -> QFrame:
        frame = QFrame()
        frame.setObjectName("toolbar")
        grid = QGridLayout(frame)
        grid.setContentsMargins(14, 12, 14, 12)
        grid.setHorizontalSpacing(9)
        grid.setVerticalSpacing(9)
        self.filter = QComboBox()
        self.filter.addItem("كل الحالات", "all")
        for key, label in STATUS_LABELS.items():
            self.filter.addItem(label, key)
        self.filter.currentIndexChanged.connect(self.refresh_table)
        self.search = QLineEdit()
        self.search.setPlaceholderText("ابحث بالـ ID أو اسم الحائز أو سبب المشكلة")
        self.search.textChanged.connect(self.refresh_table)
        self.assoc_name = QLineEdit()
        self.assoc_name.setPlaceholderText("اسم الجمعية")
        self.assoc_code = QLineEdit()
        self.assoc_code.setPlaceholderText("كود الجمعية")
        association = QPushButton("تطبيق بيانات الجمعية")
        association.clicked.connect(self.apply_association)
        self.fix = QPushButton("تطبيق التصحيحات الآمنة")
        self.fix.setObjectName("warning")
        self.fix.clicked.connect(self.apply_fixes)
        self.kml_export = QPushButton("تصدير KML للمراجعة")
        self.kml_export.clicked.connect(self.export_kml)
        self.excel_export = QPushButton("بناء Excel")
        self.excel_export.setObjectName("primary")
        self.excel_export.clicked.connect(self.export_excel)
        self.growth_inputs = {}
        for stage, placeholder in [
            ("مرحلة النمو الخضري", "نسبة النمو الخضري %"),
            ("مرحلة الإنبات", "نسبة الإنبات %"),
            ("مرحلة الإزهار والإثمار", "نسبة الإزهار/الإثمار %"),
        ]:
            field = QLineEdit(); field.setPlaceholderText(placeholder); field.setMaxLength(3)
            field.setValidator(QIntValidator(0, 100, self)); field.textChanged.connect(self.update_growth_total)
            self.growth_inputs[stage] = field
        self.growth_total = QLabel("المجموع: 0٪ — أدخل 100٪")
        self.growth_total.setObjectName("muted")
        self.growth_distribution = QPushButton("توزيع مراحل النمو")
        self.growth_distribution.setObjectName("success")
        self.growth_distribution.clicked.connect(self.apply_growth_distribution)
        grid.addWidget(QLabel("التصفية"), 0, 0)
        grid.addWidget(self.filter, 0, 1)
        grid.addWidget(self.search, 0, 2, 1, 5)
        grid.addWidget(QLabel("بيانات جمعية موحدة"), 1, 0)
        grid.addWidget(self.assoc_name, 1, 1)
        grid.addWidget(self.assoc_code, 1, 2)
        grid.addWidget(association, 1, 3)
        grid.addWidget(self.fix, 1, 4)
        grid.addWidget(self.kml_export, 1, 5)
        grid.addWidget(self.excel_export, 1, 6)
        grid.addWidget(QLabel("مراحل النمو"), 2, 0)
        grid.addWidget(self.growth_inputs["مرحلة النمو الخضري"], 2, 1)
        grid.addWidget(self.growth_inputs["مرحلة الإنبات"], 2, 2)
        grid.addWidget(self.growth_inputs["مرحلة الإزهار والإثمار"], 2, 3)
        grid.addWidget(self.growth_distribution, 2, 4)
        grid.addWidget(self.growth_total, 2, 5, 1, 2)
        return frame

    def _build_results_panel(self) -> QFrame:
        panel = QFrame()
        panel.setObjectName("panel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(14, 14, 14, 14)
        title_row = QHBoxLayout()
        title = QLabel("نتائج الفحص")
        title.setObjectName("sectionTitle")
        self.result_note = QLabel("اختر صفًا لفتح محرر القطعة.")
        self.result_note.setObjectName("muted")
        title_row.addWidget(title)
        title_row.addStretch()
        title_row.addWidget(self.result_note)
        layout.addLayout(title_row)
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(
            [
                "#",
                "الحالة",
                "سبب المراجعة",
                "عنوان القطعة",
                "ID المستخرج",
                "الحقول الناقصة",
            ]
        )
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.itemSelectionChanged.connect(self.select_row)
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table)
        return panel

    def _set_controls_enabled(self, enabled: bool) -> None:
        controls = [
            self.filter,
            self.search,
            self.assoc_name,
            self.assoc_code,
            self.fix,
            self.kml_export,
            self.excel_export,
            *self.growth_inputs.values(),
        ]
        for control in controls:
            control.setEnabled(enabled)
        self.update_growth_total()

    def update_growth_total(self) -> None:
        values = [field.text().strip() for field in self.growth_inputs.values()]
        total = sum(int(value) for value in values if value.isdigit())
        complete = len(values) == 3 and all(value.isdigit() for value in values)
        valid = complete and total == 100
        if valid:
            self.growth_total.setText("المجموع: 100٪ — جاهز للتوزيع")
            self.growth_total.setStyleSheet("color:#166534;font-weight:700;background:transparent;")
        else:
            self.growth_total.setText(f"المجموع: {total}٪ — يجب أن يساوي 100٪")
            self.growth_total.setStyleSheet("color:#B42318;font-weight:700;background:transparent;")
        self.growth_distribution.setEnabled(bool(self.document) and valid)

    def load_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "اختيار ملف", "", "KML/KMZ (*.kml *.kmz)"
        )
        if not path:
            return
        try:
            self.document = KmlDocument.load(path)
            self.file_label.setText(Path(path).name)
            self._set_controls_enabled(True)
            self.refresh_all()
            self.status.showMessage(f"تم تحليل {len(self.document.parcels)} قطعة.")
        except Exception as exc:
            QMessageBox.critical(self, "تعذر قراءة الملف", str(exc))

    def refresh_all(self) -> None:
        self.refresh_cards()
        self.refresh_table()
        self.excel_export.setEnabled(
            bool(self.document and not self.document.summary()["blocking"])
        )

    def refresh_cards(self) -> None:
        while self.cards.count():
            item = self.cards.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        if not self.document:
            return
        summary = self.document.summary()
        cards = [
            ("إجمالي القطع", summary["total"], "#0E7490"),
            ("سليمة", summary["valid"], "#166534"),
            ("غير محيزة", summary["unallocated"], "#166534"),
            ("تم إصلاحها", summary["auto_fixed"], "#2563EB"),
            ("تم تجاوز النواقص", summary["skipped"], "#A16207"),
            (
                "تحتاج مراجعة",
                summary["duplicate_identical"] + summary["title_mismatch"],
                "#A16207",
            ),
            ("مشاكل مانعة", summary["blocking"], "#B42318"),
        ]
        for label, value, color in cards:
            self.cards.addWidget(SummaryCard(label, value, color))

    def filtered_parcels(self) -> list[Parcel]:
        if not self.document:
            return []
        status, query = self.filter.currentData(), self.search.text().strip().casefold()
        parcels = []
        for parcel in self.document.parcels:
            searchable = " ".join(
                [parcel.title, parcel.id, parcel.reason, *parcel.data.values()]
            ).casefold()
            if (status == "all" or parcel.status == status) and (
                not query or query in searchable
            ):
                parcels.append(parcel)
        return parcels

    def refresh_table(self) -> None:
        parcels = self.filtered_parcels()
        self.table.setRowCount(len(parcels))
        self.result_note.setText(f"{len(parcels)} قطعة ظاهرة")
        for row, parcel in enumerate(parcels):
            values = [
                str(parcel.index),
                STATUS_LABELS.get(parcel.status, parcel.status),
                parcel.reason,
                parcel.title,
                parcel.id,
                "، ".join(parcel.missing),
            ]
            for col, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setData(Qt.UserRole, parcel)
                if col == 1:
                    item.setForeground(
                        QColor(STATUS_COLORS.get(parcel.status, "#172033"))
                    )
                self.table.setItem(row, col, item)
        self.table.resizeColumnsToContents()

    def select_row(self) -> None:
        selected = self.table.selectedItems()
        if selected:
            self.current = self.table.item(selected[0].row(), 0).data(Qt.UserRole)
            self.editor.set_parcel(self.current)

    def apply_fixes(self) -> None:
        if not self.document:
            return
        count = self.document.apply_safe_fixes()
        self.refresh_all()
        self.status.showMessage(f"تم تطبيق {count} تصحيحًا آمنًا.")

    def apply_association(self) -> None:
        if self.document:
            self.document.apply_association(
                self.assoc_name.text(), self.assoc_code.text()
            )
            self.refresh_all()
            self.status.showMessage("تم تطبيق بيانات الجمعية على القطع.")

    def apply_growth_distribution(self) -> None:
        if not self.document:
            return
        try:
            count = self.document.distribute_growth_stages(
                {stage: field.text() for stage, field in self.growth_inputs.items()}
            )
        except ValueError as exc:
            QMessageBox.warning(self, "تعذر توزيع مراحل النمو", str(exc))
            return
        self.refresh_all()
        self.status.showMessage(f"تم توزيع مراحل النمو على {count} قطعة زراعية.")

    def save_manual(self, values: dict[str, str], approved: bool) -> None:
        if not self.document or not self.current:
            return
        self.document.update_parcel(self.current, values, approved)
        self.refresh_all()
        self.status.showMessage("تم حفظ تعديل القطعة.")

    def skip_missing_fields(self, values: dict[str, str]) -> None:
        if not self.document or not self.current:
            return
        self.document.update_parcel(self.current, values, skip_missing=True)
        self.refresh_all()
        self.editor.set_parcel(self.current)
        self.status.showMessage(
            "تم تجاوز النواقص؛ لم تعد القطعة مشكلة مانعة ويمكن تصديرها بالقيم الفارغة."
        )

    def export_kml(self) -> None:
        if not self.document:
            return
        suggested = self.document.source.with_name(
            self.document.source.stem + "_مراجعة.kml"
        )
        path, _ = QFileDialog.getSaveFileName(
            self, "حفظ ملف المراجعة", str(suggested), "KML (*.kml);;KMZ (*.kmz)"
        )
        if not path:
            return
        try:
            self.document.export_kml(path)
            self.status.showMessage("تم حفظ ملف KML للمراجعة.")
        except Exception as exc:
            QMessageBox.critical(self, "تعذر التصدير", str(exc))

    def export_excel(self) -> None:
        if not self.document:
            return
        path, _ = QFileDialog.getSaveFileName(
            self,
            "حفظ Excel",
            str(self.document.source.with_suffix(".xlsx")),
            "Excel (*.xlsx)",
        )
        if not path:
            return
        try:
            self.document.export_excel(path)
            self.status.showMessage("تم بناء ملف Excel بنجاح.")
        except Exception as exc:
            QMessageBox.critical(self, "تعذر بناء Excel", str(exc))


def main() -> None:
    app = QApplication(sys.argv)
    app.setLayoutDirection(Qt.RightToLeft)
    app.setStyle("Fusion")
    app.setStyleSheet(APP_STYLESHEET)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
