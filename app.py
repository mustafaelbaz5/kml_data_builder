from __future__ import annotations

from pathlib import Path
import sys

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QColor, QPalette
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
    QComboBox,
    QTextEdit,
)

from kml_engine import FIELDS, KmlDocument, Parcel, STATUS_LABELS


class Card(QFrame):
    def __init__(self, title: str, value: str, color: str):
        super().__init__()
        self.setObjectName("card")
        self.setStyleSheet(
            f"#card{{border:1px solid #d8e1ea;border-radius:12px;background:white;border-top:4px solid {color};}}"
        )
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 10, 14, 12)
        label = QLabel(title)
        label.setStyleSheet("color:#536273;font-size:12px;")
        number = QLabel(value)
        number.setStyleSheet("color:#122235;font-size:25px;font-weight:700;")
        layout.addWidget(label)
        layout.addWidget(number)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.document: KmlDocument | None = None
        self.current: Parcel | None = None
        self.editors: dict[str, QLineEdit] = {}
        self.setWindowTitle("مراجع بيانات الحيازة")
        self.resize(1450, 880)
        self._build_ui()

    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        root.setLayout(QVBoxLayout())
        root.layout().setContentsMargins(22, 18, 22, 18)
        root.layout().setSpacing(14)
        title_row = QHBoxLayout()
        title = QLabel("مراجع بيانات الحيازة")
        title.setStyleSheet("font-size:27px;font-weight:750;color:#102a43;")
        subtitle = QLabel("تحليل وإصلاح ملفات KML / KMZ قبل بناء Excel")
        subtitle.setStyleSheet("color:#62748a;")
        title_stack = QVBoxLayout()
        title_stack.addWidget(title)
        title_stack.addWidget(subtitle)
        title_row.addLayout(title_stack)
        title_row.addStretch()
        self.file_label = QLabel("لم يتم اختيار ملف")
        self.file_label.setStyleSheet("color:#536273;")
        load = QPushButton("رفع KML أو KMZ")
        load.clicked.connect(self.load_file)
        load.setObjectName("primary")
        title_row.addWidget(self.file_label)
        title_row.addWidget(load)
        root.layout().addLayout(title_row)

        self.cards = QHBoxLayout()
        root.layout().addLayout(self.cards)
        control = QFrame()
        control.setObjectName("control")
        control.setStyleSheet(
            "#control{background:#f6f9fc;border:1px solid #d8e1ea;border-radius:12px;}"
        )
        controls = QHBoxLayout(control)
        controls.setContentsMargins(14, 10, 14, 10)
        self.filter = QComboBox()
        self.filter.addItem("كل الحالات", "all")
        for key, label in STATUS_LABELS.items():
            self.filter.addItem(label, key)
        self.filter.currentIndexChanged.connect(self.refresh_table)
        self.search = QLineEdit()
        self.search.setPlaceholderText("ابحث بالـ ID أو الاسم أو السبب…")
        self.search.textChanged.connect(self.refresh_table)
        self.assoc_name = QLineEdit()
        self.assoc_name.setPlaceholderText("اسم الجمعية")
        self.assoc_code = QLineEdit()
        self.assoc_code.setPlaceholderText("كود الجمعية")
        association = QPushButton("تطبيق بيانات الجمعية")
        association.clicked.connect(self.apply_association)
        fix = QPushButton("تطبيق التصحيحات الآمنة")
        fix.clicked.connect(self.apply_fixes)
        self.kml_export = QPushButton("تصدير KML للمراجعة")
        self.kml_export.clicked.connect(self.export_kml)
        self.excel_export = QPushButton("بناء Excel")
        self.excel_export.clicked.connect(self.export_excel)
        self.excel_export.setObjectName("primary")
        controls.addWidget(QLabel("التصفية:"))
        controls.addWidget(self.filter)
        controls.addWidget(self.search, 1)
        controls.addWidget(self.assoc_name)
        controls.addWidget(self.assoc_code)
        controls.addWidget(association)
        controls.addWidget(fix)
        controls.addWidget(self.kml_export)
        controls.addWidget(self.excel_export)
        root.layout().addWidget(control)

        splitter = QSplitter(Qt.Horizontal)
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(
            ["#", "الحالة", "السبب", "ID في العنوان", "ID في الوصف", "الحقول الناقصة"]
        )
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.itemSelectionChanged.connect(self.select_row)
        self.table.horizontalHeader().setStretchLastSection(True)
        splitter.addWidget(self.table)
        detail = QWidget()
        detail_layout = QVBoxLayout(detail)
        detail_layout.setContentsMargins(14, 8, 4, 8)
        detail_layout.addWidget(
            QLabel(
                "مراجعة وتعديل القطعة",
                styleSheet="font-size:18px;font-weight:700;color:#102a43;",
            )
        )
        self.detail_note = QLabel("اختر قطعة من الجدول لرؤية التفاصيل.")
        self.detail_note.setWordWrap(True)
        self.detail_note.setStyleSheet("color:#62748a;")
        detail_layout.addWidget(self.detail_note)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        form_host = QWidget()
        self.form = QFormLayout(form_host)
        self.form.setLabelAlignment(Qt.AlignRight)
        scroll.setWidget(form_host)
        detail_layout.addWidget(scroll, 1)
        self.coordinates = QTextEdit()
        self.coordinates.setReadOnly(True)
        self.coordinates.setMaximumHeight(90)
        self.coordinates.setPlaceholderText(
            "الإحداثيات ستظهر هنا ويمكن نسخها إلى Google Earth."
        )
        detail_layout.addWidget(QLabel("الإحداثيات"))
        detail_layout.addWidget(self.coordinates)
        buttons = QHBoxLayout()
        save = QPushButton("حفظ التعديل")
        save.clicked.connect(self.save_manual)
        approve = QPushButton("حفظ واعتماد المراجعة")
        approve.clicked.connect(lambda: self.save_manual(True))
        buttons.addWidget(save)
        buttons.addWidget(approve)
        detail_layout.addLayout(buttons)
        splitter.addWidget(detail)
        splitter.setSizes([870, 530])
        root.layout().addWidget(splitter, 1)
        self.status = self.statusBar()
        self.status.showMessage("ارفع ملفًا للبدء.")
        self._set_disabled(True)

    def _set_disabled(self, disabled: bool):
        for widget in [
            self.filter,
            self.search,
            self.assoc_name,
            self.assoc_code,
            self.kml_export,
            self.excel_export,
        ]:
            widget.setDisabled(disabled)

    def load_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "اختيار ملف", "", "KML/KMZ (*.kml *.kmz)"
        )
        if not path:
            return
        try:
            self.document = KmlDocument.load(path)
            self.file_label.setText(Path(path).name)
            self._set_disabled(False)
            self.refresh_all()
            self.status.showMessage(f"تم تحليل {len(self.document.parcels)} قطعة.")
        except Exception as exc:
            QMessageBox.critical(self, "تعذر قراءة الملف", str(exc))

    def refresh_all(self):
        self.refresh_cards()
        self.refresh_table()
        self.excel_export.setEnabled(
            bool(self.document and not self.document.summary()["blocking"])
        )

    def refresh_cards(self):
        while self.cards.count():
            item = self.cards.takeAt(0)
            item.widget().deleteLater() if item.widget() else None
        if not self.document:
            return
        s = self.document.summary()
        definitions = [
            ("إجمالي القطع", s["total"], "#1f4e78"),
            ("سليمة", s["valid"], "#2f855a"),
            ("تم إصلاحها", s["auto_fixed"], "#2563eb"),
            ("تحتاج مراجعة", s["duplicate_identical"] + s["title_mismatch"], "#d69e2e"),
            ("مشاكل مانعة", s["blocking"], "#dc2626"),
        ]
        for label, value, color in definitions:
            self.cards.addWidget(Card(label, str(value), color))

    def filtered(self):
        if not self.document:
            return []
        wanted, needle = (
            self.filter.currentData(),
            self.search.text().strip().casefold(),
        )
        result = []
        for parcel in self.document.parcels:
            if wanted != "all" and parcel.status != wanted:
                continue
            haystack = " ".join(
                [parcel.title, parcel.id, parcel.reason, *parcel.data.values()]
            ).casefold()
            if needle and needle not in haystack:
                continue
            result.append(parcel)
        return result

    def refresh_table(self):
        rows = self.filtered()
        self.table.setRowCount(len(rows))
        for row, parcel in enumerate(rows):
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
                self.table.setItem(row, col, item)
        self.table.resizeColumnsToContents()

    def select_row(self):
        selected = self.table.selectedItems()
        if not selected:
            return
        parcel = self.table.item(selected[0].row(), 0).data(Qt.UserRole)
        self.current = parcel
        self.detail_note.setText(
            f"{STATUS_LABELS.get(parcel.status, parcel.status)} — {parcel.reason}"
        )
        self.coordinates.setPlainText(parcel.coordinates)
        while self.form.rowCount():
            self.form.removeRow(0)
        self.editors.clear()
        for key, label in FIELDS:
            editor = QLineEdit(parcel.data.get(key, ""))
            self.editors[key] = editor
            self.form.addRow(label + ":", editor)

    def apply_fixes(self):
        if not self.document:
            return
        count = self.document.apply_safe_fixes()
        self.refresh_all()
        self.status.showMessage(f"تم تطبيق {count} تصحيحًا آمنًا.")

    def apply_association(self):
        if not self.document:
            return
        self.document.apply_association(self.assoc_name.text(), self.assoc_code.text())
        self.refresh_all()
        self.status.showMessage("تم تطبيق بيانات الجمعية على كل القطع.")

    def save_manual(self, approved=False):
        if not self.document or not self.current:
            return
        values = {key: editor.text() for key, editor in self.editors.items()}
        self.document.update_parcel(self.current, values, approved)
        self.refresh_all()
        self.status.showMessage("تم حفظ تعديل القطعة.")

    def export_kml(self):
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

    def export_excel(self):
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


def main():
    app = QApplication(sys.argv)
    app.setLayoutDirection(Qt.RightToLeft)
    app.setStyle("Fusion")
    app.setStyleSheet(
        "QWidget{font-family:'Segoe UI';font-size:13px;} QPushButton{padding:8px 12px;border:1px solid #b9c6d3;border-radius:7px;background:white;color:#18324a;} QPushButton:hover{background:#eef5fb;} QPushButton#primary{background:#1f4e78;color:white;border-color:#1f4e78;font-weight:700;} QTableWidget{border:1px solid #d8e1ea;border-radius:8px;gridline-color:#e7eef5;} QHeaderView::section{background:#edf3f8;color:#27425d;padding:8px;border:0;border-bottom:1px solid #d8e1ea;font-weight:700;} QLineEdit,QComboBox,QTextEdit{border:1px solid #bdc9d5;border-radius:6px;padding:6px;background:white;}"
    )
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
