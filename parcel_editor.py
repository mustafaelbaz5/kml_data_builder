"""Dedicated editor widget: parse a full Arabic description before fine editing fields."""
from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import QFormLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QScrollArea, QTabWidget, QTextEdit, QVBoxLayout, QWidget

from kml_engine import FIELDS, Parcel, STATUS_LABELS, parse_description


class ParcelEditor(QWidget):
    save_requested = Signal(dict, bool)

    def __init__(self):
        super().__init__(); self._parcel: Parcel | None = None; self._fields: dict[str, QLineEdit] = {}; self._build()

    def _build(self) -> None:
        layout = QVBoxLayout(self); layout.setContentsMargins(14, 14, 14, 14); layout.setSpacing(10)
        heading = QLabel("تعديل بيانات القطعة"); heading.setObjectName("sectionTitle"); layout.addWidget(heading)
        self.note = QLabel("اختر قطعة من جدول النتائج."); self.note.setObjectName("statusNote"); self.note.setWordWrap(True); layout.addWidget(self.note)
        self.tabs = QTabWidget(); layout.addWidget(self.tabs, 1); self._build_quick_entry(); self._build_fields()
        layout.addWidget(QLabel("إحداثيات القطعة", objectName="sectionTitle"))
        coordinate_row = QHBoxLayout(); self.coordinates = QTextEdit(); self.coordinates.setReadOnly(True); self.coordinates.setMaximumHeight(74)
        copy = QPushButton("نسخ الإحداثيات"); copy.clicked.connect(self.copy_coordinates); coordinate_row.addWidget(self.coordinates, 1); coordinate_row.addWidget(copy); layout.addLayout(coordinate_row)
        actions = QHBoxLayout(); save = QPushButton("حفظ التعديل"); save.clicked.connect(lambda: self.save_requested.emit(self.values(), False)); approve = QPushButton("حفظ واعتماد المراجعة"); approve.setObjectName("success"); approve.clicked.connect(lambda: self.save_requested.emit(self.values(), True)); actions.addWidget(save); actions.addWidget(approve); layout.addLayout(actions)

    def _build_quick_entry(self) -> None:
        page = QWidget(); layout = QVBoxLayout(page); layout.setContentsMargins(12, 12, 12, 12)
        guide = QLabel("الصق الوصف كاملًا كما هو. اضغط «قراءة وتوزيع الحقول» ثم راجع النتيجة في تبويب الحقول."); guide.setObjectName("muted"); guide.setWordWrap(True); layout.addWidget(guide)
        self.raw_description = QTextEdit(); self.raw_description.setPlaceholderText("id: ... | رقم الحيازة: ... | اسم الحائز: ..."); layout.addWidget(self.raw_description, 1)
        parse = QPushButton("قراءة وتوزيع الحقول"); parse.setObjectName("primary"); parse.clicked.connect(self.parse_raw); layout.addWidget(parse); self.tabs.addTab(page, "لصق وصف كامل")

    def _build_fields(self) -> None:
        page = QWidget(); outer = QVBoxLayout(page); outer.setContentsMargins(8, 8, 8, 8)
        scroll = QScrollArea(); scroll.setWidgetResizable(True); form_host = QWidget(); self.form = QFormLayout(form_host); self.form.setLabelAlignment(Qt.AlignRight)
        for key, label in FIELDS:
            editor = QLineEdit(); self._fields[key] = editor; self.form.addRow(label + ":", editor)
        scroll.setWidget(form_host); outer.addWidget(scroll); self.tabs.addTab(page, "الحقول المستخرجة")

    def set_parcel(self, parcel: Parcel) -> None:
        self._parcel = parcel; self.note.setText(f"{STATUS_LABELS.get(parcel.status, parcel.status)} — {parcel.reason}"); self.raw_description.setPlainText(parcel.description); self.coordinates.setPlainText(parcel.coordinates)
        for key, editor in self._fields.items(): editor.setText(parcel.data.get(key, ""))

    def parse_raw(self) -> None:
        raw = self.raw_description.toPlainText().strip()
        if not raw: self.note.setText("ألصق الوصف أولًا ثم أعد المحاولة."); return
        parsed = parse_description(raw); populated = 0
        for key, editor in self._fields.items():
            if parsed.get(key): editor.setText(parsed[key]); populated += 1
        self.note.setText(f"تم توزيع {populated} حقلًا. راجع الحقول ثم احفظ التعديل."); self.tabs.setCurrentIndex(1)

    def values(self) -> dict[str, str]:
        return {key: editor.text() for key, editor in self._fields.items()}

    def copy_coordinates(self) -> None:
        QGuiApplication.clipboard().setText(self.coordinates.toPlainText()); self.note.setText("تم نسخ الإحداثيات. الصقها في Google Earth للعثور على القطعة.")
