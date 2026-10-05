"""Focused parcel editor with quick parsing and spacious field review."""
from __future__ import annotations

from PySide6.QtCore import Signal
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QScrollArea,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from kml_engine import FIELDS, Parcel, STATUS_LABELS, parse_description


class ParcelEditor(QWidget):
    save_requested = Signal(dict, bool)
    skip_requested = Signal(dict)

    def __init__(self):
        super().__init__()
        self._fields: dict[str, QLineEdit] = {}
        self._search_coordinate = ""
        self._build()

    def _build(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(10)

        heading = QLabel("تفاصيل القطعة")
        heading.setObjectName("sectionTitle")
        layout.addWidget(heading)

        self.note = QLabel("اختر قطعة من جدول النتائج.")
        self.note.setObjectName("statusNote")
        self.note.setWordWrap(True)
        layout.addWidget(self.note)

        self.tabs = QTabWidget()
        self._build_fields_tab()
        self._build_quick_entry_tab()
        layout.addWidget(self.tabs, 1)

        self.copy_button = QPushButton("نسخ نقطة البحث في Google Earth")
        self.copy_button.setObjectName("primary")
        self.copy_button.setEnabled(False)
        self.copy_button.clicked.connect(self.copy_search_coordinate)
        layout.addWidget(self.copy_button)

        actions = QHBoxLayout()
        save = QPushButton("حفظ التعديل")
        save.clicked.connect(lambda: self.save_requested.emit(self.values(), False))
        approve = QPushButton("حفظ واعتماد المراجعة")
        approve.setObjectName("success")
        approve.clicked.connect(lambda: self.save_requested.emit(self.values(), True))
        self.skip_missing = QPushButton("تجاوز النواقص وتركها فارغة")
        self.skip_missing.setObjectName("warning")
        self.skip_missing.setEnabled(False)
        self.skip_missing.setToolTip(
            "يحفظ الحقول الحالية، ويرفع حظر Excel مع الإبقاء على الخانات الناقصة فارغة."
        )
        self.skip_missing.clicked.connect(
            lambda: self.skip_requested.emit(self.values())
        )
        actions.addWidget(save)
        actions.addWidget(approve)
        actions.addWidget(self.skip_missing)
        layout.addLayout(actions)

        skip_hint = QLabel(
            "عند اكتمال الحقول المطلوبة استخدم «حفظ التعديل». وللسماح بترك النواقص فارغة، استخدم زر التجاوز."
        )
        skip_hint.setObjectName("muted")
        skip_hint.setWordWrap(True)
        layout.addWidget(skip_hint)

    def _build_fields_tab(self) -> None:
        page = QWidget()
        outer = QVBoxLayout(page)
        outer.setContentsMargins(10, 12, 10, 8)
        guide = QLabel("راجع القيم المستخرجة وعدّل أي خانة عند الحاجة.")
        guide.setObjectName("muted")
        outer.addWidget(guide)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        form_host = QWidget()
        grid = QGridLayout(form_host)
        grid.setContentsMargins(8, 8, 8, 8)
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(3, 1)

        for index, (key, label) in enumerate(FIELDS):
            row, pair = divmod(index, 2)
            label_widget = QLabel(label)
            label_widget.setObjectName("fieldLabel")
            editor = QLineEdit()
            editor.setMinimumHeight(34)
            self._fields[key] = editor
            grid.addWidget(label_widget, row, pair * 2)
            grid.addWidget(editor, row, pair * 2 + 1)

        scroll.setWidget(form_host)
        outer.addWidget(scroll, 1)
        self.tabs.addTab(page, "الحقول المستخرجة")

    def _build_quick_entry_tab(self) -> None:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(12, 12, 12, 12)
        guide = QLabel(
            "الصق الوصف كاملًا كما هو، ثم اضغط «قراءة وتوزيع الحقول»."
        )
        guide.setObjectName("muted")
        guide.setWordWrap(True)
        layout.addWidget(guide)
        self.raw_description = QTextEdit()
        self.raw_description.setPlaceholderText(
            "id: ... | رقم الحيازة: ... | اسم الحائز: ..."
        )
        layout.addWidget(self.raw_description, 1)
        parse = QPushButton("قراءة وتوزيع الحقول")
        parse.setObjectName("primary")
        parse.clicked.connect(self.parse_raw)
        layout.addWidget(parse)
        self.tabs.addTab(page, "لصق وصف كامل")

    def set_parcel(self, parcel: Parcel) -> None:
        self.note.setText(
            f"{STATUS_LABELS.get(parcel.status, parcel.status)} — {parcel.reason}"
        )
        self.raw_description.setPlainText(parcel.description)
        self._search_coordinate = parcel.search_coordinate
        self.copy_button.setEnabled(bool(self._search_coordinate))
        self.skip_missing.setEnabled(
            parcel.status in {"empty", "no_description", "incomplete", "unreadable"}
        )
        for key, editor in self._fields.items():
            editor.setText(parcel.data.get(key, ""))
        self.tabs.setCurrentIndex(0)

    def parse_raw(self) -> None:
        raw = self.raw_description.toPlainText().strip()
        if not raw:
            self.note.setText("ألصق الوصف أولًا ثم أعد المحاولة.")
            return
        parsed = parse_description(raw)
        populated = 0
        for key, editor in self._fields.items():
            if parsed.get(key):
                editor.setText(parsed[key])
                populated += 1
        self.note.setText(f"تم توزيع {populated} حقلًا. راجع القيم ثم احفظ التعديل.")
        self.tabs.setCurrentIndex(0)

    def values(self) -> dict[str, str]:
        return {key: editor.text() for key, editor in self._fields.items()}

    def copy_search_coordinate(self) -> None:
        QGuiApplication.clipboard().setText(self._search_coordinate)
        self.note.setText("تم نسخ نقطة مركز القطعة. الصقها في بحث Google Earth.")
