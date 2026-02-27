from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QListWidget,
    QListWidgetItem, QPushButton, QLabel, QApplication,
)
from PyQt6.QtCore import Qt

VIDEO_EXTENSIONS = (".mp4", ".avi", ".mov", ".mkv")

BROWSER_STYLE = """
    QDialog, QWidget {
        background-color: #1e1e1e;
        color: #f0f0f0;
    }
    QListWidget {
        background-color: #2d2d2d;
        border: 1px solid #444;
        border-radius: 6px;
        font-size: 13px;
        padding: 4px;
    }
    QListWidget::item {
        padding: 6px 8px;
        border-radius: 4px;
    }
    QListWidget::item:selected {
        background-color: #0078d4;
    }
    QListWidget::item:hover:!selected {
        background-color: #3a3a3a;
    }
    QPushButton {
        background-color: #2d2d2d;
        color: #f0f0f0;
        border: 1px solid #555;
        border-radius: 6px;
        padding: 6px 16px;
        font-size: 13px;
    }
    QPushButton:hover { background-color: #3a3a3a; border-color: #888; }
    QPushButton:pressed { background-color: #0078d4; border-color: #0078d4; }
    QPushButton#select_btn {
        background-color: #0078d4;
        border-color: #0078d4;
        font-weight: bold;
    }
    QPushButton#select_btn:hover { background-color: #1a8fe3; }
    QPushButton#select_btn:disabled { background-color: #333; border-color: #444; color: #666; }
    QLabel { font-size: 12px; color: #aaa; }
"""


class OneDriveBrowser(QDialog):
    def __init__(self, client, parent=None):
        super().__init__(parent)
        self.client = client
        self.setWindowTitle("Browse OneDrive")
        self.setMinimumSize(640, 480)
        self.resize(720, 520)
        self.setStyleSheet(BROWSER_STYLE)

        self.nav_stack = []          # list of (item_id, display_name) tuples
        self.selected_item = None    # dict with id, name, parent_id
        self._current_folder_id = None
        self._current_folder_name = "OneDrive"

        self._setup_ui()
        self.load_folder(None, "OneDrive")

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(10)

        self.path_label = QLabel("📂  OneDrive")
        layout.addWidget(self.path_label)

        self.list_widget = QListWidget()
        self.list_widget.itemClicked.connect(self._on_item_clicked)
        self.list_widget.itemDoubleClicked.connect(self._on_item_double_clicked)
        layout.addWidget(self.list_widget)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)

        self.back_btn = QPushButton("← Back")
        self.back_btn.setEnabled(False)
        self.back_btn.clicked.connect(self._go_back)
        btn_row.addWidget(self.back_btn)

        btn_row.addStretch()

        self.select_btn = QPushButton("Select")
        self.select_btn.setObjectName("select_btn")
        self.select_btn.setEnabled(False)
        self.select_btn.clicked.connect(self.accept)
        btn_row.addWidget(self.select_btn)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)

        layout.addLayout(btn_row)

    def load_folder(self, item_id, name):
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            self.list_widget.clear()
            self.select_btn.setEnabled(False)
            self._current_folder_id = item_id
            self._current_folder_name = name

            items = self.client.list_folder(item_id)
            folders = sorted(
                [i for i in items if "folder" in i],
                key=lambda x: x["name"].lower(),
            )
            files = sorted(
                [i for i in items if "file" in i],
                key=lambda x: x["name"].lower(),
            )

            for item in folders:
                li = QListWidgetItem(f"📁  {item['name']}")
                li.setData(Qt.ItemDataRole.UserRole, {
                    "type": "folder",
                    "id": item["id"],
                    "name": item["name"],
                })
                self.list_widget.addItem(li)

            for item in files:
                if not item["name"].lower().endswith(VIDEO_EXTENSIONS):
                    continue
                li = QListWidgetItem(f"🎬  {item['name']}")
                li.setData(Qt.ItemDataRole.UserRole, {
                    "type": "file",
                    "id": item["id"],
                    "name": item["name"],
                    "parent_id": item.get("parentReference", {}).get("id"),
                })
                self.list_widget.addItem(li)

            crumbs = [p[1] for p in self.nav_stack] + [name]
            self.path_label.setText("📂  " + "  /  ".join(crumbs))
            self.back_btn.setEnabled(bool(self.nav_stack))
        finally:
            QApplication.restoreOverrideCursor()

    def _on_item_clicked(self, item):
        data = item.data(Qt.ItemDataRole.UserRole)
        if data["type"] == "file":
            self.selected_item = data
            self.select_btn.setEnabled(True)
        else:
            self.select_btn.setEnabled(False)

    def _on_item_double_clicked(self, item):
        data = item.data(Qt.ItemDataRole.UserRole)
        if data["type"] == "folder":
            self.nav_stack.append((self._current_folder_id, self._current_folder_name))
            self.load_folder(data["id"], data["name"])
        elif data["type"] == "file":
            self.selected_item = data
            self.accept()

    def _go_back(self):
        if self.nav_stack:
            parent_id, parent_name = self.nav_stack.pop()
            self.load_folder(parent_id, parent_name)
