from __future__ import annotations

import sys
from pathlib import Path

from PyQt6.QtCore import pyqtSlot, QSize, QByteArray, Qt
from PyQt6.QtWidgets import QMainWindow, QApplication, QFileDialog, QWidget, QPushButton, QGridLayout, QCheckBox
from PyQt6.QtGui import QPalette, QColor, QPixmap, QIcon

DOC_IN_EXTS = ('docx',)

B64_DOC_PLURAL = b"iVBORw0KGgoAAAANSUhEUgAAADcAAABABAMAAABM0vIMAAAAGFBMVEUAAAD///8AAABAQECAgIC/v79hYWEWFhbk7YeIAAAAAXRSTlMAQObYZgAAAPNJREFUOMvt0U1uwjAQhmFTmgNYlAMwDWaduGm3BUWw5n8bOEIE4vokFuBvRmCQgBW8q0iPvpEiq5eqPfdJi0ry/QtsZvoUGYEffUAqQmgCKKYSTQDFVKK5hCVOJa5xKrFFfuoRS00IqQihEXh4t7HDeoq4JFdSfdo8NQzbZVYPpzX+6AEViMtd4n7liLHHavjHsQf4RcTQpoidmjzqye04tVXZBWw5fJ+9dvbRGOeu0VmkQ2cQeh5uAJtbbrMuYFQSb+Wx0kU2xzSi+kw0JrAj0AI2epYXA6qZ/Ub7HSOqSBzm2BU4AGzsctbQKGhheX11R3v4kY3dSHXipgAAAABJRU5ErkJggg=="

SIZE_BUTTON = QSize(27, 32)
STYLE = """
QPushButton {
    border: none;
    background: #333;
    color: white;
    width: 250px;
    height: 80px;
    border-radius: 6px;
    margin: 5px;
    padding: 10px;
    font-size: 14px;
}

QPushButton:hover {
    background-color: #555;
}

QPushButton:pressed {
    background-color: #666;
}
"""

# https://stackoverflow.com/a/52298774
def iconFromB64(b64: bytes):
    pixmap = QPixmap()
    pixmap.loadFromData(QByteArray.fromBase64(b64))
    return QIcon(pixmap)    

class Main(QMainWindow):

    def __init__(self: Main) -> None:
        super().__init__()
        self.do_layout()

    def do_layout(self: Main) -> None:
        self.setWindowTitle("Bulletin Collator")
        self.setFixedSize(QSize(350, 200))

        b_choose_folder = QPushButton(iconFromB64(B64_DOC_PLURAL), "  Choose input document folder")

        b_choose_folder.setIconSize(SIZE_BUTTON)

        b_choose_folder.clicked.connect(lambda _: self.run())

        layout = QGridLayout()
        layout.addWidget(b_choose_folder, 0, 0, 0, 2, Qt.AlignmentFlag.AlignCenter)

        widget = QWidget()
        widget.setLayout(layout)

        widget.setAutoFillBackground(True)
        palette = widget.palette()
        palette.setColor(QPalette.ColorRole.Window, QColor('#111'))
        widget.setPalette(palette)

        self.setCentralWidget(widget)

    @pyqtSlot()
    def ask_folder(self: QMainWindow, keyword: str) -> Path:
        path = QFileDialog.getExistingDirectory(
            self, f"Choose folder of {keyword} to convert", ".",
            options=QFileDialog.Option.DontUseNativeDialog
        )

        return Path(path) if path else None
    
    def run(self):
        pass

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLE)
    main_gui = Main()
    main_gui.show()
    sys.exit(app.exec())
