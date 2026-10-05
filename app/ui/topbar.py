from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QFrame, QHBoxLayout, QToolButton, QLabel


TOPBAR_QSS = """
QFrame#TopBar {
    min-height: 56px;
    max-height: 56px;
    border-bottom: 1px solid rgba(0,0,0,0.08);
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
        stop:0 #0B5CAD, stop:1 #10B981);
}

/* Botão voltar “pill” */
QToolButton#BackButton {
    background: rgba(255,255,255,0.16);
    border: 1px solid rgba(255,255,255,0.30);
    color: white;
    padding: 6px 12px;
    border-radius: 12px;
    font-weight: 600;
}
QToolButton#BackButton:hover {
    background: rgba(255,255,255,0.24);
    border: 1px solid rgba(255,255,255,0.38);
}
QToolButton#BackButton:pressed {
    background: rgba(0,0,0,0.10);
}
QToolButton#BackButton:focus {
    border: 2px solid rgba(255,255,255,0.70);
}

/* Botão de ajuda */
QToolButton#HelpButton {
    background: rgba(255,255,255,0.16);
    border: 1px solid rgba(255,255,255,0.30);
    color: white;
    min-width: 34px;
    max-width: 34px;
    min-height: 34px;
    max-height: 34px;
    border-radius: 17px;
    font-weight: 900;
    font-size: 16px;
}
QToolButton#HelpButton:hover {
    background: rgba(255,255,255,0.26);
    border: 1px solid rgba(255,255,255,0.45);
}
QToolButton#HelpButton:pressed {
    background: rgba(0,0,0,0.10);
}
QToolButton#HelpButton:focus {
    border: 2px solid rgba(255,255,255,0.70);
}

/* Chip do título */
QFrame#TitleChip {
    background: rgba(255,255,255,0.16);
    border: 1px solid rgba(255,255,255,0.22);
    border-radius: 14px;
    padding: 6px 12px;
}
QLabel#TitleText {
    color: white;
    font-weight: 700;
    font-size: 13px;
}
"""


class TopBar(QFrame):
    """
    Topbar padrão para todas as páginas.
    Exponha self.btn_back para manter compatibilidade com seu código atual.
    Agora também expõe self.btn_help para abrir o tutorial da página.
    """
    def __init__(
        self,
        title: str,
        icon_path: str | None = None,
        back_icon_path: str | None = None,
        show_help: bool = False
    ):
        super().__init__()
        self.setObjectName("TopBar")

        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 10, 16, 10)
        layout.setSpacing(12)

        # Botão voltar
        self.btn_back = QToolButton()
        self.btn_back.setObjectName("BackButton")
        self.btn_back.setCursor(Qt.PointingHandCursor)
        self.btn_back.setText("Voltar")
        self.btn_back.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.btn_back.setIconSize(QSize(18, 18))

        if back_icon_path:
            self.btn_back.setIcon(QIcon(back_icon_path))

        layout.addWidget(self.btn_back)

        # Espaço antes do título
        layout.addStretch(1)

        # Chip do título
        chip = QFrame()
        chip.setObjectName("TitleChip")

        chip_layout = QHBoxLayout(chip)
        chip_layout.setContentsMargins(10, 0, 10, 0)
        chip_layout.setSpacing(8)

        self.icon_label = QLabel()
        self.icon_label.setFixedSize(18, 18)
        self.icon_label.setScaledContents(True)

        if icon_path:
            self.icon_label.setPixmap(QIcon(icon_path).pixmap(18, 18))

        chip_layout.addWidget(self.icon_label)

        self.title_label = QLabel(title)
        self.title_label.setObjectName("TitleText")
        chip_layout.addWidget(self.title_label)

        layout.addWidget(chip)

        # Espaço entre título e ajuda
        layout.addStretch(1)

        # Botão ajuda/tutorial
        self.btn_help = QToolButton()
        self.btn_help.setObjectName("HelpButton")
        self.btn_help.setCursor(Qt.PointingHandCursor)
        self.btn_help.setText("?")
        self.btn_help.setToolTip("Abrir tutorial desta página")
        self.btn_help.setFixedSize(34, 34)

        layout.addWidget(self.btn_help)

        if not show_help:
            self.btn_help.hide()

    @staticmethod
    def qss():
        return TOPBAR_QSS