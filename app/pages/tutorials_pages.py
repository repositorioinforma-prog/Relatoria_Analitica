from app.core.resources import resource_base, resource_path
import os
import sys
from pathlib import Path

from PySide6.QtCore import Qt, QSize, QUrl
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFrame, QLabel,
    QGridLayout, QCommandLinkButton, QStyle, QSizePolicy
)
from PySide6.QtWebEngineCore import QWebEnginePage, QWebEngineSettings, QWebEngineProfile
from PySide6.QtWebEngineWidgets import QWebEngineView

from app.ui.topbar import TopBar






class TutorialsHomePage(QWidget):
    def __init__(self):
        super().__init__()
        self.setObjectName("Page")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        topbar = TopBar(
            title="Tutoriais",
            icon_path=resource_path("assets/icons/tutorial.ico"),
            back_icon_path=resource_path("assets/icons/arrow-left.ico")
        )
        topbar.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.btn_back = topbar.btn_back
        layout.addWidget(topbar)

        center_row = QHBoxLayout()
        center_row.setContentsMargins(30, 30, 30, 30)
        center_row.addStretch(1)

        card = QFrame()
        card.setObjectName("CenterCard")
        card.setMinimumWidth(940)
        card.setMaximumWidth(1080)

        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(34, 26, 34, 26)
        card_layout.setSpacing(16)

        hero = QLabel("Guias passo a passo")
        hero.setObjectName("HeroTitle")

        subtitle = QLabel("Selecione o tutorial que deseja abrir.")
        subtitle.setObjectName("HeroSubtitle")

        card_layout.addWidget(hero)
        card_layout.addWidget(subtitle)

        grid = QGridLayout()
        grid.setHorizontalSpacing(16)
        grid.setVerticalSpacing(16)
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)
        grid.setAlignment(Qt.AlignTop)

        self.btn_tut_heatmap = QCommandLinkButton(
            "Mapa de Calor",
            "Importar dados, configurar visual, pintar, numerar e exportar PPTX."
        )
        self.btn_tut_heatmap.setProperty("variant", "feature")
        self.btn_tut_heatmap.setCursor(Qt.PointingHandCursor)
        self.btn_tut_heatmap.setMinimumHeight(92)
        self.btn_tut_heatmap.setIcon(self.style().standardIcon(QStyle.SP_FileDialogInfoView))
        self.btn_tut_heatmap.setIconSize(QSize(22, 22))
        self.btn_tut_heatmap.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)


        self.btn_tut_crosstabs = QCommandLinkButton(
            "Cruzamentos",
            "Importar dados, configurar variáveis, gerar tabelas, ordenar linhas e colunas e exportar Excel."
        )
        self.btn_tut_crosstabs.setProperty("variant", "feature")
        self.btn_tut_crosstabs.setCursor(Qt.PointingHandCursor)
        self.btn_tut_crosstabs.setMinimumHeight(92)
        self.btn_tut_crosstabs.setIcon(self.style().standardIcon(QStyle.SP_FileDialogInfoView))
        self.btn_tut_crosstabs.setIconSize(QSize(22, 22))
        self.btn_tut_crosstabs.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.btn_tut_correlation = QCommandLinkButton(
            "Matriz de Correlação",
            "Importar dados, selecionar variáveis, configurar peso e missing, calcular a matriz e exportar CSV/XLSX."
        )
        self.btn_tut_correlation.setProperty("variant", "feature")
        self.btn_tut_correlation.setCursor(Qt.PointingHandCursor)
        self.btn_tut_correlation.setMinimumHeight(92)
        self.btn_tut_correlation.setIcon(self.style().standardIcon(QStyle.SP_FileDialogInfoView))
        self.btn_tut_correlation.setIconSize(QSize(22, 22))
        self.btn_tut_correlation.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        buttons = [
            self.btn_tut_heatmap,
            self.btn_tut_crosstabs,
            self.btn_tut_correlation,
        ]

        for i, btn in enumerate(buttons):
            r, c = divmod(i, 2)
            grid.addWidget(btn, r, c)

        card_layout.addLayout(grid)

        footer = QLabel("Dica: novos tutoriais podem ser adicionados aqui depois.")
        footer.setStyleSheet("color:#7A8791; font-size:11px;")
        card_layout.addWidget(footer, alignment=Qt.AlignRight)

        center_row.addWidget(card)
        center_row.addStretch(1)

        layout.addStretch(2)
        layout.addLayout(center_row)
        layout.addStretch(3)


class TutorialViewerPage(QWidget):
    def __init__(self, main_window, profile: QWebEngineProfile):
        super().__init__()
        self.main_window = main_window
        self._root = resource_base()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.topbar = TopBar(
            title="Tutorial",
            icon_path=resource_path("assets/icons/tutorial.ico"),
            back_icon_path=resource_path("assets/icons/arrow-left.ico")
        )
        self.topbar.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.btn_back = self.topbar.btn_back
        layout.addWidget(self.topbar)

        self.view = QWebEngineView(self)
        layout.addWidget(self.view, 1)

        page = QWebEnginePage(profile, self.view)
        self.view.setPage(page)

        s = self.view.settings()
        s.setAttribute(QWebEngineSettings.LocalContentCanAccessFileUrls, True)
        s.setAttribute(QWebEngineSettings.LocalContentCanAccessRemoteUrls, True)

    def set_title(self, text: str):
        self.topbar.title_label.setText(f"Tutorial: {text}")

    def load_html_file(self, rel_path: str):
        abs_path = Path(resource_path(rel_path)).resolve()
        if not abs_path.exists():
            self.view.setHtml(f"<h3>Arquivo não encontrado:</h3><pre>{abs_path}</pre>")
            return

        html = abs_path.read_text(encoding="utf-8", errors="replace")

        # Como o HTML usa src="assets/tutorial/correlacao/..."
        # a base deve apontar para a raiz do projeto
        base = QUrl.fromLocalFile(str(self._root) + os.sep)
        self.view.setHtml(html, baseUrl=base)