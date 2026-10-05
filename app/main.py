from app.core.resources import resource_base, resource_path
import sys
import os
from pathlib import Path

from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QFont, QIcon, QPixmap
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QStackedWidget,
    QVBoxLayout, QLabel, QGridLayout, QFrame, QHBoxLayout,
    QCommandLinkButton, QStyle, QSizePolicy
)
from PySide6.QtWebEngineCore import QWebEngineProfile

# Import das páginas do App
from app.pages.heatmap_page import HeatmapPage
from app.pages.correlation_page import CorrelationPage
from app.pages.crosstabs_page import CrosstabsPage
from app.pages.residuals_page import ResidualsPage

# Tutoriais
from app.pages.tutorials_pages import TutorialsHomePage, TutorialViewerPage
from app.pages.tutorial_popup import TutorialPopup

from app.ui.topbar import TopBar


APP_QSS = """
QMainWindow {
    background: #0B5CAD;
}

QWidget#Page {
    background: transparent;
}

QLabel#HeroTitle {
    font-size: 27px;
    font-weight: 750;
    color: #0B1F2A;
}

QLabel#HeroSubtitle {
    font-size: 12px;
    color: #60717D;
}

QFrame#CenterCard {
    background: rgba(248, 251, 255, 0.98);
    border: 1px solid #DCE5EE;
    border-radius: 20px;
}

QFrame#HeaderDivider {
    background: #DDE6EE;
    min-height: 1px;
    max-height: 1px;
    border: none;
}

QCommandLinkButton[variant="feature"] {
    background: #123A66;
    border: 1px solid #245B8F;
    border-radius: 14px;
    padding: 13px 15px;
    text-align: left;
    color: #FFFFFF;
}

QCommandLinkButton[variant="feature"]:hover {
    background: #174A7D;
    border: 1px solid #6FA8DC;
    color: #FFFFFF;
}

QCommandLinkButton[variant="feature"]:pressed {
    background: #0E3156;
    border: 1px solid #8AB8E6;
    color: #FFFFFF;
}

QCommandLinkButton[variant="feature"]:focus {
    border: 2px solid #7FB13D;
    color: #FFFFFF;
}

QCommandLinkButton[variant="feature"] QLabel {
    color: #FFFFFF;
}
"""




class HomePage(QWidget):
    def __init__(self):
        super().__init__()
        self.setObjectName("Page")

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)

        center_row = QHBoxLayout()
        center_row.setContentsMargins(30, 30, 30, 30)
        center_row.addStretch(1)

        card = QFrame()
        card.setObjectName("CenterCard")
        card.setMinimumWidth(780)
        card.setMaximumWidth(1120)
        card.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Preferred
        )

        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(34, 28, 34, 24)
        card_layout.setSpacing(14)

        # Cabeçalho da Home: título à esquerda e logo no canto superior direito
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(18)

        text_layout = QVBoxLayout()
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(6)

        title = QLabel("Ferramentas Estatísticas")
        title.setObjectName("HeroTitle")

        subtitle = QLabel("Escolha uma funcionalidade para iniciar.")
        subtitle.setObjectName("HeroSubtitle")

        text_layout.addWidget(title)
        text_layout.addWidget(subtitle)

        logo = QLabel()
        logo.setAlignment(Qt.AlignRight | Qt.AlignTop)
        logo.setMinimumWidth(180)
        logo.setSizePolicy(
            QSizePolicy.Policy.Preferred,
            QSizePolicy.Policy.Fixed
        )

        logo_pixmap = QPixmap(resource_path("assets/logo_home.png"))
        if not logo_pixmap.isNull():
            logo.setPixmap(
                logo_pixmap.scaled(
                    215,
                    72,
                    Qt.KeepAspectRatio,
                    Qt.SmoothTransformation
                )
            )

        header_layout.addLayout(text_layout, 1)
        header_layout.addWidget(logo, 0, Qt.AlignRight | Qt.AlignTop)

        card_layout.addLayout(header_layout)

        divider = QFrame()
        divider.setObjectName("HeaderDivider")
        card_layout.addWidget(divider)

        grid = QGridLayout()
        grid.setHorizontalSpacing(14)
        grid.setVerticalSpacing(14)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)

        self.btn_heatmap = QCommandLinkButton(
            "Mapa de Calor",
            "Importe um Relatório Descritivo Quantitativo e o KMLs do mapa e faça o Mapa de Calor."
        )


        self.btn_correlation = QCommandLinkButton(
            "Matriz de Correlação",
            "Importe o banco com peso no Excel e exporte um arquivo em Excel da tabela de Correlação."
        )

        self.btn_crosstabs = QCommandLinkButton(
            "Cruzamentos (Tabelas)",
            "Importe o banco com Peso do SPSS e faça os cruzamentos das tabelas com as variáveis que desejar."
        )

        self.btn_residuals = QCommandLinkButton(
            "Análise Residual",
            "Crosstab ponderado + resíduos (Ajustado Z/Pearson) e exportação em Excel."
        )


        self.btn_tutorials = QCommandLinkButton(
            "Tutoriais",
            "Acesse guias passo a passo de cada ferramenta do programa."
        )

        buttons = [
            self.btn_heatmap,
            self.btn_correlation,
            self.btn_crosstabs,
            self.btn_residuals,
            self.btn_tutorials,
        ]

        icons = [
            self.style().standardIcon(QStyle.SP_DriveNetIcon),
            self.style().standardIcon(QStyle.SP_FileDialogInfoView),
            self.style().standardIcon(QStyle.SP_FileDialogListView),
            self.style().standardIcon(QStyle.SP_DialogHelpButton),
            self.style().standardIcon(QStyle.SP_FileDialogContentsView),
        ]

        for btn, ic in zip(buttons, icons):
            btn.setProperty("variant", "feature")
            btn.setCursor(Qt.PointingHandCursor)
            btn.setMinimumHeight(86)
            btn.setMinimumWidth(0)
            btn.setSizePolicy(
                QSizePolicy.Policy.Expanding,
                QSizePolicy.Policy.Fixed
            )
            btn.setIcon(ic)
            btn.setIconSize(QSize(22, 22))

        grid.addWidget(self.btn_heatmap, 0, 0)
        grid.addWidget(self.btn_correlation, 0, 1)
        grid.addWidget(self.btn_crosstabs, 1, 0)
        grid.addWidget(self.btn_residuals, 1, 1)
        grid.addWidget(self.btn_tutorials, 2, 0, 1, 2)

        card_layout.addLayout(grid)

        footer = QLabel("Instituto Informa • Ferramentas Estatísticas v2.0")
        footer.setStyleSheet("color:#7A8791; font-size:10px; padding-top:4px;")
        card_layout.addWidget(footer, alignment=Qt.AlignRight)

        center_row.addWidget(card)
        center_row.addStretch(1)

        root.addStretch(2)
        root.addLayout(center_row)
        root.addStretch(3)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("App Desktop - Processamentos da Relatória - Instituto Informa")
        self.resize(1400, 850)

        # Guarda referências dos popups abertos para o Python não destruir a janela.
        self._tutorial_popups = []

        # Profile compartilhado do app
        self.web_profile = QWebEngineProfile("AppProfile", self)
        self.web_profile.downloadRequested.connect(lambda d: d.cancel())

        self.stack = QStackedWidget(self)
        self.setCentralWidget(self.stack)

        # Páginas
        self.home = HomePage()
        self.heatmap = HeatmapPage(self, profile=self.web_profile)
        self.correlation = CorrelationPage(self, profile=self.web_profile)
        self.crosstabs_page = CrosstabsPage(self, profile=self.web_profile)
        self.residuals = ResidualsPage(self, profile=self.web_profile)

        # Tutoriais antigos, pela página de tutoriais
        self.tutorials_home = TutorialsHomePage()
        self.tutorial_viewer = TutorialViewerPage(self, profile=self.web_profile)

        # Stack
        self.stack.addWidget(self.home)
        self.stack.addWidget(self.heatmap)
        self.stack.addWidget(self.correlation)
        self.stack.addWidget(self.crosstabs_page)
        self.stack.addWidget(self.residuals)
        self.stack.addWidget(self.tutorials_home)
        self.stack.addWidget(self.tutorial_viewer)

        # Ícones da home
        self.home.btn_heatmap.setIcon(QIcon(resource_path("assets/icons/heatmap.ico")))
        self.home.btn_correlation.setIcon(QIcon(resource_path("assets/icons/correlation.ico")))
        self.home.btn_crosstabs.setIcon(QIcon(resource_path("assets/icons/table.ico")))
        self.home.btn_residuals.setIcon(QIcon(resource_path("assets/icons/residual_graph.png")))
        self.home.btn_tutorials.setIcon(QIcon(resource_path("assets/icons/tutorial.ico")))

        # Navegação - Home
        self.home.btn_heatmap.clicked.connect(lambda: self.stack.setCurrentWidget(self.heatmap))
        self.home.btn_correlation.clicked.connect(lambda: self.stack.setCurrentWidget(self.correlation))
        self.home.btn_crosstabs.clicked.connect(lambda: self.stack.setCurrentWidget(self.crosstabs_page))
        self.home.btn_residuals.clicked.connect(lambda: self.stack.setCurrentWidget(self.residuals))
        self.home.btn_tutorials.clicked.connect(lambda: self.stack.setCurrentWidget(self.tutorials_home))

        # Voltar para Home
        self.heatmap.btn_back.clicked.connect(lambda: self.stack.setCurrentWidget(self.home))
        self.correlation.btn_back.clicked.connect(lambda: self.stack.setCurrentWidget(self.home))
        self.crosstabs_page.btn_back.clicked.connect(lambda: self.stack.setCurrentWidget(self.home))
        self.residuals.btn_back.clicked.connect(lambda: self.stack.setCurrentWidget(self.home))

        # Botões de ajuda nas páginas principais
        self.connect_help_buttons()

        # Tutoriais: voltar
        self.tutorials_home.btn_back.clicked.connect(lambda: self.stack.setCurrentWidget(self.home))
        self.tutorial_viewer.btn_back.clicked.connect(lambda: self.stack.setCurrentWidget(self.tutorials_home))

        # Abrir tutorial específico pela página antiga de tutoriais
        self.tutorials_home.btn_tut_heatmap.clicked.connect(
            lambda: self.open_tutorial(
                "Mapa de Calor",
                "assets/tutorial/guia_mapa_de_calor.html"
            )
        )


        self.tutorials_home.btn_tut_crosstabs.clicked.connect(
            lambda: self.open_tutorial(
                "Cruzamentos",
                "assets/tutorial/cruzamentos/guia_cruzamentos.html"
            )
        )

        self.tutorials_home.btn_tut_correlation.clicked.connect(
            lambda: self.open_tutorial(
                "Matriz de Correlação",
                "assets/tutorial/correlacao/guia_correlacao.html"
            )
        )

    def connect_help_buttons(self):
        """
        Conecta o botão ? de cada página ao tutorial correspondente.
        Usa hasattr para evitar erro caso alguma página ainda não tenha self.topbar.
        """

        help_map = [
            (
                self.heatmap,
                "Mapa de Calor",
                "assets/tutorial/guia_mapa_de_calor.html"
            ),
            (
                self.crosstabs_page,
                "Cruzamentos",
                "assets/tutorial/cruzamentos/guia_cruzamentos.html"
            ),
            (
                self.correlation,
                "Matriz de Correlação",
                "assets/tutorial/correlacao/guia_correlacao.html"
            ),

            # Quando você criar o tutorial da Análise Residual, descomente:
            # (
            #     self.residuals,
            #     "Análise Residual",
            #     "assets/tutorial/residuos/guia_residuos.html"
            # ),
        ]

        for page, title, rel_html in help_map:
            if hasattr(page, "topbar") and hasattr(page.topbar, "btn_help"):
                page.topbar.btn_help.show()
                page.topbar.btn_help.clicked.connect(
                    lambda checked=False, t=title, h=rel_html: self.show_tutorial_popup(t, h)
                )

    def open_tutorial(self, title: str, rel_html: str):
        """
        Mantém o funcionamento antigo da página de tutoriais.
        """
        self.tutorial_viewer.set_title(title)
        self.tutorial_viewer.load_html_file(rel_html)
        self.stack.setCurrentWidget(self.tutorial_viewer)

    def show_tutorial_popup(self, title: str, rel_html: str):
        """
        Novo funcionamento: abre o tutorial em popup, sem sair da página atual.
        """
        popup = TutorialPopup(
            title=title,
            rel_html=rel_html,
            profile=self.web_profile,
            parent=self
        )

        self._tutorial_popups.append(popup)
        popup.destroyed.connect(
            lambda _=None, p=popup: self._remove_tutorial_popup(p)
        )

        popup.show()

    def _remove_tutorial_popup(self, popup):
        try:
            if popup in self._tutorial_popups:
                self._tutorial_popups.remove(popup)
        except Exception:
            pass

    def closeEvent(self, event):
        try:
            for page in (
                self.heatmap,
                self.correlation,
                self.crosstabs_page,
                self.residuals,
            ):
                try:
                    page.cleanup()
                except Exception:
                    pass
        finally:
            event.accept()


def main():
    app = QApplication(sys.argv)

    app.setFont(QFont("DIN", 10))
    app.setStyle("Fusion")
    app.setStyleSheet(APP_QSS + TopBar.qss())

    w = MainWindow()
    w.showMaximized()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()