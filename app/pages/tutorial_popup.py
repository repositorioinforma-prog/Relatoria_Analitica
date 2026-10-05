from app.core.resources import resource_base, resource_path
import os
import sys
from pathlib import Path

from PySide6.QtCore import QUrl
from PySide6.QtWidgets import QDialog, QVBoxLayout
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWebEngineCore import QWebEnginePage, QWebEngineSettings, QWebEngineProfile






class TutorialPopup(QDialog):
    def __init__(
        self,
        title: str,
        rel_html: str,
        profile: QWebEngineProfile,
        parent=None
    ):
        super().__init__(parent)

        self.setWindowTitle(f"Tutorial: {title}")
        self.resize(1100, 720)
        self.setModal(False)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.view = QWebEngineView(self)
        layout.addWidget(self.view)

        page = QWebEnginePage(profile, self.view)
        self.view.setPage(page)

        settings = self.view.settings()
        settings.setAttribute(QWebEngineSettings.LocalContentCanAccessFileUrls, True)
        settings.setAttribute(QWebEngineSettings.LocalContentCanAccessRemoteUrls, True)

        self.load_html_file(rel_html)

    def load_html_file(self, rel_path: str):
        abs_path = Path(resource_path(rel_path)).resolve()

        if not abs_path.exists():
            self.view.setHtml(f"<h3>Arquivo não encontrado:</h3><pre>{abs_path}</pre>")
            return

        html = abs_path.read_text(encoding="utf-8", errors="replace")

        # Mantém a mesma lógica do TutorialViewerPage:
        # a base aponta para a raiz do projeto, para funcionar assets/tutorial/...
        base = QUrl.fromLocalFile(str(resource_base()) + os.sep)
        self.view.setHtml(html, baseUrl=base)