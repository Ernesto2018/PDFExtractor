__author__ = "Ernesto"
__email__ = "ernestondieki12@gmail.com"
__dated__ = "12-09-22"  # DDMMYY

from PyQt5.QtCore import QUrl
from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEngineSettings


class PDFViewer(QWebEngineView):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.settings().setAttribute(QWebEngineSettings.WebAttribute.PluginsEnabled, True)
        self.settings().setAttribute(QWebEngineSettings.WebAttribute.PdfViewerEnabled, True)

    def changeUrl(self, url: str):
        """ open new file """
        qurl = QUrl.fromLocalFile(url)
        if qurl.isValid():
            self.load(qurl)
