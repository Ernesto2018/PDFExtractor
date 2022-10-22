__author__ = "Ernesto"
__email__ = "ernestondieki12@gmail.com"
__dated__ = "12-09-22"  # DDMMYY

from PyQt5.QtCore import Qt, pyqtSignal, QThreadPool
from custom import CustomDialog, r_path
from viewer import PDFViewer
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (
    QWidget, QApplication, QPushButton,
    QLabel, QVBoxLayout, QHBoxLayout,
    QMainWindow, QFrame, QLineEdit,
    QListWidget, QFileDialog, QProgressBar,
    QMessageBox,
)
from os.path import (
    normpath,
    basename, expanduser, join,
)
from utils import (
    getTextBbox, XLWriter,
    getName, getEmail,
    getInvoiceNum,
)
from os import scandir
import sys


STYLE = """
QWidget {
    font: 16px;
    font-family: Consolas;
}
QProgressBar {
    border: 2px solid #58c1d1;
    border-radius: 5px;
}
QProgressBar::chunk {
    background-color: #58c1d1;
    width: 10px;
}
QPushButton {
    background-color: #0cb3aa;
    border-radius: 5px;
    min-height: 30px;
}
QPushButton:hover {
    background-color: gray;
    border-radius: 5px;
}
QPushButton:disabled {
    color: gray;
}
QPushButton:pressed {
    background-color: #cacaca;
}
QMessageBox QPushButton {
    min-width: 60px;
    min-height: 30px;
}
QDialog QPushButton {
    min-width: 60px;
    min-height: 30px;
}
QPushButton#selectf {
    min-width: 150px;
}
QLineEdit {
    border-radius: 4px;
    border: 1px solid gray;
    min-height: 25px;
}
QLineEdit:focus {
    border: 1px solid #0cb3aa;
}
QListWidget {
    border: 2px solid white;
    border-radius: 4px;
}
QPushButton#stopbtn {
    background-color: #fa7e75;
}
QPushButton#stopbtn:hover {
    background-color: #ed473b;
}
QPushButton#stopbtn:pressed {
    background-color: #fca8a2;
}
QPushButton#stopbtn:disabled {
    color: gray;
    background-color: #edc1be;
}
"""


class RowEdit(QFrame):
    """ top btn """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        layout = QVBoxLayout(self)
        title = QLabel(f"Where the data will start")
        title.setObjectName("where")
        layout.addWidget(title)
        self.edit = QLineEdit("1")
        self.edit.setFixedWidth(120)
        layout.addWidget(self.edit)


class Zone(QFrame):
    """ """

    def __init__(self, num, *args, **kwargs):
        super().__init__(*args, **kwargs)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"Zone # {num}"))
        self.edit = QLineEdit()
        self.edit.setFixedWidth(120)
        layout.addWidget(self.edit)


class Home(QWidget):
    """ home widgets """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # zone record
        self.zones = []

        # main layout for all the widgets and inner layouts
        self.main_layout = QHBoxLayout(self)
        self.create_widgets()

    def create_widgets(self):
        """ create widgets for the main win """

        # pdf viewer instance
        self.pdf_viewer = PDFViewer()
        # add it to the main layout first
        self.main_layout.addWidget(self.pdf_viewer, stretch=3)

        # right side layout
        self.right_layout = QVBoxLayout()
        create_zone_title = QLabel("Create a PDF zone")
        # map the top title
        self.right_layout.addWidget(create_zone_title)
        # parse zone layout
        parser_zone_layout = QHBoxLayout()
        # create parse zone btn
        self.create_parse_zone = QPushButton("+")
        self.create_parse_zone.setFixedWidth(40)
        # map widgets to layout
        parser_zone_layout.addWidget(self.create_parse_zone, stretch=1)
        title = QLabel("Create Parsing Zone")
        title.setObjectName("parseZone")
        parser_zone_layout.addWidget(title, stretch=4)
        # map the btn and label
        self.right_layout.addLayout(parser_zone_layout)
        zones_title = QLabel("Zones")
        zones_title.setObjectName("zones")
        # map zones title
        self.right_layout.addWidget(zones_title)
        # btns
        self.row_edit = RowEdit()
        # map btns
        self.right_layout.addWidget(self.row_edit)
        # create zones

        self.zoneA = Zone(1)
        self.zoneA.edit.setText("A")
        # map zone
        self.right_layout.addWidget(self.zoneA)

        self.zoneB = Zone(2)
        self.zoneB.edit.setText("B")
        # map zone
        self.right_layout.addWidget(self.zoneB)

        self.zoneC = Zone(3)
        self.zoneC.edit.setText("C")
        # map zone
        self.right_layout.addWidget(self.zoneC)

        self.zoneD = Zone(4)
        self.zoneD.edit.setText("D")
        # map zone
        self.right_layout.addWidget(self.zoneD)

        # list widget
        listview_title = QLabel("Uploaded PDFs (double-click to preview)")
        self.right_layout.addWidget(listview_title)
        self.listview = QListWidget()

        self.right_layout.addWidget(self.listview)

        # select folder btn
        self.selectbtn = QPushButton("Select folder")
        self.selectbtn.setObjectName("selectf")
        self.right_layout.addWidget(self.selectbtn, alignment=Qt.AlignmentFlag.AlignLeft)

        # parse btn
        self.parsebtn = QPushButton("Parse PDFs")
        self.right_layout.addWidget(self.parsebtn)

        # progress bar
        self.progressbar = QProgressBar()
        self.progressbar.setRange(0, 100)
        self.progressbar.setTextVisible(True)
        self.progressbar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progressbar.hide()
        self.right_layout.addWidget(self.progressbar)

        self.stopbtn = QPushButton("STOP")
        self.stopbtn.setObjectName("stopbtn")
        self.stopbtn.hide()
        self.right_layout.addWidget(self.stopbtn)

        self.right_layout.addStretch()
        self.main_layout.addLayout(self.right_layout, stretch=1)

    def uploadPDFs(self, pdfs):
        """ add pdfs to the list view """
        # clear the previous items
        self.listview.clear()
        self.listview.addItems(pdfs)


class MainWin(QMainWindow):

    """ main application window """

    thread_started = pyqtSignal()
    thread_progress = pyqtSignal(int)
    thread_done = pyqtSignal()
    error = pyqtSignal(str, str)

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # set window title and icon
        self.setWindowTitle("PDFs Extractor")
        self.setWindowIcon(QIcon(r_path("app.png")))
        self.setMinimumSize(900, 300)

        # last selected dir
        self.lastopendir = expanduser("~\\Documents")
        self.filenames = []
        self.text_bbox = {}
        self.cancel_parsing = 0

        self.zone_1_bbox = ()
        self.zone_2_bbox = ()
        self.zone_3_bbox = ()
        self.zone_4_bbox = ()

        self.zone_1_text = ""
        self.zone_2_text = ""
        self.zone_3_text = ""
        self.zone_4_text = ""

        # map the main widgets
        self.home = Home()
        self.setCentralWidget(self.home)

        self.threadpool = QThreadPool(self)
        self.threadpool.setMaxThreadCount(1)

        # asign signals to slots
        self.home.selectbtn.clicked.connect(self.choose_dir)
        # on listview double-click
        self.home.listview.itemActivated.connect(self.preview_item)
        # on create parse zone click
        self.home.create_parse_zone.clicked.connect(self.addZone)
        # parse PDFs
        self.home.parsebtn.clicked.connect(self.process_zones)
        # stop btn
        self.home.stopbtn.clicked.connect(self.stop_parsing)

        # custom signals
        self.thread_started.connect(self.parsing_started)
        self.thread_progress.connect(self.receive_progress)
        self.thread_done.connect(self.parsing_done)
        self.error.connect(self.showinfo)

    def showinfo(self, txt, info):
        """ show info/error pop-up """
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Information)
        msg.setText(txt)
        msg.setInformativeText(info)
        msg.setWindowTitle("Feedback - PDFs Extractor")
        return msg.exec()

    def choose_dir(self):
        """ select folder for pdf parsing """
        folder = normpath(QFileDialog.getExistingDirectory(caption="Select folder with PDFs",
                                                           directory=self.lastopendir))
        if folder != ".":
            self.lastopendir = folder
            self.filenames = [e.name for e in scandir(folder) if e.name.endswith((".pdf", ".PDF"))]
            # add files to list view
            self.home.uploadPDFs(self.filenames)
            self.home.listview.setCurrentRow(0)
            self.preview_item(self.filenames[0])

    def preview_item(self, item):
        """ on listview item double-click """
        try:
            name = item.text()
        except AttributeError:  # if it's a str (called from choose_dir)
            name = item
        path = join(self.lastopendir, name)
        self.home.pdf_viewer.changeUrl(path)
        try:
            self.text_bbox = {text: bbox for text, bbox in getTextBbox(path)}
        except Exception as e:
            self.showinfo("The selected PDF may not be allowing text extraction", str(e))

    def addZone(self):
        """ add another to the existing zones """
        if not self.text_bbox:
            self.showinfo("PDF mapping could not be/has not been generated",
                          "Make sure:\n- a PDF file is selected for previewing\n- the previewed PDF is text-extractable")
            return

        parse_dialog = CustomDialog(self, self.text_bbox.keys())
        if parse_dialog.exec():
            # text
            self.zone_1_text = parse_dialog.zone_1.text()
            self.zone_2_text = parse_dialog.zone_2.text()
            self.zone_3_text = parse_dialog.zone_3.text()
            self.zone_4_text = parse_dialog.zone_4.text()
            # bbox
            self.zone_1_bbox = self.text_bbox.get(self.zone_1_text, ())
            self.zone_2_bbox = self.text_bbox.get(self.zone_2_text, ())
            self.zone_3_bbox = self.text_bbox.get(self.zone_3_text, ())
            self.zone_4_bbox = self.text_bbox.get(self.zone_4_text, ())

    def _process_zones(self):
        """ scan for similar texts """

        try:
            self.thread_started.emit()
            total_files = len(self.filenames)
            progress = 0
            # get the starting points
            self.row, self.col1, self.col2, self.col3, self.col4 = (
                int(self.home.row_edit.edit.text()),
                self.home.zoneA.edit.text(),
                self.home.zoneB.edit.text(),
                self.home.zoneC.edit.text(),
                self.home.zoneD.edit.text(),
            )
            writer = XLWriter()

            for index, name in enumerate(self.filenames):
                # pdf full path
                pdf = join(self.lastopendir, name)
                progress = (((index + 1) * 100) / total_files)
                self.thread_progress.emit(int(progress))
                # parse
                for text, bbox in getTextBbox(pdf):
                    # left, bottom, right, and top
                    if (bbox == self.zone_1_bbox):
                        # process zone A
                        writer.add_data(f"{self.col1}{self.row}", getInvoiceNum(text))

                    elif (bbox == self.zone_2_bbox):
                        # process zone B
                        writer.add_data(f"{self.col2}{self.row}", getEmail(text))

                    elif (bbox == self.zone_3_bbox):
                        # process zone C
                        writer.add_data(f"{self.col3}{self.row}", getName(text))

                    elif (bbox == self.zone_4_bbox):
                        # process zone D
                        writer.add_data(f"{self.col4}{self.row}", text.strip(), wrapText=True)
                # increment row count
                self.row += 1
                # check if cancelled
                if self.cancel_parsing:
                    self.thread_done.emit()
                    # break out of the function
                    return
            # finally write data to an Excel sheet
            output_file = f"{self.lastopendir}\\{basename(self.lastopendir)}.xlsx"
            writer.save(output_file)

        except Exception as e:
            self.error.emit("An error has interrupted the parsing of PDF files", str(e))

        self.thread_done.emit()

    def process_zones(self):
        """ add _process_zones to thread """
        if not self.filenames:
            self.showinfo("No PDF files have been selected for parsing",
                          "Click on 'Select folder' first before proceeding")
            return
        elif not (self.zone_1_bbox and self.zone_2_bbox and self.zone_3_bbox):
            self.showinfo("No parsing zones have been created",
                          "Create parsing zones on the top-right corner of the main window")
            return
        self.threadpool.start(self._process_zones)

    def parsing_started(self):
        """ when parsing starts """
        self.home.progressbar.show()
        self.home.stopbtn.show()
        self.home.parsebtn.setDisabled(True)

    def parsing_done(self):
        """ when parsing is done """
        self.home.progressbar.hide()
        self.home.stopbtn.hide()
        self.home.parsebtn.setEnabled(True)

    def stop_parsing(self):
        """ change flag """
        self.cancel_parsing = 1

    def receive_progress(self, value):
        """ update progress bar """
        self.home.progressbar.setValue(value)
        self.home.progressbar.setFormat(f"Parsing... {value}%")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLE)
    main = MainWin()
    main.showMaximized()

    sys.exit(app.exec())
