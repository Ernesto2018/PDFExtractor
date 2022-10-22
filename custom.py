__author__ = "Ernesto"
__email__ = "ernestondieki12@gmail.com"
__dated__ = "14-09-22"  # DDMMYY

from PyQt5.QtWidgets import (
    QCompleter, QComboBox,
    QDialog, QDialogButtonBox,
    QVBoxLayout, QFormLayout,
    QLabel, QLineEdit,
    QAction,
)
from PyQt5.QtCore import (
    QSortFilterProxyModel,
    QStringListModel, Qt,
    QRegExp,
)
from PyQt5.QtGui import QIcon
from os.path import abspath, dirname, join
import sys

BASE_DIR = dirname(abspath(__file__))


def r_path(relpath):
    """ get absolute path """

    base_path = getattr(sys, "_MEIPASS", BASE_DIR)
    return join(base_path, relpath)


class Completer(QCompleter):
    """ custom line edit completer """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.local_prefix = ""
        self.source_model = None
        self.filter_model = QSortFilterProxyModel(self)
        self.using_original_model = False

    def setModel(self, model):
        if isinstance(model, (list, tuple, set)):
            model = QStringListModel(model)
        self.source_model = model
        self.filter_model = QSortFilterProxyModel(self)
        self.filter_model.setSourceModel(self.source_model)
        super().setModel(self.filter_model)
        self.using_original_model = True

    def updateModel(self):
        if not self.using_original_model:
            self.filter_model.setSourceModel(self.source_model)

        pattern = QRegExp(self.local_prefix, Qt.CaseSensitivity.CaseInsensitive,
                          QRegExp.PatternSyntax.FixedString)
        self.filter_model.setFilterRegExp(pattern)

    def splitPath(self, path):
        self.local_prefix = path
        self.updateModel()
        if self.filter_model.rowCount() == 0:
            self.using_original_model = False
            self.filter_model.setSourceModel(QStringListModel([path]))
            return [path]
        return []


class ParseEdit(QComboBox):
    """ parser line edit """

    __slots__ = ("comp", "prev_text", "word_suggestions", "edit", "error_action")

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.setEditable(True)
        self.setFixedWidth(600)
        self.editTextChanged.connect(self.validate)

        self.edit = self.lineEdit()
        self.word_suggestions = set()
        self.prev_text = ""
        self.error_action = QAction()
        self.error_action.setIcon(QIcon(r_path(("error.png"))))

        self.comp = Completer(self)
        self.comp.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
        self.setCompleter(self.comp)

    def text(self):
        """ raise error on an empty string else return """
        text = super().currentText()
        if not text:
            self.error_action.setToolTip("This field is required")
            self.edit.addAction(self.error_action, QLineEdit.ActionPosition.LeadingPosition)
            return ""
        return text

    def drop_error(self):
        """ remove error QAction from line edit """
        self.edit.removeAction(self.error_action)

    def validate(self, text):
        """ raise or drop error icon accordingly """
        if not text:
            self.drop_error()
        elif text not in self.word_suggestions:
            self.error_action.setToolTip("No match found in the PDF")
            self.edit.addAction(self.error_action, QLineEdit.ActionPosition.LeadingPosition)
        else:
            self.drop_error()

    def add_words(self, words):
        """ add unique parse words """
        # if words not in self.word_suggestions:
        self.word_suggestions = set(words)
        self.addItems(words)
        self.comp.setModel(self.model())

    def remove_word(self, name):
        """ drop parse word """
        if name in self.word_suggestions:
            self.word_suggestions.remove(name)
            index = self.findText(name)
            self.removeItem(index)

    def focusInEvent(self, e):
        current_text = self.currentText()
        if current_text:
            self.prev_text = current_text
        # self.clearEditText()
        super().focusInEvent(e)

    def focusOutEvent(self, e):
        current_text = self.currentText()
        if not current_text:
            self.drop_error()
            self.setEditText(self.prev_text)
        super().focusOutEvent(e)

    def keyPressEvent(self, e):
        key = e.key()
        if (key == Qt.Key.Key_Return) or (key == Qt.Key.Key_Enter):
            # avoid setting an empty str
            self.comp.popup().hide()
            text = self.comp.currentCompletion()
            self.setEditText(text)
        return super().keyPressEvent(e)


class CustomDialog(QDialog):
    def __init__(self, parent, suggestion_data, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)

        self.setWindowTitle("Create Parsing Zone")

        QBtn = QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel

        self.buttonBox = QDialogButtonBox(QBtn)
        self.buttonBox.accepted.connect(self.okay)
        self.buttonBox.rejected.connect(self.cancel)

        self.mainlayout = QVBoxLayout(self)
        form_layout = QFormLayout()

        message = QLabel("Type the text for the zones you want to be extracted from the PDFs\nThen select from the pop-up suggestions")
        message.setObjectName("popupMsg")
        message.setAlignment(Qt.AlignmentFlag.AlignCenter)
        # zone # 1
        self.zone_1 = ParseEdit()
        self.zone_1.add_words(suggestion_data)
        self.zone_1.setToolTip("Zone # 1 (Invoice number)")
        # add to form layout
        form_layout.addRow("Zone # 1:", self.zone_1)
        # zone # 2
        self.zone_2 = ParseEdit()
        self.zone_2.add_words(suggestion_data)
        self.zone_2.setToolTip("Zone # 2 (Email address)")
        # add to form layout
        form_layout.addRow("Zone # 2:", self.zone_2)
        # zone # 3
        self.zone_3 = ParseEdit()
        self.zone_3.add_words(suggestion_data)
        self.zone_3.setToolTip("Zone # 3 (Names)")
        # add to form layout
        form_layout.addRow("Zone # 3:", self.zone_3)
        # zone # 4
        self.zone_4 = ParseEdit()
        self.zone_4.add_words(suggestion_data)
        self.zone_4.setToolTip("Zone # 4 (Billing address)")
        # add to form layout
        form_layout.addRow("Zone # 4:", self.zone_4)

        self.mainlayout.addWidget(message)
        self.mainlayout.addLayout(form_layout)
        self.mainlayout.addWidget(self.buttonBox)

    def okay(self):
        """ OK clicked """
        z1, z2, z3, z4 = self.zone_1.text(), self.zone_2.text(), self.zone_3.text(), self.zone_4.text()
        if all((z1, z2, z3, z4)):
            self.accept()

    def cancel(self):
        self.reject()
