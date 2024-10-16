import subprocess
import sys
from functools import partial

from PyQt5 import uic, QtGui
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QDialog

from modules.acts_analysis import *
from modules.acts_create import *
from modules.acts_incidents import *
from modules.pdf_rotation import *
from modules.title_page import *
from modules.title_page_analysis import *


class MainWindow(QMainWindow):
    """"Главное окно."""

    def __init__(self):
        super(MainWindow, self).__init__()
        uic.loadUi(self.resource_path('ui/main.ui'), self)
        self.dialog = Dialog_acts()
        self.setWindowIcon(
            QtGui.QIcon(self.resource_path("ui/icons/logo.ico")))

        # Кнопки раздела "СОЗДАТЬ АКТЫ"
        self.b_browse_template_acts.clicked.connect(partial(
            self.get_file, self.le_path_template_acts, 'Excel Files (*.xlsx)'))
        self.b_browse_csv_data_acts.clicked.connect(partial(
            self.get_file, self.le_path_csv_data_acts, 'CSV Files (*.csv)'))
        self.b_browse_output_acts.clicked.connect(partial(
            self.get_directory, self.le_path_output_acts))
        self.b_create_acts.clicked.connect(self.start_acts_create)
        self.b_open_folder_acts.clicked.connect(partial(
            self.open_folder, self.le_path_output_acts))

        self.b_other_menu_acts.clicked.connect(self.open_dialog)

        # Кнопки раздела "ПОЛУЧИТЬ ДАННЫЕ ИЗ АКТОВ"
        self.b_browse_input_acts.clicked.connect(partial(
            self.get_directory, self.le_path_analysis_acts))
        self.b_browse_output_csv_data_acts.clicked.connect(partial(
            self.save_file, self.le_path_output_csv_data_acts))
        self.b_analysis_acts.clicked.connect(self.start_acts_analysis)
        self.b_open_folder_csv_data_acts.clicked.connect(partial(
            self.open_folder, self.le_path_output_csv_data_acts, True))

        # Кнопки раздела "СОЗДАТЬ ТИТУЛЬНЫЕ ЛИСТЫ"
        self.b_browse_template_title_page.clicked.connect(partial(
            self.get_file,
            self.le_path_template_title_page,
            'Word Files (*.docx)'))
        self.b_browse_csv_data_title_page.clicked.connect(partial(
            self.get_file,
            self.le_path_csv_data_title_page,
            'CSV Files (*.csv)'))
        self.b_browse_output_title_page.clicked.connect(partial(
            self.get_directory,
            self.le_path_output_title_page))
        self.b_create_title_page.clicked.connect(self.start_title_page_create)
        self.b_open_folder_title_page.clicked.connect(partial(
            self.open_folder,
            self.le_path_output_title_page))

        # Кнопки раздела "ПОЛУЧИТЬ ДАННЫЕ ИЗ ТИТУЛЬНЫХ ЛИСТОВ"
        self.b_browse_input_title_page.clicked.connect(partial(
            self.get_directory, self.le_path_analysis_title_page))
        self.b_browse_output_csv_data_title_page.clicked.connect(partial(
            self.save_file, self.le_path_output_csv_data_title_page))
        self.b_analysis_title_page.clicked.connect(
            self.start_title_page_analysis)
        self.b_open_folder_csv_data_title_page.clicked.connect(partial(
            self.open_folder, self.le_path_output_csv_data_title_page, True))

        # Кнопки раздела "ВНЕСТИ ЗАЯВКИ В АКТЫ"
        self.b_browse_input_acts_incidents.clicked.connect(partial(
            self.get_directory, self.le_path_input_acts_incidents))
        self.b_browse_csv_data_incidents.clicked.connect(partial(
            self.get_file,
            self.le_path_csv_data_incidents,
            'CSV Files (*.csv)'))
        self.b_browse_output_acts_incidents.clicked.connect(partial(
            self.get_directory, self.le_path_output_acts_incidents))
        self.b_insert_incidents.clicked.connect(
            self.start_acts_incidents)
        self.b_open_folder_acts_incidents.clicked.connect(partial(
            self.open_folder, self.le_path_output_acts_incidents))

        # Кнопки раздела "ОБРАБОТАТЬ PDF"
        self.b_browse_input_pdf.clicked.connect(
            partial(self.get_directory, self.le_path_input_pdf))
        self.b_browse_output_pdf.clicked.connect(
            partial(self.get_directory, self.le_path_output_pdf))
        self.b_rotate_pdf.clicked.connect(self.start_pdf_rotation)
        self.b_open_folder_new_pdf.clicked.connect(partial(
            self.open_folder, self.le_path_output_pdf))

    def check_button_state(self, button, conditions):
        """Проверка кнопки для включения."""
        button.setEnabled(all(conditions))

    def check_buttons(self):
        """Подготовка данных кнопок для проверки перед включением."""
        # Включение кнопок раздела "ОСОЗДАТЬ АКТЫ"
        self.check_button_state(self.b_create_acts, [
            self.le_path_template_acts.text(),
            self.le_path_csv_data_acts.text(),
            self.le_path_output_acts.text(),
            self.dialog.te_post.toPlainText(),
            self.dialog.le_fio.text(),
        ])

        self.check_button_state(self.b_open_folder_acts, [
            self.le_path_output_acts.text(),
        ])

        # Включение кнопок раздела "ПОЛУЧИТЬ ДАННЫЕ ИЗ АКТОВ"
        self.check_button_state(self.b_analysis_acts, [
            self.le_path_analysis_acts.text(),
            self.le_path_output_csv_data_acts.text(),
        ])

        self.check_button_state(self.b_open_folder_csv_data_acts, [
            self.le_path_output_csv_data_acts.text(),
        ])

        # Включение кнопок раздела "СОЗДАТЬ ТИТУЛЬНЫЕ ЛИСТЫ"
        self.check_button_state(self.b_create_title_page, [
            self.le_path_template_title_page.text(),
            self.le_path_csv_data_title_page.text(),
            self.le_path_output_title_page.text(),
        ])

        self.check_button_state(self.b_open_folder_title_page, [
            self.le_path_output_title_page.text(),
        ])

        # Включение кнопок раздела "ПОЛУЧИТЬ ДАННЫЕ ИЗ ТИТУЛЬНЫХ ЛИСТОВ"
        self.check_button_state(self.b_analysis_title_page, [
            self.le_path_analysis_title_page.text(),
            self.le_path_output_csv_data_title_page.text(),
        ])

        self.check_button_state(self.b_open_folder_csv_data_title_page, [
            self.le_path_output_csv_data_title_page.text(),
        ])

        # Включение кнопок раздела "ВНЕСТИ ЗАЯВКИ В АКТЫ"
        self.check_button_state(self.b_insert_incidents, [
            self.le_path_input_acts_incidents.text(),
            self.le_path_csv_data_incidents.text(),
            self.le_path_output_acts_incidents.text(),
        ])

        self.check_button_state(self.b_open_folder_acts_incidents, [
            self.le_path_output_acts_incidents.text(),
        ])

        # Включение кнопок раздела "ОБРАБОТАТЬ PDF"
        self.check_button_state(self.b_rotate_pdf, [
            self.le_path_input_pdf.text(),
            self.le_path_output_pdf.text(),
        ])

        self.check_button_state(self.b_open_folder_new_pdf, [
            self.le_path_output_pdf.text(),
        ])

    def update_status(self, msg):
        """Изменение сообщения статусбара."""
        self.statusbar.showMessage(msg)
        self.statusbar.repaint()

    @staticmethod
    def resource_path(relative_path):
        """Получает абсолютный путь к ресурсу,
        работает как в режиме разработки, так и в скомпилированном виде."""
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def get_directory(self, action):
        """Выбор папки."""
        path = QFileDialog.getExistingDirectory(self, "Выбрать папку", ".")
        action.setText(path)
        self.check_buttons()

    def get_file(self, action, format_file):
        """Выбор файла."""
        try:
            filename, filetype = QFileDialog.getOpenFileName(
                self,
                f"Выбрать файл",
                ".",
                f"{format_file}"
            )
            if action:
                action.setText(filename)
                self.check_buttons()
        except Exception as e:
            print(e)

    def save_file(self, action):
        """Выбрать путь для сохранения файла."""
        filename, ok = QFileDialog.getSaveFileName(self,
                                                   "Сохранить как",
                                                   ".",
                                                   'CSV Files (*.csv)')
        action.setText(filename)
        self.check_buttons()

    def update_progress(self, value):
        """Прогресбар."""
        self.progressBar.setValue(value)

    def open_folder(self, folder, parent=False):
        """Открытие конечной папки."""
        if folder:
            folder = Path(folder.text())
            if parent:
                folder = folder.parent
            subprocess.Popen(['explorer', folder])
        else:
            self.update_status('Необходимо выбрать папку')

    def start_acts_create(self):
        """Запуск создания Актов _xlsx_."""
        path_template_acts = Path(
            self.le_path_template_acts.text()
        )
        path_csv_data_acts = Path(
            self.le_path_csv_data_acts.text()
        )
        path_output_acts = Path(
            self.le_path_output_acts.text()
        )
        dialog = self.dialog.get_data()
        self.thread = ActsCreate(
            path_template_acts,
            path_csv_data_acts,
            path_output_acts,
            dialog
        )
        self.thread.status_update.connect(self.update_status)
        self.thread.progress_update.connect(self.update_progress)
        self.thread.start()

    def start_acts_analysis(self):
        """Запуск обработки Актов _xlsx_ в _csv_."""
        path_acts = Path(
            self.le_path_analysis_acts.text()
        )
        path_output_csv_acts = Path(
            self.le_path_output_csv_data_acts.text()
        )
        stage = self.cb_stage_acts

        self.thread = ActsAnalysis(
            path_acts,
            path_output_csv_acts,
            stage
        )
        self.thread.status_update.connect(self.update_status)
        self.thread.progress_update.connect(self.update_progress)
        self.thread.start()

    def start_title_page_create(self):
        """Запуск создания Титульных листов _docx_."""
        path_template_title_page = Path(
            self.le_path_template_title_page.text()
        )
        path_csv_data_title_page = Path(
            self.le_path_csv_data_title_page.text()
        )
        path_output_title_page = Path(
            self.le_path_output_title_page.text()
        )

        self.thread = TitlePageCreate(
            path_template_title_page,
            path_csv_data_title_page,
            path_output_title_page
        )
        self.thread.status_update.connect(self.update_status)
        self.thread.progress_update.connect(self.update_progress)
        self.thread.start()

    def start_title_page_analysis(self):
        """Запуск обработки Титульных листов _docx_ в _csv_."""
        path_analysis_title_page = Path(
            self.le_path_analysis_title_page.text()
        )
        path_output_csv_data_title_page = Path(
            self.le_path_output_csv_data_title_page.text()
        )

        self.thread = TitlePageAnalysis(path_analysis_title_page,
                                        path_output_csv_data_title_page)
        self.thread.status_update.connect(self.update_status)
        self.thread.progress_update.connect(self.update_progress)
        self.thread.start()

    def start_acts_incidents(self):
        """Запуск внесение заявок в Акты _xlsx_ из _csv_."""
        path_input_acts_incidents = Path(
            self.le_path_input_acts_incidents.text()
        )
        path_csv_data_incidents = Path(
            self.le_path_csv_data_incidents.text()
        )
        path_output_acts_incidents = Path(
            self.le_path_output_acts_incidents.text()
        )

        self.thread = ActIncident(
            path_input_acts_incidents,
            path_csv_data_incidents,
            path_output_acts_incidents
        )
        self.thread.status_update.connect(self.update_status)
        self.thread.progress_update.connect(self.update_progress)
        self.thread.start()

    def start_pdf_rotation(self):
        """Запуск обработки PDF документов."""
        path_input_pdf = Path(self.le_path_input_pdf.text())
        path_output_pdf = Path(self.le_path_output_pdf.text())
        stage = self.sb_stage_pdf

        self.thread = PdfRotation(
            path_input_pdf,
            path_output_pdf,
            stage,
        )
        self.thread.status_update.connect(self.update_status)
        self.thread.progress_update.connect(self.update_progress)
        self.thread.start()

    def open_dialog(self):
        self.dialog.exec_()
        self.check_buttons()


class Dialog_acts(QDialog):
    """Диалоговое окно."""

    def __init__(self):
        super(Dialog_acts, self).__init__()
        uic.loadUi(self.resource_path(r'ui\dialog_acts.ui'), self)
        self.setWindowIcon(
            QtGui.QIcon(self.resource_path("ui/icons/logo.ico")))

    def get_data(self):
        self.post = self.te_post.toPlainText().replace(
            '"', '«', 1).replace('"', '»', 2)
        self.fio = self.le_fio.text()
        return self.post, self.fio

    @staticmethod
    def resource_path(relative_path):
        """Получает абсолютный путь к ресурсу,
        работает как в режиме разработки, так и в скомпилированном виде."""
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
