import subprocess

from PyQt5 import uic

from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog
import sys
from functools import partial

from acts_create import *
from acts_analysis import *
from pdf_rotation import *


class MainWindow(QMainWindow):
    """"Главное окно."""
    def __init__(self):
        super(MainWindow, self).__init__()
        uic.loadUi(self.resource_path('ui/open.ui'), self)
        # Создание актов
        self.path_template_acts = None
        self.path_csv_data_acts = None
        self.path_output_acts = None
        self.b_create_acts.setEnabled(False)

        self.b_browse_template_acts.clicked.connect(partial(
            self.get_file, 'template_acts', 'Excel Files (*.xlsx)'))
        self.b_browse_csv_data_acts.clicked.connect(partial(
            self.get_file, 'csv_data_acts', 'CSV Files (*.csv)'))
        self.b_browse_output_acts.clicked.connect(partial(
            self.get_directory, 'output_acts'))

        # Кнопки для Создания актов
        self.b_create_acts.clicked.connect(self.start_acts_create)
        self.b_open_folder.clicked.connect(self.open_folder)

        # Обработка актов
        self.path_analysis_acts = None
        self.path_output_csv_data_acts = None
        self.b_analysis_acts.setEnabled(False)

        self.b_browse_input_acts.clicked.connect(partial(
            self.get_directory, 'save_csv_file_acts',))
        self.b_browse_output_csv_data_acts.clicked.connect(partial(
            self.save_file, 'save_csv_file_folder_acts'))

        # Кнопки для сохранения данных из актов _xlsx_ файлов в _csv_ файл
        self.b_open_folder_csv_data_acts.clicked.connect(self.open_folder_2)
        self.b_analysis_acts.clicked.connect(self.start_acts_analysis)

        # Обработка PDF
        self.path_input_pdf: str = None
        self.path_output_pdf: str = None
        self.b_rotate_pdf.setEnabled(False)

        self.b_browse_input_pdf.clicked.connect(partial(
            self.get_directory, 'input_pdf'))
        self.b_browse_output_pdf.clicked.connect(partial(
            self.get_directory, 'output_pdf'))

        # Кнопки для обработки _pdf_ файлов
        self.b_rotate_pdf.clicked.connect(self.start_pdf_rotation)
        self.b_open_folder_new_pdf.clicked.connect(self.open_folder_3)

        self.b_clear_te_pdf.clicked.connect(self.clear_te_pdf)

        self.actions = {
            #  Создание актов _xlsx_ из _csv_ файла с данными
            'template_acts':
                ['path_template_acts',
                 self.le_path_template_acts],
            'csv_data_acts':
                ['path_csv_data_acts',
                 self.le_path_csv_data_acts],
            'output_acts':
                ['path_output_acts',
                 self.le_path_output_acts],

            # Сохранения данных из актов _xlsx_ файлов в _csv_ файл
            'save_csv_file_acts':
                ['path_analysis_acts',
                 self.le_path_analysis_acts],
            'save_csv_file_folder_acts':
                ['path_output_csv_data_acts',
                 self.le_path_output_csv_data_acts],

            # Обработка файлов _pdf_
            'input_pdf':
                ['path_input_pdf',
                 self.le_path_input_pdf],
            'output_pdf':
                ['path_output_pdf',
                 self.le_path_output_pdf]
        }

    # def check_button_create_acts_state(self):
    #     self.b_create_acts.setEnabled(
    #         all([
    #             self.path_template_acts,
    #             self.path_csv_data_acts,
    #             self.path_output_acts
    #         ])
    #     )
    #
    # def check_button_read_acts_state(self):
    #     self.b_analysis_acts.setEnabled(
    #         all([
    #             self.path_analysis_acts,
    #             self.path_output_csv_data_acts,
    #         ])
    #     )
    #
    # def check_button_rotate_pdf_state(self):
    #     self.b_rotate_pdf.setEnabled(
    #         all([
    #             self.path_input_pdf,
    #             self.path_output_pdf,
    #         ])
    #     )

    def check_button_state(self, button, conditions):
        button.setEnabled(all(conditions))

    def check_buttons(self):
        self.check_button_state(self.b_create_acts, [
            self.path_template_acts,
            self.path_csv_data_acts,
            self.path_output_acts
        ])

        self.check_button_state(self.b_analysis_acts, [
            self.path_analysis_acts,
            self.path_output_csv_data_acts,
        ])

        self.check_button_state(self.b_rotate_pdf, [
            self.path_input_pdf,
            self.path_output_pdf,
        ])

    def clear_te_pdf(self):
        self.te_pdf.clear()

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
        dirlist = QFileDialog.getExistingDirectory(self, "Выбрать папку", ".")
        if action in self.actions:
            setattr(self, self.actions[action][0], dirlist)
            self.actions[action][1].setText(dirlist)
            self.check_buttons()

    def get_file(self, action, format_file):
        """Выбор файла."""
        filename, filetype = QFileDialog.getOpenFileName(
            self,
            f"Выбрать файл",
            ".",
            f"{format_file}"
        )
        if action in self.actions:
            setattr(self, self.actions[action][0], filename)
            self.actions[action][1].setText(filename)
            self.check_buttons()

    def save_file(self, action):
        """Сохранить файл."""
        filename, ok = QFileDialog.getSaveFileName(self,
                                                   "Сохранить файл",
                                                   ".",
                                                   'CSV Files (*.csv)')
        if action in self.actions:
            setattr(self, self.actions[action][0], filename)
            self.actions[action][1].setText(filename)
            self.check_buttons()

    def update_progress(self, value):
        """Прогресбар."""
        self.progressBar.setValue(value)

    def open_folder(self):
        """Открытие папки созданных актов."""
        if self.path_output_acts:
            normalized_path = Path(self.path_output_acts)
            subprocess.Popen(['explorer', normalized_path])
        else:
            self.update_status('Необходимо выбрать папку')

    def open_folder_2(self):
        """Открытие папки."""
        if self.path_output_csv_data_acts:
            path_csv = Path(self.path_output_csv_data_acts).parent
            subprocess.Popen(['explorer', path_csv])
        else:
            self.update_status('Необходимо выбрать папку')

    def open_folder_3(self):
        """Открытие папки."""
        if self.path_output_pdf:
            path_pdf = Path(self.path_output_pdf)
            subprocess.Popen(['explorer', path_pdf])
        else:
            self.update_status('Необходимо выбрать папку')

    def start_acts_create(self):
        """Запуск создания актов EXCEL."""
        template_path = str(self.path_template_acts)
        csv_data = str(self.path_csv_data_acts)
        folder_name = str(self.path_output_acts)

        self.thread = ActsCreate(template_path, folder_name, csv_data)
        self.thread.status_update.connect(self.update_status)
        self.thread.progress_update.connect(self.update_progress)
        self.thread.start()

    def start_acts_analysis(self):
        """Запуск обработки актов EXCEL."""
        path_acts = str(self.path_analysis_acts)
        path_output_csv_acts = str(self.path_output_csv_data_acts)
        stage = self.cb_stage_acts

        self.thread = ActsAnalysis(
            path_acts, path_output_csv_acts, stage)
        self.thread.status_update.connect(self.update_status)
        self.thread.progress_update.connect(self.update_progress)
        self.thread.start()

    def start_pdf_rotation(self):
        """Запуск обработки PDF документов."""
        path_input_pdf = str(self.path_input_pdf)
        path_output_pdf = str(self.path_output_pdf)
        stage = self.sb_stage_pdf
        te_pdf = self.te_pdf

        self.thread = PdfRotation(path_input_pdf, path_output_pdf, stage, te_pdf)
        self.thread.status_update.connect(self.update_status)
        self.thread.progress_update.connect(self.update_progress)
        self.thread.start()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())