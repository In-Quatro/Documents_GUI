import subprocess

from PyQt5 import QtCore, QtGui, QtWidgets, QtQuickWidgets, uic
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import (QApplication, QMainWindow,
                             QTableWidgetItem, QDialog, QFileDialog,
                              QPushButton)
import sys
from functools import partial
from PyPDF2 import PdfFileReader, PdfFileWriter

from acts_create import *
from acts_analysis import *


class MainWindow(QMainWindow):
    """"Главное окно."""
    def __init__(self):
        super(MainWindow, self).__init__()
        uic.loadUi(self.resource_path('open.ui'), self)
        self.input_folder = None
        self.output_folder = None
        self.output_excel_folder = None
        self.folder_template_excel = None
        self.csv_data = None

        self.path_acts = None
        self.path_output_csv_acts = None

        # Создание актов
        self.b_browse_template.clicked.connect(partial(
            self.get_file, 'template_excel', 'Excel Files (*.xlsx)'))
        self.b_browse_csv_data.clicked.connect(partial(
            self.get_file, 'csv_data', 'CSV Files (*.csv)'))
        self.b_browse_output.clicked.connect(partial(
            self.get_directory, 'output_excel'))
        self.b_start.clicked.connect(self.start_acts_create)
        self.b_open_folder.clicked.connect(self.open_folder)

        # Обработка актов
        self.b_browse_output_data.clicked.connect(partial(
            self.save_file, 'save_file',))
        self.b_browse_acts.clicked.connect(partial(
            self.get_directory, 'input_acts'))
        self.b_open_folder_data_acts.clicked.connect(self.open_folder_2)

        self.b_read_acts.clicked.connect(self.start_acts_analysis)

        self.actions = {
            'template_excel':
                ['folder_template_excel', self.le_path_template],
            'csv_data':
                ['csv_data', self.le_path_csv_data],
            'output_excel':
                ['output_excel_folder', self.le_path_output],
            'save_file':
                ['path_output_csv_acts', self.le_path_output_csv_acts],
            'input_acts':
                ['path_acts', self.le_path_acts_analysis],
        }

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

    def save_file(self, action):
        """Сохранить файл."""
        filename, ok = QFileDialog.getSaveFileName(self,
                                                  "Сохранить файл",
                                                  ".",
                                                  'CSV Files (*.csv)')
        if action in self.actions:
            setattr(self, self.actions[action][0], filename)
            self.actions[action][1].setText(filename)


    def page_rotation(self):
        """Поворот страниц кроме 1-ой."""
        self.statusbar.showMessage('Начинаю обработку...')
        stage_num = 4
        input_dir = self.input_folder
        output_dir = self.output_folder
        count = 0
        directory = os.listdir(input_dir)
        for file in directory:
            if Path(file).suffix == '.pdf':
                pdf_path = Path(input_dir, file)
                pdf_reader = PdfFileReader(pdf_path)
                pdf_writer = PdfFileWriter()

            for page in range(pdf_reader.getNumPages()):
                pages = pdf_reader.getPage(page)
                if page != 0:
                    pages.rotateClockwise(90)
                pdf_writer.addPage(pages)

            file_name = Path(file).stem
            output_file_name = f'{file_name} - {stage_num}.pdf'
            output_file_path = os.path.join(output_dir, output_file_name)
            with open(output_file_path, 'wb') as output_file:
                pdf_writer.write(output_file)
            count += 1
        self.statusbar.showMessage(f'Готово. Обработано файлов - {count}')

    def update_progress(self, value):
        """Прогресбар."""
        self.progressBar.setValue(value)

    def open_folder(self):
        """Открытие папки."""
        if self.output_excel_folder:
            normalized_path = os.path.normpath(self.output_excel_folder)
            subprocess.Popen(['explorer', normalized_path])
        else:
            self.update_status('Необходимо выбрать папку')

    def open_folder_2(self):
        """Открытие папки."""
        if self.path_output_csv_acts:
            path_csv = Path(self.path_output_csv_acts).parent
            subprocess.Popen(['explorer', path_csv])
        else:
            self.update_status('Необходимо выбрать папку')

    def start_acts_create(self):
        """Запуск создания актов EXCEL."""
        template_path = str(self.folder_template_excel)
        folder_name = str(self.output_excel_folder)
        csv_data = str(self.csv_data)

        if all([
            self.folder_template_excel,
            self.output_excel_folder,
            self.csv_data
        ]):
            self.thread = ActsCreate(template_path, folder_name, csv_data)
            self.thread.status_update.connect(self.update_status)
            self.thread.progress_update.connect(self.update_progress)
            self.thread.start()
        else:
            self.update_status('Проверьте что выбраны все поля')

    def start_acts_analysis(self):
        """Запуск обработки актов EXCEL."""
        path_acts = str(self.path_acts)
        path_output_csv_acts = str(self.path_output_csv_acts)
        stage = self.cb_stage
        try:
            if all([
                self.path_acts,
                self.path_output_csv_acts,
            ]):
                # print('YES')
                self.thread = ActsAnalysis(
                    path_acts, path_output_csv_acts, stage)
                self.thread.status_update.connect(self.update_status)
                self.thread.progress_update.connect(self.update_progress)
                self.thread.start()
            else:
                # print('NO')
                self.update_status('Проверьте что выбраны все поля')
        except Exception as e:
            print(e)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())