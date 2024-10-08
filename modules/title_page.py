import csv
import re
from pathlib import Path

from docx import Document
from PyQt5.QtCore import QThread, pyqtSignal


class TitlePageCreate(QThread):
    """Создание Титульных листов _docx_ из _csv_ файла с данными."""
    status_update = pyqtSignal(str)
    progress_update = pyqtSignal(int)

    def __init__(self,
                 path_template_title_page,
                 path_csv_data_title_page,
                 path_output_title_page):
        super().__init__()
        self.path_template_title_page = path_template_title_page
        self.path_csv_data_title_page = path_csv_data_title_page
        self.path_output_title_page = path_output_title_page

    def run(self):
        try:
            self.status_update.emit(
                'Идет создание титульных листов, ожидайте...')
            self.read_csv()
        except Exception as e:
            print(e)

    def get_step(self):
        with open(self.path_csv_data_title_page, newline='') as csvfile:
            reader = csv.DictReader(csvfile, delimiter=';')
            row_count = sum(1 for _ in reader)
            cnt = 100 / row_count
            return cnt

    def read_csv(self):
        step = self.get_step()
        progress = 0
        quantity = 0
        self.progress_update.emit(progress)
        with open(self.path_csv_data_title_page, encoding='ANSI') as csv_file:
            file_reader = csv.DictReader(csv_file, delimiter=";")
            for row in file_reader:
                progress += step
                self.progress_update.emit(int(progress))
                kod_mo = row['kod']
                fio = row['client']
                post = row['position']
                file_name = f"{kod_mo}_{post}_{fio}.docx"
                self.fill_docx_template(row, file_name)
                quantity += 1
            self.progress_update.emit(100)
            self.status_update.emit(f'Готово. Создано файлов: {quantity}')

    def fill_docx_template(self, data, file_name):
        """Создание docx на основе шаблона."""
        doc = Document(self.path_template_title_page)

        for paragraph in doc.paragraphs:
            for key, value in data.items():
                if key in paragraph.text:
                    inline = paragraph.runs
                    for i in range(len(inline)):
                        if key in inline[i].text:
                            inline[i].text = re.sub(re.escape(key), value,
                                                    inline[i].text)

        with Path(self.path_output_title_page, file_name) as output_file:
            doc.save(output_file)
