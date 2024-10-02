import os
import re
import csv

from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal
from docx import Document


class TitlePageAnalysis(QThread):
    status_update = pyqtSignal(str)
    progress_update = pyqtSignal(int)

    def __init__(self, path_analysis_title_page,
                 path_output_csv_data_title_page):
        super().__init__()
        self.path_analysis_title_page = path_analysis_title_page
        self.path_output_csv_data_title_page = path_output_csv_data_title_page
        self.header = [
            'Код МО',
            'Код полностью',
            'Наименование МО',
            'Должность МО',
            'ФИО',
            'Основание подписания'
        ]
        self.patterns = {
            'kod': r'\Технический акт №\xa0(.*?)\n',
            'kod_2': r'\» \((.*?)\)',
            'naimenovanie': r'\, и (.*?)\s\(',
            'dolzhnost': r'\«МО», в лице (.*?)\s[А-ЯЁ]',
            'fio': r'[А-ЯЁ][а-яё]+\s[А-ЯЁ][а-яё]+\s[А-ЯЁ][а-яё]+',
            'osnovanie': r'\основании (.*?)\,',
        }

    def run(self):
        try:
            files = os.listdir(self.path_analysis_title_page)
            if not files:
                self.status_update.emit('Нет файлов в папке для обработки')
            else:
                self.status_update.emit(
                    'Идет обработка титульных листов, ожидайте...')
                step = 100 / len(files)
                progress = 0
                quantity = 0
                for file in files:
                    quantity += 1
                    progress += step
                    self.progress_update.emit(int(progress))
                    if Path(file).suffix == '.docx':
                        file_path = Path(self.path_analysis_title_page, file)
                        doc = Document(file_path)
                        text = '\n\n'.join(
                            [par.text for par in doc.paragraphs])
                        data = [re.findall(match, text)[-1] for match in
                                self.patterns.values()]
                        self.write_to_csv(data)
                self.progress_update.emit(100)
                self.status_update.emit(
                    f'Готово. Обработано файлов: {quantity}')
        except Exception as e:
            print(e)

    def write_to_csv(self, data):
        """Создание CSV файла и внесение данных."""
        file_exists = os.path.isfile(self.path_output_csv_data_title_page)
        with open(self.path_output_csv_data_title_page, mode='a',
                  newline='') as file:
            writer = csv.writer(file, delimiter=';')
            if not file_exists:
                writer.writerow(self.header)
            writer.writerow(data)
