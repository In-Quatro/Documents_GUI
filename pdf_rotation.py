from datetime import datetime
import os
import csv
from PyPDF2 import PdfFileReader, PdfFileWriter
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal


class PdfRotation(QThread):
    status_update = pyqtSignal(str)
    progress_update = pyqtSignal(int)

    def __init__(self, path_input_pdf, path_output_pdf, stage):
        super().__init__()
        self.path_input_pdf = path_input_pdf
        self.path_output_pdf = path_output_pdf
        self.stage = stage
        # self.te_pdf = te_pdf

    def run(self):
        try:
            self.status_update.emit('Начинаю обработку, ожидайте..')
            self.page_rotation()
        except Exception as e:
            print(e)

    def page_rotation(self):
        """Поворот страниц кроме 1-ой."""
        stage = self.stage.value()
        quantity = 0
        files = os.listdir(self.path_input_pdf)
        if not files:
            self.status_update.emit('Нет файлов в папке для обработки')
        else:

            step = 100 / len(files)
            progress = 0
            self.progress_update.emit(progress)

            for file in files:
                if Path(file).suffix == '.pdf':
                    pdf_path = Path(self.path_input_pdf, file)
                    pdf_reader = PdfFileReader(pdf_path)
                    pdf_writer = PdfFileWriter()

                    for page in range(pdf_reader.getNumPages()):
                        pages = pdf_reader.getPage(page)
                        if page != 0:
                            pages.rotateClockwise(90)
                        pdf_writer.addPage(pages)

                    file_name = Path(file).stem
                    output_file_name = f'{file_name} - {stage}.pdf'
                    output_file_path = os.path.join(
                        self.path_output_pdf, output_file_name)

                    with open(output_file_path, 'wb') as output_file:
                        pdf_writer.write(output_file)

                    quantity += 1
                    self.write_csv(stage, file_name, pdf_reader)
                    progress += step
                    self.progress_update.emit(int(progress))
            self.progress_update.emit(100)
            self.status_update.emit(f'Готово. Обработано файлов - {quantity}')

    def write_csv(self, stage, file_name, pdf_reader):
        """Запись количества листов в csv документ."""
        csv_file = f'Количество листов (Этап {stage}).csv'
        csv_file_path = os.path.join(self.path_output_pdf, csv_file)
        date = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
        num_pages = pdf_reader.getNumPages()
        row = [file_name, file_name[4::], num_pages, date]
        # self.te_pdf.append(f'{file_name} - {num_pages}')

        # Проверка на существование файла и запись данных
        file_exists = os.path.isfile(csv_file_path)

        with open(csv_file_path, mode='a', newline='') as file:
            writer = csv.writer(file, delimiter=';')
            # Записываем заголовок, если файл новый
            if not file_exists:
                writer.writerow(['МО', 'МО', 'Страниц', 'Дата'])
            writer.writerow(row)
