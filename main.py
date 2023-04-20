"""
App name: Skarb - profit converter
License: GNU GPL v.3

J1703502.xsd XML scheme converter for personal profit analysis

Description:
Обробка та зведення загальних показників експортованих таблиць доходів. Придатні до
опрацювання файли формату *.XML (файл формату PDF не опрацьовується). Власноручне
внесення змін до файлу або збереження формату сторонніми програмами може призвести
до унеможливлення конвертування.

(с) 2023 https://github.com/OlehOleinikov/Skarb

Used in GUI:
https://www.flaticon.com/free-icons/excel - Excel icons created by Freepik - Flaticon
https://www.flaticon.com/free-icons/microsoft-word - Microsoft word icons created by Bharat Icons - Flaticon
https://www.flaticon.com/free-icons/excel - Excel icons created by Bharat Icons - Flaticon
"""

import sys
import os
from pathlib import Path

from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox, QProgressBar

from gui.main_gui import Ui_MainWindow
from xml_converter import FileProfitXML, MultiFileDrfoData
from word_reporter import DocEditor


class AppWin(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.data = MultiFileDrfoData()

        self.b_import.clicked.connect(self.import_file)
        self.b_word.clicked.connect(self.save_word)
        self.b_excel.clicked.connect(self.save_excel)
        self.l_cur_file.setText(f'Статус: Готовий до роботи')
        self.label_4.setOpenExternalLinks(True)  # дозвіл на відкриття браузера (посилання на Github)

    def import_file(self):
        """Вибір файлів для опрацювання (відкриття вікна вибору, валідація, препроцесінг)"""

        result_info = ''  # звіт про результати (накопичується під час виконання)
        user_files = QFileDialog.getOpenFileNames(self, 'Додати файл (файли) для опрацювання (*.xml)',
                                                  str(Path.cwd().absolute()),
                                                  'Файли ДРФО (*.xml)')
        # Якщо користувач не обрав файли:
        if not len(user_files[0]):
            self.statusbar.showMessage('Не обрані файли XML...', 5000)
            return 0

        # Створення тимчасового прогресбару
        self.statusbar.showMessage('Завантаження XML...', 5000)
        self.progressBar = QProgressBar()
        self.progressBar.setMaximumHeight(18)
        self.progressBar.setMinimumHeight(18)
        self.progressBar.setStyleSheet("""QProgressBar {
                                                border: 2px solid rgb(211, 211, 211);
                                                border-radius: 7px;
                                                background-color: rgb(211, 211, 211);
                                                text-align: center;
                                            }
                                            QProgressBar::chunk {
                                                background-color: rgb(246, 191, 39);
                                                width: 7px; 
                                                margin: 0.5px;
                                                border-radius :2px;
                                            }""")
        self.statusBar().addPermanentWidget(self.progressBar)
        self.progressBar.setMaximum(len(user_files[0]) - 1)
        self.progressBar.setMinimum(0)
        self.progressBar.setValue(0)

        # Ітерування файлів:
        for pos, file in enumerate(user_files[0]):
            file_path = file
            file_name = os.path.basename(file)
            result_info += f"----------------------------------\n" \
                           f"File: {file_name}\n"
            cur_df = FileProfitXML(file_path)  # створення інстансу обробки одного файлу
            success = (not bool(cur_df.read_xml()))  # спроба прочитати XML файл (доступ, збір відомих тегів)
            if not success:
                result_info += f'Помилка читання файлу: можливо файл відкритий іншою програмою або не є файлом ДРФО\n\n'
                continue
            warnings = cur_df.fill_df()  # валідація та форматування даних файлу
            result_info += warnings

            if cur_df.df.shape[0] > 0:  # якщо є хоча б один розпізнаний запис
                self.data.add_df(cur_df.df)
                result_info += f'OK. Додано записів: {cur_df.df.shape[0]}\n'
            else:
                result_info += 'ЗАПИСИ ВІДСУТНІ\n'

            result_info += f'\n'
            self.progressBar.setValue(pos)
            QApplication.processEvents()

        # Оновлення статусу в вікні GUI
        if self.data.df.shape[0] == 0:
            self.l_cur_file.setText(f'Файлів: {len(user_files[0])}\nСтатус: відсутні валідні дані')
            self.l_cur_file.setStyleSheet("QLabel{color: rgb(150, 0, 0);}")
            self._disable_gui('Відсутні дані в обраних XML файлах')
        else:
            persons_total = len([x for x in self.data.df['g3s'].dropna().unique().tolist() if len(x) > 6])
            self.l_cur_file.setText(f'Файлів: {len(user_files[0])}\n'
                                    f'Статус: записів {self.data.df.shape[0]} (платників: {persons_total})')
            self.l_cur_file.setStyleSheet("QLabel{color: rgb(0, 145, 0);}")
            self.gb_word.setEnabled(True)
            self.gb_excel.setEnabled(True)

        self.statusbar.removeWidget(self.progressBar)
        self.statusbar.showMessage('Опрацювання XML завершено', 5000)

        # Звіт користувачу після завершення опрацювання всіх обраних файлів:
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("Опрацювання обраних XML файлів завершено.")
        msg.setInformativeText("Перевірте (!) успішність опрацювання за кнопкою 'Show details'")
        msg.setWindowTitle("Результати опрацювання XML")
        msg.setDetailedText(result_info)
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

    def _disable_gui(self, message='Помилка завантаження'):
        """Вимкнення кнопок формування звітів та вибірок (у випадку відсутності даних)"""
        self.gb_word.setDisabled(True)
        self.gb_excel.setDisabled(True)
        self.l_cur_file.setText(f'Статус: Не вдалось завантажити XML')
        self.statusbar.showMessage(message, 5000)

    def save_excel(self):
        new_file = QFileDialog.getSaveFileName(self, "Збереження таблиці доходів", '', 'Файл Excel (*.xlsx)')
        if new_file[0] != '':
            self.statusbar.showMessage('Збереження Excel...', 5000)
            QApplication.processEvents()
            self.data: MultiFileDrfoData
            self.data.save_excel(new_file[0],
                                 separate=self.rb_excel_sep.isChecked(),
                                 format_float=self.cb_float_format.isChecked(),
                                 add_profit_column=self.cb_add_profi_col.isChecked())
            self.statusbar.showMessage('Запис Excel файлу завершено', 5000)

    def save_word(self):
        new_file = QFileDialog.getSaveFileName(self, "Збереження звіту", '', 'Файл Word (*.docx)')
        if new_file[0] != '':
            self.statusbar.showMessage('Збереження Word...', 60000)
            QApplication.processEvents()
            word_doc = DocEditor(self.data,
                                 add_years=self.cb_det_years.isChecked(),
                                 add_signs=self.cb_det_types.isChecked(),
                                 add_tab=self.cb_det_tab.isChecked(),
                                 sub_list_text=self.rb_sublist_text.isChecked(),
                                 sub_list_table=self.rb_sublist_table.isChecked())
            available_persons = word_doc.get_available_persons()
            self.progressBar = QProgressBar()
            self.progressBar.setMaximumHeight(18)
            self.progressBar.setMinimumHeight(18)
            self.progressBar.setStyleSheet("""QProgressBar {
                                                    border: 2px solid rgb(211, 211, 211);
                                                    border-radius: 7px;
                                                    background-color: rgb(211, 211, 211);
                                                    text-align: center;
                                                }
                                                QProgressBar::chunk {
                                                    background-color: rgb(246, 191, 39);
                                                    width: 7px; 
                                                    margin: 0.5px;
                                                    border-radius :2px;
                                                }""")
            self.statusBar().addPermanentWidget(self.progressBar)
            self.progressBar.setMaximum(len(available_persons)-1)
            self.progressBar.setMinimum(0)
            self.progressBar.setValue(0)

            for pos, person in enumerate(available_persons):
                word_doc.write_person_to_document(person)
                self.progressBar.setValue(pos)
                QApplication.processEvents()

            word_doc.save_docx(new_file[0])
            self.statusbar.removeWidget(self.progressBar)
            self.statusbar.showMessage('Запис Word файлу завершено', 5000)


def run_gui():
    app = QApplication(sys.argv)
    app.setApplicationName("Skarb - profit converter")
    window = AppWin()
    window.show()
    app.exec_()


if __name__ == '__main__':
    run_gui()
"""
Для заміни у генерованому файлі інтерфейсу:
import gui.res_icons
"""
