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
from pathlib import Path

from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox

from gui.main_gui import Ui_MainWindow
from xml_converter import FileProfitXML
from word_reporter import DocEditor


class AppWin(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.data = None

        self.b_import.clicked.connect(self.import_file)
        self.b_word.clicked.connect(self.save_word)
        self.b_excel.clicked.connect(self.save_excel)
        self.l_cur_file.setText(f'Статус: Готовий до роботи')

    def import_file(self):
        chosen_file = QFileDialog.getOpenFileName(self, 'Вибір файлу відомостей про доходи', str(Path.cwd().absolute()),
                                                  'Файли XML (*.xml *.XML)')[0]
        if chosen_file:
            self.data = FileProfitXML(chosen_file)
            success = (not bool(self.data.read_xml()))  # спроба прочитати XML файл
            if not success:
                self._disable_gui('Помилка читання XML файлу')
                self.l_cur_file.setText(f'Файл: {Path(chosen_file).name}\nСтатус: Не вдалось прочитати XML')
                self.l_cur_file.setStyleSheet("QLabel{color: rgb(150, 0, 0);}")
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Warning)
                msg.setText("Помилка читання XML файлу.")
                msg.setInformativeText("Можливо файл відкритий іншою програмою.")
                msg.setWindowTitle("Помилка XML")
                msg.setStandardButtons(QMessageBox.Ok)
                msg.exec_()
                return

            warnings = self.data.fill_df()
            print(self.data.df.dtypes)
            if warnings != '':
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Information)
                msg.setText("Під час імпорту виявлені невалідні записи.")
                msg.setInformativeText("Перевірте критичність помилок за кнопкою 'Show details'")
                msg.setWindowTitle("Попередження")
                msg.setDetailedText(warnings)
                msg.setStandardButtons(QMessageBox.Ok)
                msg.exec_()

            if self.data.df.shape[0] == 0:
                self._disable_gui('Формат XML неочікуваний (0 записів)')
                self.l_cur_file.setText(f'Файл: {Path(chosen_file).name}\nСтатус: Не вдалось прочитати XML')
                self.l_cur_file.setStyleSheet("QLabel{color: rgb(150, 0, 0);}")
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Information)
                msg.setText("Не виявлено валідних записів")
                msg.setInformativeText("Особи не мали офіційних джерел доходів, відмовлено у видачі з БД або файл "
                                       "не відноситься до формату реєстру ДРФО/змінювався сторонніми програмами. "
                                       "Деталі імпорту за кнопкою 'Show details'")
                msg.setWindowTitle("Неочікуваний формат")
                msg.setDetailedText(warnings)
                msg.setStandardButtons(QMessageBox.Ok)
                msg.exec_()
                return
            else:
                # self.gb_import.setDisabled(True)
                self.gb_word.setEnabled(True)
                self.gb_excel.setEnabled(True)
                self.statusbar.showMessage(f'XML опрацьовано', 5000)
                # self.l_cur_file.setText(f'Статус: записів {self.data.df.shape[0]} ({Path(chosen_file).name})')
                persons_count = self.data.df['g3s'].nunique()
                self.l_cur_file.setText(f'Файл: {Path(chosen_file).name}\n'
                                        f'Статус: записів {self.data.df.shape[0]} (платників: {persons_count})')
                self.l_cur_file.setStyleSheet("QLabel{color: rgb(0, 145, 0);}")

    def _disable_gui(self, message='Помилка завантаження'):
        self.gb_word.setDisabled(True)
        self.gb_excel.setDisabled(True)
        self.l_cur_file.setText(f'Статус: Не вдалось завантажити XML')
        self.statusbar.showMessage(message, 5000)

    def save_excel(self):
        new_file = QFileDialog.getSaveFileName(self, "Збереження таблиці доходів", '', 'Файл Excel (*.xlsx)')
        if new_file[0] != '':
            self.data: FileProfitXML
            self.data.save_excel(new_file[0],
                                 separate=self.rb_excel_sep.isChecked(),
                                 format_float=self.cb_float_format.isChecked(),
                                 add_profit_column=self.cb_add_profi_col.isChecked())
            self.statusbar.showMessage('Збереження EXCEL завершено', 5000)

    def save_word(self):
        new_file = QFileDialog.getSaveFileName(self, "Збереження звіту", '', 'Файл Word (*.docx)')
        if new_file[0] != '':
            word_doc = DocEditor(self.data,
                                 add_years=self.cb_det_years.isChecked(),
                                 add_signs=self.cb_det_types.isChecked(),
                                 add_plots=self.cb_det_plot.isChecked(),
                                 add_tab=self.cb_det_tab.isChecked(),
                                 sub_list_text=self.rb_sublist_text.isChecked(),
                                 sub_list_table=self.rb_sublist_table.isChecked())
            word_doc.save_docx(new_file[0])
            self.statusbar.showMessage('Збереження WORD завершено', 5000)


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
