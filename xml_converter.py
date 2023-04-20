"""
Модуль опрацювання XML файлу:
    - читання файлу
    - опрацювання структури
    - формування датафрейму
    - підготовка датафрейму до експорту
    - окремий клас накопичення даних декількох файлів
"""

import re
from pathlib import Path
from typing import Union

import pandas as pd
import numpy as np
import xml.etree.ElementTree as ET

from defines import dict_long as sign_dict_default, response, service_col_names


class CellProfit:
    def __init__(self, cell_adr: str, row_num: int, value: Union[int, str, float]):
        self.cell = cell_adr
        self.row = row_num
        self.col = 'G2S'
        self.value = value
        self.status = False
        self.valid()

    def valid(self):
        if len(self.cell.lower().split('xxxx')) == 2:
            self.col = self.cell.lower().split('xxxx', -1)[1].strip()
            self.status = True
        else:
            return


class FileProfitXML:
    headers = {'g2s': 'Особа №',
               'g3s': 'РНОКПП',
               'g4s': 'Результат обробки',
               'g5': 'Тип ФО',
               'g6s': 'Код агента',
               'g7s': 'Назва агента',
               'g8': 'Дохід',
               'g9': 'Податок',
               'profit': 'Прибуток',
               'g10': 'Ознака доходу',
               'g11': 'Квартал',
               'g12': 'Рік'}
    col_int = ['g5', 'g10', 'g11', 'g12']
    col_float = ['g8', 'g9']
    signs = sign_dict_default

    def __init__(self, file: Union[str, Path]):
        assert type(file) in [str, Path], "Тип посилання на файл - string або екземпляр Path"
        if type(file) == str:
            file = Path(file)
        self.file = file
        self.max_rows = 0
        self.columns = set()
        self.df = pd.DataFrame()
        self.cells_collection = []

    def read_xml(self) -> int:
        """
        Читання файлу XML, перевірка відповідності схеми

        :return: error code: 0 - OK, 1 - ERROR
        """
        try:
            tree = ET.parse(self.file)
        except Exception:
            return 1

        body = tree.find('DECLARBODY')
        for elem in body:
            adr = str(elem.tag)
            if adr.startswith("T1R"):
                row_num = int(elem.attrib.get('ROWNUM', 0))
                if row_num != 0:
                    self.max_rows = row_num if row_num > self.max_rows else self.max_rows
                    cur_cell_inst = CellProfit(cell_adr=elem.tag,
                                               row_num=row_num,
                                               value=elem.text)
                    if cur_cell_inst.status:
                        self.cells_collection.append(cur_cell_inst)
        for cell_inst in self.cells_collection:
            cell_inst: CellProfit
            self.columns.add(cell_inst.col)
        return 0

    def check_columns_set(self, df_new=None) -> bool:
        """
        Перевірка чи наявний достатній набір колонок у імпортованому файлі

        :return: True - Ok, False - недостатньо колонок для опрацювання
        """
        if df_new is None:
            if len(set(service_col_names.keys()).intersection(self.columns)) == 11:
                return True
            else:
                return False
        else:
            if len(set(service_col_names.keys()).intersection(df_new.columns)) == 11:
                return True
            else:
                return False

    def fill_df(self) -> str:
        """
        Створення порожнього датафрейму відповідно отриманої розмірності (рядки/колонки) та заповнення
        його записами файлу XML

        :return: текстовий опис виявлених помилок
        """
        warnings = ''

        # Перевірка достатності даних для побудови датафрейму:
        if not self.check_columns_set():
            absent_columns = ', '.join([str(x).upper() for x in list(set(service_col_names.keys()) - self.columns)])
            warnings += f'Неправильний формат. У файлі відсутні необхідні колонки: {absent_columns}\n'
            return warnings
        if self.max_rows == 0:
            warnings += f'Неправильний формат. У файлі відсутні записи.\n'
            return warnings

        # Створення датафрейму з розмірами, що відповідають кількості записів/колонок:
        self.df = pd.DataFrame(np.nan, np.arange(self.max_rows), columns=list(self.columns))

        # Внесення кожного запису (клітинки) до датафрейму:
        for c in self.cells_collection:
            c: CellProfit
            self.df.at[c.row - 1, c.col] = c.value  # заповнення "запис XML - клітинка таблиці"

        # Видалення рядку "Декларація фізичної особи" - не приймає участі у аналізі
        rows_before = self.df.shape[0]
        self.df.drop(self.df[self.df['g10'].isin([888, '888'])].index, inplace=True)
        if rows_before != self.df.shape[0]:
            pass
            # warnings += "У таблиці наявні записи щодо декларування фізичної особи (код 888), записи про подані " \
            #             "декларації не враховуються при подальшій роботі Skarb (не вносяться в аналітичні звіти " \
            #             "та не додаються у експорт Excel). Зазначене не стосується доходів підприємницької " \
            #             "діяльності.\n"

        # Вирішення помилкового коду вивантаження з БД:
        self.df.fillna(np.nan, inplace=True)  # Перетворення None до np.nan
        missing_persons = self.df['g3s'].isna().sum()
        if missing_persons:
            warnings += f'Видалено {missing_persons} записів у яких відсутні значення РНОКПП\n'
            self.df.dropna(subset=['g3s'], inplace=True)
        self.df['g4s'].fillna(10)
        self.df['g4s'] = self.df['g4s'].astype(int)

        if len(set(self.df['g4s'].unique()).intersection(set(response.keys()))) > 0:
            warnings += 'Наявні записи, що свідчать про негативну відповідь на запит:\n'
            for err_code in response.keys():
                failed_persons = self.df.loc[self.df['g4s'] == err_code]['g3s'].unique().tolist()
                for p in failed_persons:
                    to_del = ', '.join([str(x+1) for x in list(self.df.loc[(self.df['g4s'] == err_code) &
                                                                           (self.df['g3s'] == p)].index)])
                    warnings += f"- РНОКПП {p}: {response.get(err_code, 'помилковий код відповіді')} " \
                                f"(видалено рядки № {to_del})\n"
                    self.df.drop(self.df[(self.df['g4s'] == err_code) & (self.df['g3s'] == p)].index, inplace=True)

        # Виправлення дублювання коштів у звітах (6-місяців, 9-місяців, річних) для декларацій єдиного податку:
        self.df = self._tax_declaration_fix(self.df)

        # Вирішення місінгів, які можна відновити:
        self.df = self.df.apply(lambda row: self.fill_na_tax_codes(row), axis=1)  # заповнення агенту для ФОП
        self.df['g11'].fillna(4, inplace=True)  # заповнення кварталу у разі порожнього значення

        na_income = self.df['g8'].isna().sum()  # місінги у значенні доходів
        na_tax = self.df['g9'].isna().sum()  # місінги у значенні податків
        na_name_employer = self.df['g7s'].isna().sum()  # місінги у найменуванні роботодавця
        na_income_type = self.df['g10'].isna().sum()  # місінги у видах доходу

        if na_income > 0:
            ind_na_income = ', '.join([str(x+1) for x in list(self.df.loc[pd.isna(self.df["g8"]), :].index)])
            warnings += f'Відсутні суми доходу у {na_income} рядках, замінені на 0.00 (№: {ind_na_income})\n'
            self.df['g8'].fillna(0.0, inplace=True)

        if na_tax > 0:
            ind_na_tax = ', '.join([str(x+1) for x in list(self.df.loc[pd.isna(self.df["g9"]), :].index)])
            warnings += f'Відсутні суми податку у {na_tax} рядках, замінені на 0.00 (№: {ind_na_tax})\n'
            self.df['g9'].fillna(0.0, inplace=True)

        if na_name_employer > 0:
            ind_na_empname = ', '.join([str(x+1) for x in list(self.df.loc[pd.isna(self.df["g7s"]), :].index)])
            warnings += f'Відсутні назви джерела у {na_name_employer} рядках, замінені на "Не відомо" (№: {ind_na_empname})\n'
            self.df['g7s'].fillna("Не зазначено", inplace=True)

        if na_income_type > 0:
            ind_na_type = ', '.join([str(x+1) for x in list(self.df.loc[pd.isna(self.df["g7s"]), :].index)])
            warnings += f'Відсутні види доходу у {na_income_type} рядках, замінені на "код 14 Інші доходи" (№: {ind_na_type})\n'
            self.df['g7s'].fillna(14, inplace=True)

        # # Вирішення місінгів в обовязкових колонках:
        # req_columns = {'g7s': 'Назва агента', "g10": "Вид доходу", "g12": "Рік"}
        # for column in req_columns.keys():
        #     missing_count = self.df[column].isna().sum()
        #     if missing_count:
        #         warnings += f'Видалено {missing_count} записів у яких відсутні значення поля ' \
        #                     f'"{service_col_names.get(column, column)}"\n'
        #         self.df.dropna(subset=[column], inplace=True)

        # Приведення числових типів у відповідність:
        for col in self.col_int:
            if col in self.df.columns:
                self.df[col] = self.df[col].astype(int, errors="ignore")
        for col in self.col_float:
            if col in self.df.columns:
                self.df[col] = self.df[col].astype(float, errors="ignore")

        # Перевірка, чи залишились записи після видалення місінгів:
        if self.df.shape[0] == 0:
            warnings += 'Після очищення помилкових значень не залишилось валідних записів.\n'
            return warnings

        # Розрахунок колонки прибутку:
        self.df['profit'] = self.df['g8'] - self.df['g9']
        return warnings

    def _get_formatted_df(self, external_df=None, format_float=True, add_profit=True) -> pd.DataFrame:
        if not type(external_df) == pd.DataFrame:
            df = self.df
        else:
            df = external_df

        if add_profit:
            df_view = df[['g2s', 'g3s', 'g6s', 'g7s', 'g8', 'g9', 'profit', 'g10', 'g11', 'g12']].copy()
        else:
            df_view = df[['g2s', 'g3s', 'g6s', 'g7s', 'g8', 'g9', 'g10', 'g11', 'g12']].copy()

        def f2s(amount):
            try:
                thou_sep = ' '
                deci_sep = '.'
                w_dec = '%.2f' % amount
                part_int = w_dec.split('.')[0]
                part_int = re.sub(r"\B(?=(?:\d{3})+$)", thou_sep, part_int)
                part_dec = w_dec.split('.')[1]
                return part_int + deci_sep + part_dec
            except Exception:
                print(f'Error with float value {amount} (type {type(amount)}) - cant convert to string')
                return '0.00'

        if format_float:
            df_view['g8'] = df_view['g8'].apply(lambda x: f2s(x))
            df_view['g9'] = df_view['g9'].apply(lambda x: f2s(x))
            if 'profit' in df_view.columns:
                df_view['profit'] = df_view['profit'].apply(lambda x: f2s(x))
        df_view.replace({'g10': self.signs}, inplace=True)
        df_view.rename(columns=self.headers, inplace=True)
        df_view.fillna('Не зазначено', inplace=True)
        return df_view

    def save_excel(self, file: Union[str, Path], separate=False, format_float=True, add_profit_column=True):
        """
        Збереження форматованого файлу таблиці Excel

        :param file: назва створюваного файлу
        :param separate: розділення на декілька файлів в разі записів щодо декількох осіб
        :param format_float: форматування сум (12300,00 -> 12 300.00)
        :param add_profit_column: додати колонку розрахунку прибутку (дохід - податок)
        """
        if type(file) == str:
            file = Path(file)

        if not separate:
            df = self._get_formatted_df(format_float=format_float, add_profit=add_profit_column)
            df.to_excel(file, index=False)
        else:
            persons = self.df['g3s'].dropna().unique().tolist()
            for p in persons:
                df = self.df.loc[self.df['g3s'] == p]
                cur_path = file.with_name(f"{file.stem}_{str(p)}{file.suffix}")
                df_f = self._get_formatted_df(df, format_float=format_float, add_profit=add_profit_column)
                df_f.to_excel(cur_path, index=False)

    @staticmethod
    def _tax_declaration_fix(df: pd.DataFrame):
        """
        Нормалізація доходів зазначених в деклараціях платника єдиного податку:
        виключення піврічних звітів, які включаються 9-річними, формування окремого виду доходу щодо
        доходу отриманого від підприємницької діяльності (коди 506, 509, 512).
        Дублювання звітів визначається для кожного року та кожної окремої особи.
        Передбачається використання методу під час опрацювання місінгів (тобто до приведення
        типів до цілих чисел).
        """
        df = df.copy()
        # Визначення років, у які подавались декларації:
        years_with_declar = df.loc[df['g10'].isin(['503', '506', '509', '512'])]['g12'].unique()
        # Перевірка чи в кожному році прийшли річні звіти:
        for year in years_with_declar:
            pers_with_declar = df.loc[(df['g10'].isin(['506', '509', '512'])) & (df['g12'] == year)]['g3s'].unique()
            for person in pers_with_declar:
                tax_signs_present = df.loc[(df['g3s'] == person) &
                                           (df['g12'] == year) &
                                           (df['g10'].isin(['506', '509', '512'])), 'g10'].unique()
                # Якщо є річний звіт - видалити проміжні
                if '512' in tax_signs_present:
                    df.drop(df[(df['g3s'] == person) &
                               (df['g12'] == year) &
                               (df['g10'].isin(['506', '509', '503']))].index, inplace=True)
                # Якщо немає річного, але є за 9 місяців - видалити попередні звіти:
                elif '509' in tax_signs_present:
                    df.drop(df[(df['g3s'] == person) &
                               (df['g12'] == year) &
                               (df['g10'].isin(['503', '506']))].index, inplace=True)
                # Якщо немає річного та 9 місячного звітів, але є за 6 місяців - видалити  звіт першого кварталу:
                elif '506' in tax_signs_present:
                    df.drop(df[(df['g3s'] == person) &
                               (df['g12'] == year) &
                               (df['g10'].isin(['503']))].index, inplace=True)

        # Привести ознаки залишених звітів до загального:
        df.reset_index(inplace=True, drop=True)
        df.replace({'g10': {'503': '512', '509': '512', '506': '512'}}, inplace=True)
        return df

    @staticmethod
    def fill_na_tax_codes(row):
        """
        Заповнення значення роботодавця в разі коли запис стосується ФОП. Код 512 - річний звіт ФОП.
        Передбачається, що метод викликається після видалення записів про 6 та 9-місячні звіти.
        """
        if row['g10'] in [512, '512', 503, "503", 506, "506", 509, "509"]:
            row['g6s'] = row['g3s']
            row['g7s'] = 'ДОХОДИ ВЛАСНОЇ ПІДПРИЄМНИЦЬКОЇ ДІЯЛЬНОСТІ'
        return row


class MultiFileDrfoData(FileProfitXML):
    def __init__(self):
        super().__init__(file='dummy path')  # dummy path

    def add_df(self, df_new: pd.DataFrame):
        self.df = pd.concat([self.df, df_new], ignore_index=True, sort=False)
        self.df.reset_index(inplace=True, drop=True)

    def read_xml(self) -> int:
        """
        Читання файлу XML, перевірка відповідності схеми

        :return: error code: 0 - OK, 1 - ERROR
        """
        raise AttributeError('Multi data instance not allowed the method. Use parent class')

    def fill_df(self) -> str:
        """
        Створення порожнього датафрейму відповідно отриманої розмірності (рядки/колонки) та заповнення
        його записами файлу XML

        :return: текстовий опис виявлених помилок
        """
        raise AttributeError('Multi data instance not allowed the method. Use parent class')

    @staticmethod
    def _tax_declaration_fix(df: pd.DataFrame):
        """
        Нормалізація доходів зазначених в деклараціях платника єдиного податку:
        виключення піврічних звітів, які включаються 9-річними, формування окремого виду доходу щодо
        доходу отриманого від підприємницької діяльності (коди 506, 509, 512).
        Дублювання звітів визначається для кожного року та кожної окремої особи.
        Передбачається використання методу під час опрацювання місінгів (тобто до приведення
        типів до цілих чисел).
        """
        raise AttributeError('Multi data instance not allowed the method. Use parent class')

    @staticmethod
    def fill_na_tax_codes(row):
        """
        Заповнення значення роботодавця в разі коли запис стосується ФОП. Код 512 - річний звіт ФОП.
        Передбачається, що метод викликається після видалення записів про 6 та 9-місячні звіти.
        """
        raise AttributeError('Multi data instance not allowed the method. Use parent class')

