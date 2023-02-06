"""
Модуль опрацювання XML файлу:
    - читання файлу
    - опрацювання структури
    - формування датафрейму
    - підготовка датафрейму до експорту
"""

import re
from pathlib import Path
from typing import Union

import pandas as pd
import numpy as np
import xml.etree.ElementTree as ET

from defines import dict_long as sign_dict_default


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

    def check_col_set(self) -> int:
        """
        Перевірка чи наявний достатній набір колонок у імпортованому файлі

        :return: error code: 0 - OK, 1 - ERROR
        """
        if len(set(self.headers.keys()).intersection(set(self.columns))) == 11:
            return 0
        else:
            return 1

    def fill_df(self) -> str:
        """
        Створення порожнього датафрейму відповідно отриманої розмірності (рядки/колонки) та заповнення
        його записами файлу XML

        :return: текстовий опис виявлених помилок
        """
        warnings = ''
        # Створення датафрейму з розмірами, що відповідають кількості записів/колонок:
        self.df = pd.DataFrame(np.nan, np.arange(self.max_rows), columns=list(self.columns))

        # Внесення кожного запису до датафрейму:
        for c in self.cells_collection:
            c: CellProfit
            self.df.at[c.row - 1, c.col] = c.value  # заповнення "запис XML - клітинка таблиці"
        # Приведення типів у відповідність:
        for col in self.col_int:
            if col in self.df.columns:
                self.df[col] = self.df[col].astype(int, errors="ignore")
        for col in self.col_float:
            if col in self.df.columns:
                self.df[col] = self.df[col].astype(float, errors="ignore")
        if self.df.shape[0] == 0:
            return ''
        # Видалення рядку "Загалом"
        self.df.drop(self.df[self.df['g10'] == 888].index, inplace=True)

        # Перевірка грошових сум, розрахунок прибутку:
        self.df.fillna(np.nan, inplace=True)  # Перетворення None до np.nan
        na_income = self.df['g8'].isna().sum()
        na_tax = self.df['g9'].isna().sum()

        if na_income > 0:
            ind_na_income = ', '.join([str(x+1) for x in list(self.df.loc[pd.isna(self.df["g8"]), :].index)])
            warnings += f'Відсутні суми доходу у {na_income} рядках (№: {ind_na_income})\n'
            print(f'NA values in: INCOME={na_income} (index: {ind_na_income})')
            self.df['g8'].fillna(0.0, inplace=True)

        if na_tax > 0:
            ind_na_tax = ', '.join([str(x+1) for x in list(self.df.loc[pd.isna(self.df["g9"]), :].index)])
            warnings += f'Відсутні суми податку у {na_tax} рядках (№: {ind_na_tax})\n'
            print(f'NA values in: TAX={na_tax} (index: {ind_na_tax})')
            self.df['g9'].fillna(0.0, inplace=True)

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
