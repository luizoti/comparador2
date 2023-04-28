import datetime
import itertools
import os
import sys
from collections import deque
from os.path import abspath, join, basename

import pandas as pd
from openpyxl import load_workbook
from xlsxwriter.utility import xl_rowcol_to_cell

# from xlsxwriter.utility import xl_rowcol_to_cell

# from unittest import TestCase


BASE_DIR = sys.argv[1] if len(sys.argv) >= 2 else abspath(".")


def type_converter(column):
    if column is None:
        return ""

    elif isinstance(column, datetime.datetime):
        return column.strftime('%Y-%m-%d').strip()
    return str(column).strip()


class LeitorDeXlsx:
    """docstring for LeitorDeXlsx"""

    def __init__(self, caminho_do_xlsx):
        super(LeitorDeXlsx, self).__init__()
        self.caminho_do_xlsx = caminho_do_xlsx
        print(f"self.caminho_do_xlsx {self.caminho_do_xlsx}")
        self.wb = load_workbook(self.caminho_do_xlsx)
        self.sheet = self.wb.worksheets[0]
        self.rows = ([[type_converter(x.value) for x in row] for row in self.sheet.rows if
                      set([x.value for x in row]) != {None}])

    def header(self):
        return ["CONCAT"] + list(self.rows)[0]

    def as_dict(self) -> dict:
        rows = (
            ("".join(x), ["".join(x)] + x) for x in itertools.islice(self.rows, 1, None)
        )
        return dict(sorted(rows, key=lambda x: x[0].lower()))

    def dataframe_like(self):
        return itertools.chain([self.header()], self.as_dict().values())


class ComparadorDeXlsx(object):
    """Docstring for ComparadorDeXlsx"""

    def __init__(self, arquivo_orm, arquivo_sql):
        super(ComparadorDeXlsx, self).__init__()
        self.arquivo_comparativo = join(BASE_DIR, f'comparativo_{basename(arquivo_sql.lower().replace(" ", "_"))}')
        self.linhas_arquivo_orm = LeitorDeXlsx(arquivo_orm)
        self.linhas_arquivo_sql = LeitorDeXlsx(arquivo_sql)

        orm_dataframe = pd.DataFrame(self.linhas_arquivo_orm.dataframe_like())
        sql_dataframe = pd.DataFrame(self.linhas_arquivo_sql.dataframe_like())

        # comparativo = pd.DataFrame(dataframe)
        writer = pd.ExcelWriter(self.arquivo_comparativo, engine='xlsxwriter')
        orm_dataframe.style.set_properties(subset=self.linhas_arquivo_orm.header(), width=50)
        sql_dataframe.style.set_properties(subset=self.linhas_arquivo_sql.header(), width=50)

        orm_dataframe.to_excel(writer, sheet_name="ORM")
        print("Folha do ORM processada!")
        sql_dataframe.to_excel(writer, sheet_name="SQL")
        print("Folha do SQL processada!")
        sql_dataframe.to_excel(writer, sheet_name="COMPARATIVO")
        print("Folha do COMPARATIVO processada!")

        writer.close()
        print("Arquivo gravado!!!")
        print(self.arquivo_comparativo)


if __name__ == '__main__':
    # print(LeitorDeXlsx("/home/luiz/v2_SQL_efd_fiscal_c100_c190.xlsx").as_dict())
    orm_file = None
    sql_file = None
    for file in [x for x in os.listdir(BASE_DIR) if x.endswith(".xlsx")]:
        full_file_path = join(BASE_DIR, file)
        if (file.lower().startswith("orm") or "sql" not in file.lower()) and not file.lower().startswith(
                "comparativo_"):
            orm_file = full_file_path
            continue
        elif file.lower().startswith("sql"):
            sql_file = full_file_path
            continue

    ComparadorDeXlsx(orm_file, sql_file)
