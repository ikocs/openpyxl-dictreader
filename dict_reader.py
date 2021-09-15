from __future__ import annotations
from typing import List, Dict, Tuple, Optional, TYPE_CHECKING
if TYPE_CHECKING:
    from openpyxl.worksheet.worksheet import Worksheet
    from openpyxl.cell.cell import Cell


class XlDictReader:
    """Читалка листа эксель в словарь"""
    def __init__(self,
                 sheet: Worksheet,
                 header_row: int = 1,
                 fieldnames: Optional[list[str]] = None,
                 restval: Optional[str] = None,
                 restkey: Optional[str] = None):
        self.sheet = sheet
        self.rows = self.sheet.iter_rows()
        self.header_row = header_row
        self.restval = restval
        self.restkey = restkey
        self.line_num = self.header_row - 1
        self.header: List[str] = fieldnames or self.read_header()

    def __repr__(self):
        head_strings = map(str, self.header)
        return f'Head: {" | ".join(head_strings)}'

    def read_header(self):
        """Считыватель заголовка таблицы"""
        for _ in range(self.header_row - 1):
            next(self.rows)

        header_cells: Tuple[Cell] = next(self.rows)
        header = [str(cell.value) for cell in header_cells]
        self.line_num += 1
        return header

    def __iter__(self):
        return self

    def __next__(self) -> Dict[str, Optional[Cell]]:
        row = next(self.rows)
        self.line_num += 1
        d = dict(zip(self.header, row))

        # Обработка случаев, когда количество элементов
        # в fieldnames не равно размеру строки
        lf = len(self.header)
        lr = len(row)
        if lf < lr:
            d[self.restkey] = row[lf:]
        elif lf > lr:
            for key in self.header[lr:]:
                d[key] = self.restval
        return d
