import openpyxl

from . import patch, xltypes


class Reader():
    _book: openpyxl.Workbook

    def __init__(self, file_name: str):
        self.excel_file_name = file_name
        self._book = None

    @property
    def book(self):
        if not self._book:
            self._book = self.read()

        return self._book


    def read(self):
        with patch.openpyxl_WorksheetReader_patch():
            return openpyxl.load_workbook(self.excel_file_name)

    def read_defined_names(self, ignore_sheets=[], ignore_hidden=False) -> dict[str, str]:
        return {
            defn.name: defn.value
            for name, defn in self.book.defined_names.items()
            if not defn.hidden and defn.value != '#REF!'
        }

    def read_cells(self, ignore_sheets=[], ignore_hidden=False):
        cells = {}
        formulae = {}
        ranges = {}
        for sheet_name in self.book.sheetnames:
            if sheet_name in ignore_sheets:
                continue
            sheet = self.book[sheet_name]
            for row in sheet.rows:
                for cell in row:
                    addr = f'{sheet_name}!{cell.coordinate}'
                    if cell.data_type == 'f':
                        value = cell.value
                        if isinstance(
                                value,
                                openpyxl.worksheet.formula.ArrayFormula
                        ):
                            value = value.text
                        formula = xltypes.XLFormula(value, sheet_name)
                        formulae[addr] = formula
                        value = cell.cvalue
                    else:
                        formula = None
                        value = cell.value

                    cells[addr] = xltypes.XLCell(
                        addr, value=value, formula=formula)

        return [cells, formulae, ranges]
