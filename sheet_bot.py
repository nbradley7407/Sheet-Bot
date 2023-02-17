import gspread
from range import SheetRange
from openpyxl.utils.cell import column_index_from_string
from openpyxl.utils.cell import get_column_letter


class SheetCP:
    # key refers to the key between /d/ and /edit in the url
    def __init__(self, key_c, key_p, sheet_c, sheet_p, range_c, range_p, credentials):
        self.credentials = credentials
        gc = gspread.service_account(filename=self.credentials)

        self.key_c = gc.open_by_key(key_c)
        self.key_p = gc.open_by_key(key_p)

        self.sheet_c = self.key_c.worksheet(sheet_c)
        self.sheet_p = self.key_p.worksheet(sheet_p)

        self.range_c = range_c
        self.range_p = range_p

        self.range_values = self.sheet_c.get_values(self.range_c)

    def _set_values(self):
        # grab values to copy and paste, and format null values to empty strings
        self.range_values = self.sheet_c.get_values(self.range_c)
        return self.range_values

    def _paste_values(self):
        # paste values in destination sheet
        self.sheet_p.update(self.range_p, self.range_values, raw=False)

    def _clear_bottom(self):
        # format range_p to split the row numbers and column letters as strings
        # "A1:BB500" -> "A", "1", "BB", "500"
        sheet_range = SheetRange(self.range_p)

        # find first empty row (beginning paste row + length of copy range that have values) & concatenate the strings
        begin_clear_range = sheet_range.start_column + str(int(sheet_range.start_row) + len(self.range_values))
        end_clear_range = sheet_range.end_column + str(int(sheet_range.end_row))
        clear_range = begin_clear_range + ':' + end_clear_range

        # if any remaining rows in range, clear them
        if (int(sheet_range.start_row) + len(self.range_values)-1) < int(sheet_range.end_row):
            self.sheet_p.batch_clear([clear_range])

    def _clear_side(self):
        # format range_p to split the row numbers and column letters as strings
        sheet_range = SheetRange(self.range_p)

        # find first empty row (beginning paste row + length of copy range that have values) & concatenate the strings
        start_column_calc = get_column_letter(column_index_from_string(sheet_range.start_column)
                                              + len(self.range_values[0]))
        begin_clear_range = start_column_calc + sheet_range.start_row
        end_clear_range = sheet_range.end_column + sheet_range.end_row
        clear_cols = begin_clear_range + ':' + end_clear_range
        # if any remaining rows in range, clear them
        if (column_index_from_string(sheet_range.start_column) + len(self.range_values[0]) - 1)\
                < column_index_from_string(sheet_range.end_column):
            self.sheet_p.batch_clear([clear_cols])

    def run(self):
        self._set_values()
        self._paste_values()
        self._clear_bottom()
        self._clear_side()

