class SheetRange:
    def __init__(self, range):
        self.range = range

        self.end_row = ''
        self.start_row = ''
        self.end_column = ''
        self.start_column = ''

        self._split_range()

    def _split_range(self):
        # format range_p to split the row numbers and column letters as strings
        # "A1:BB500" -> "A", "1", "BB", "500"
        split_range_end = self.range.rsplit(':', 1)[1]
        split_range_start = self.range.rsplit(':', 1)[0]
        for i in split_range_end:
            if i.isdigit():
                self.end_row += i
            else:
                self.end_column += i
        for x in split_range_start:
            if x.isdigit():
                self.start_row += x
            else:
                self.start_column += x