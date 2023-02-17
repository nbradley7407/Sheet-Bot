# sheet_bot

SheetCP builds upon the gspread module and is used to copy values over a range within a Google Sheets worksheet and paste them into another.

The problem is that gspread does not pick up empty cells at the end of a range, so if the range to be copied shrinks in size over time, old data will still remain. For example, imagine the previous day you had data in range A1:B10 in 'Sheet1'. If the desired range to be pasted over is copied from A1:B10 from 'Sheet2', but only A1:B8 has data, cells A9:B10 in 'Sheet1' will keep the old data and not be replaced by empty cells.

SheetCP addresses this issue by deleting old data after the new data has been copied over.
