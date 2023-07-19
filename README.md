# sheet_bot

SheetCP builds upon the gspread module and is used to copy values over a range within a Google Sheets worksheet and paste them into another.

The problem is that gspread does not pick up empty cells at the end of a range, so if the range to be copied shrinks in size over time, old data will still remain. SheetCP addresses this issue by deleting old data after the new data has been copied over.
