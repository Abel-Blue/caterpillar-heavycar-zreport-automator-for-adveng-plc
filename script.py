import win32com.client as win32


excel = win32.gencache.EnsureDispatch('Excel.Application')
filename = 'sheets/soodo.xlsx'
sheetname = 'sheets/soodo.xlsx'

# the file is processed on the next day of the file received
file_received_date = datetime.date.today() - timedelta(days = 1)
# add 1 day for everyday delayed
day_delay = datetime.date.today().weekday()
# the process file dated the next day of the file received
date_process_file = datetime.date.today() - timedelta(days = day_delay)

# Days Encode
# -------------
# Monday: 0
# Tuesday: 1
# Wednesday: 2
# Thursday: 3
# Friday: 4
# Saturday: 5
# Sunday: 6

# Scheduling on Wednesday
date_process_file = datetime.date.today() - timedelta(days = day_delay - 2)
actual_date = datetime.date.today() - timedelta(days = weekday_script_run - weekday_report_scheduled)

# copy paste FORMAT only, which includes font colour, fill colour, and border
wb2.Sheets(sheetname).Range("A:D").Copy()
wb.Sheets(sheetname).Range("A1").PasteSpecial(Paste=win32c.xlPasteFormats)

wb.Sheets(sheetname).Columns("A:A").AutoFit()
wb.Sheets(sheetname).Columns("B:B").ColumnWidth = 19.71
wb.Sheets(sheetname).Columns("C:C").ColumnWidth = 43
wb.Sheets(sheetname).Columns("D:I").AutoFit()
wb.Sheets(sheetname).Columns("J:N").ColumnWidth = 23

wb.Sheets(sheetname).Range("B1").WrapText = True