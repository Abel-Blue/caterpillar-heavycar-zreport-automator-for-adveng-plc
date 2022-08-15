import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
filename = 'sheets/soodo.xlsx'
sheetname = 'sheets/soodo.xlsx'
# excel can be visible or not
excel.Visible = False  # False     
wb = excel.Workbooks.Open(filename)
sheetname = sheetname
select_table = wb.Sheets(sheetname).Range("A3:N3")

egde_list = [win32c.xlEdgeLeft, win32c.xlEdgeTop, win32c.xlEdgeBottom, win32c.xlEdgeRight, win32c.xlInsideVertical, win32c.xlInsideHorizontal]
# add border line to selected style
for edge in egde_list:
  wb.Sheets(sheetname).Range(select_table, select_table.End(win32c.xlDown)).Borders(edge).LineStyle = win32c.xlContinuous

# remove border line for selected cells
for edge in egde_list:
  wb.Sheets(sheetname).Range("P8:S10").Borders(edge).LineStyle = win32c.xlNone
  
# set thick outer border
# only first 4 in the list, as we only want to set thick border for the edges.
for edge in egde_list[:4]:
  wb.Sheets(sheetname).Range("P3:S11").Borders(edge).Weight = win32c.xlMedium