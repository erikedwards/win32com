import win32com.client

# create and instance of Excel & make visible
ExcelApp = win32com.client.Dispatch("Excel.Application")
ExcelApp.Visible = True

# create a new workbook in the Excel app
ExcelWorkbook = ExcelApp.Workbooks.Add()

# create a new sheet in the workbook
ExcelWrkSht = ExcelWorkbook.Worksheets.Add()

# set a reference to a range of cells
ExcelRng1 = ExcelWrkSht.Range("A1:A11")
ExcelRng1.Value = 69

# set a reference to a range of cells
Cell1 = ExcelWrkSht.Cells(1,4)
Cell2 = ExcelWrkSht.Cells(7,7)
ExcelRng2 = ExcelWrkSht.Range(Cell1, Cell2)
ExcelRng2.Value = 420

# count the cells in a range
print(ExcelRng2.Cells.Count)

# Loop through the range of cells
for cellItem in ExcelRng2:
    print(cellItem.Address)

for i in range(len(ExcelRng2)):
    print(cellItem.Value)