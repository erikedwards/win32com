import win32com.client as win32
import pythoncom

# define Application Events
class ApplicationEvents:
    # define an event inside our Application
    def OnSheetActivate(self, *args):
        print("You activated a new sheet.")

# define Workbook Events
class WorkbookEvents:
    # define an event inside out workbook
    def OnSheetSelectionChange(self, *args):
        print(args)
        print(args[1].Address)
        args[0].Range('A1').Value = 'You selected cell ' + str(args[1].Address)

# Get the active instance of Excel
xl = win32.GetActiveObject('Excel.Application')

# assign our event to the Excel Application Object
xl_events = win32.WithEvents(xl, ApplicationEvents)

# grab the work book
xl_workbook = xl.Workbooks('Book1')

# assign events to Workbook
xl_workbook_events = win32.WithEvents(xl_workbook, WorkbookEvents)

# while there are messages keep displaying them
while True:
    # display the message
    pythoncom.PumpWaitingMessages()