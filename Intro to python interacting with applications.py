import win32com.client as win32
import pythoncom
import sys

# define application events
class ApplicationEvents:

    # define an event inside our application
    def OnSheetActivate(self, *args):
        print('You created a new sheet.')


# define workbook events
class WorkbookEvents:

    def OnSheetSelectionChange(self, *args):

        # print args
        print(args)
        print(args[1].Address)
        args[0].Range('A1').Value = 'You selected cell ' + str(args[1].Address)



# get active instance of excel
excel = win32.GetActiveObject('Excel.Application')

# grab workbook specify the name of the workbook your working in and give the entension i.e. "book1.xlsm"
excel_workbook = excel.Workbooks('Book1')

# assign events to workbook
excel_workbook_events = win32.WithEvents(excel_workbook, WorkbookEvents)

# assign event to excel application object
excel_events = win32.WithEvents(excel, ApplicationEvents)


# define initializer
keepOpen = True

# while there are messages keep displaying them, and also as long as the excel application is still open
while keepOpen:
    
    # display the message
    pythoncom.PumpWaitingMessages()

    try:
        # If there are NO open workbooks stop running
        if excel.Workbooks.Count == 0:
            keepOpen = False
            excel = None
            sys.exit()

    # this basically just says if theres an error just close the application
    except:
        keepOpen = False
        excel = None
        sys.exit()
