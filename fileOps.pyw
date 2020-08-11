import os, re, datetime, pprint
import win32com.client as win32

_desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

def getDesktopFilePath(filename):

    return os.path.join(_desktop, filename)

def xlsToXlsx(filepath, folder):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(filepath)
    try:
        fileName = os.path.basename(filepath)+'x'
        wb.SaveAs(os.path.join(folder, fileName), FileFormat = 51)
    except Exception as exc:
        wb.Close()
        raise Exception(exc)
    wb.Close()
    excel.Application.Quit()

    return os.path.join(folder, fileName)

def deleteOldReportsFromFolder(folder, fileTarget=20):
    # deletes the oldest .xlsx files from folder to maintain the target number of files in directory
    files = [file for file in os.listdir(folder) if file[:8].isnumeric() and file.endswith('.xlsx')]
    dateDict = {}
    for file in files:
        day = int(file[:2])
        month = int(file[2:4])
        year = int(file[4:8])
        dateDict.setdefault(datetime.datetime(year = year, month = month, day = day), os.path.join(folder, file))
    dates = list(dateDict.keys())
    dates.sort()
    while len(dates) > fileTarget:
        dateToDelete = dates[0]
        fileToDelete = dateDict[dateToDelete]
        os.unlink(fileToDelete)
        dates.remove(dateToDelete)


