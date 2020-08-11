from windowOps import *
from dateOps import *
from fileOps import *
import time, os
from errorLogs import errorLog

@errorLog
def dailyReport(date=getNetworkDay(-1)):

    try:
        app = getToolingApp()
    except:
        raise Exception('Nie znaleziono okna aplikacji PWSK')

    closeOtherWindowsByTitle(app, 'EVIMASTER')
    
    window = app.top_window()
    runReportByMenuItemIndex(window, 11)

    app.window(title='Rejestr dokumentów').wait('active')

    reportWindow = app.window(title='Rejestr dokumentów')

    runDocumentReport(reportWindow, date)

    contextWindow = app.top_window()
    contextWindow['MenuItem3'].click_input(button='left')
    
    while len(app.windows()) < 3:
        time.sleep(0.5)

    xlsFilename = f"{date}.xls"
    saveWindow = app.top_window()
    saveToDesktop(saveWindow, xlsFilename)

    excelApp = None
    timer = 0
    while excelApp == None:
        try:
            excelApp = getExcelApp()
        except:
            if timer == 10:
                getExcelApp(greedy=True)
                break
            time.sleep(1)
            timer +=1

    refuseExcelOpening(excelApp)

    xlsPath = getDesktopFilePath(xlsFilename)
    xlsxPath = xlsToXlsx(xlsPath, _dataFolder)

    if os.path.exists(xlsxPath):
        os.unlink(xlsPath)

    deleteOldReportsFromFolder(_dataFolder)

@errorLog
def monthlyReport(dateRange = getExpiryRange()):

    try:
        app = getToolingApp()
    except:
        raise Exception('Nie znaleziono okna aplikacji PWSK')

    closeOtherWindowsByTitle(app, 'EVIMASTER')
    
    window = app.top_window()
    
    runReportByMenuItemIndex(window, 3)

    app.window(title_re=r'Kontrola.*').wait('active')
    
    reportWindow = app.top_window()

    runExpirationReport(reportWindow, dateRange)
    contextWindow = app.top_window()
    contextWindow['MenuItem3'].click_input(button='left')

    while len(app.windows()) < 3:
        time.sleep(0.5)

    xlsFilename = f"{dateRange[1]}-terminy.xls"
    
    saveWindow = app.top_window()
    saveToDesktop(saveWindow, xlsFilename)

    excelApp = None
    timer = 0
    while excelApp == None:
        try:
            excelApp = getExcelApp()
        except:
            if timer == 5:
                break
            time.sleep(1)
            timer +=1

    try:
        refuseExcelOpening(excelApp)
    except Exception as exc:
        print(exc)

    xlsPath = getDesktopFilePath(xlsFilename)
    xlsxPath = xlsToXlsx(xlsPath, _monthlyDataFolder)

    print(xlsxPath)

    if os.path.exists(xlsxPath):
        os.unlink(xlsPath)

    deleteOldReportsFromFolder(_monthlyDataFolder, fileTarget=1)
    
    
if __name__ == "__main__":
    pass
