import PySimpleGUI as sg
import toolshop as ts

sg.change_look_and_feel('Reddit')
reportList = ['Dzienny', 'Miesięczny']
exitFlag = False

while not exitFlag:
    reportSelectLayout = [
                [sg.Txt('Wybierz typ raportu', size=(30,2), auto_size_text=False)],
                [sg.DropDown(reportList, default_value=reportList[0], size=(30,6), readonly=True)],
                [sg.OK()]
                ]

    window = sg.Window('Raporty narzędziownia').Layout(reportSelectLayout)
    firstEvent, values = window.Read()
    report = values[0]
    window.close()
    if firstEvent in [None, '', 'Exit']:
        exitFlag = True
        break
    else:
        break

if report == reportList[0] and not exitFlag:
    while not exitFlag:
        defaultDate = ts.getNetworkDay(-1)
        dailyConfirmLayout = [
                [sg.Txt('Wykonaj raport dzienny dla daty:', size=(30,2), auto_size_text=False)],
                [sg.InputText((f'{defaultDate[:2]}'),size=(2,2)),sg.InputText((f'{defaultDate[2:4]}'),size=(2,2)),sg.InputText((f'{defaultDate[4:]}'),size=(4,2)),],
                [sg.OK()],
                ]
        window = sg.Window('Raport dzienny').Layout(dailyConfirmLayout)
        event, values = window.Read()
        try:
            userDate = ''.join([values[0], values[1], values[2]])
        except:
            pass
        window.close()
        if event in [None, '', 'Exit']:
            exitFlag = True
            break
        elif not ts.isValidDate(userDate):
            sg.Popup('Błędna data', title='Error')
            continue
        else:
            try:
                ts.dailyReport(userDate)
            except Exception as exc:
                sg.Popup(f'{exc}', title='Error')
                exitFlag = True
            break

if report == reportList[1] and not exitFlag:
    while not exitFlag:
        defaultStartDate, defaultEndDate = ts.getExpiryRange()
        monthlyConfirmLayout = [
                [sg.Txt('Wykonaj raport masowy dla zakresu:', size=(30,2), auto_size_text=False)],
                [sg.InputText((f'{defaultStartDate[:2]}'),size=(2,2)),sg.InputText((f'{defaultStartDate[2:4]}'),size=(2,2)),sg.InputText((f'{defaultStartDate[4:]}'),size=(4,2)),
                 sg.Txt('do:', size=(3,1), auto_size_text=False),
                 sg.InputText((f'{defaultEndDate[:2]}'),size=(2,2)),sg.InputText((f'{defaultEndDate[2:4]}'),size=(2,2)),sg.InputText((f'{defaultEndDate[4:]}'),size=(4,2))],
                [sg.OK()]
                ]
        window = sg.Window('Raport miesięczny').Layout(monthlyConfirmLayout)
        event, values = window.Read()
        try:
            userStart = ''.join([values[0], values[1], values[2]])
            userEnd = ''.join([values[3], values[4], values[5]])
        except:
            pass
        window.close()
        if event in [None, '', 'Exit']:
            exitFlag = True
            break
        elif not ts.isValidDate(userStart) or not ts.isValidDate(userEnd):
            sg.Popup('Błędna data', title='Error')
            continue
        else:
            try:
                ts.monthlyReport((userStart, userEnd))
            except Exception as exc:
                sg.Popup(f'{exc}', title='Error')
                exitFlag = True
            break
