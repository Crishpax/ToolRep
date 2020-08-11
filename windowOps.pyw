import pywinauto

def getToolingApp():
    app = pywinauto.application.Application(backend="uia").connect(title_re = r'.*PWSK.*')

    return app

def getExcelApp(greedy=False):
    app = pywinauto.application.Application(backend="win32").connect(title = 'Microsoft Excel')
    if app == None and greedy:
        app = pywinauto.application.Application(backend="win32").connect(title_re = r'.+Excel$')

    return app

def refuseExcelOpening(app):
    window = app.top_window()
    window['Button2'].click()

def runReportByMenuItemIndex(window, integer):

    window.menu_select('Raporty')
    window.descendants(control_type="MenuItem")[integer].select()

def closeOtherWindowsByTitle(app, desiredTitle):
    if len(app.windows()) == 1:
        return
    for window in app.windows():
        if desiredTitle not in str(window):
            window.close()

def runDocumentReport(window, dateStr):

    window.descendants()[42].select('Wszystkie')
    detailCheckbox = window.descendants()[6]
    if detailCheckbox.get_toggle_state() == 0:
        detailCheckbox.toggle()
    startField = window.descendants()[59]
    endField = window.descendants()[57]
    pane = window.descendants()[0]
    
    for field in [startField, endField]:
        field.click_input()
        pywinauto.keyboard.send_keys('{DEL}')
        for i in range(8):
            pywinauto.keyboard.send_keys('{BACKSPACE}')
        pywinauto.keyboard.send_keys(dateStr)
    
    paneRect = pane.get_properties()['rectangle']
    clickX = int(paneRect.width()/10)
    clickY = int(paneRect.height()/2)
    pane.click_input(button='left', coords=(clickX, clickY))

def runExpirationReport(window, dateRange):
    print(dateRange[0])
    #for child in window.descendants():
    #    print(child, window.descendants().index(child))
    #window.print_control_identifiers()
    window.descendants()[40].select('Wszystkie')
    
    startField = window.descendants()[55]
    endField = window.descendants()[50]

    pane = window.descendants()[0]
    
    startField.click_input()
    pywinauto.keyboard.send_keys('{DEL}')
    for i in range(8):
        pywinauto.keyboard.send_keys('{BACKSPACE}')
    pywinauto.keyboard.send_keys(dateRange[0])
    
    endField.click_input()
    pywinauto.keyboard.send_keys('{DEL}')
    for i in range(8):
        pywinauto.keyboard.send_keys('{BACKSPACE}')
    pywinauto.keyboard.send_keys(dateRange[1])

    paneRect = pane.get_properties()['rectangle']
    clickX = int(paneRect.width()/10)
    clickY = int(paneRect.height()/2)
    pane.click_input(button='left', coords=(clickX, clickY))
    
def saveToDesktop(window, filename):
    try:
        window['Pulpit'].click()
        window['Edit'].set_edit_text(filename)
        window['Zapisz'].click()
    except:
        window['Desktop'].click()
        window['Edit'].set_edit_text(filename)
        window['Save'].click()

