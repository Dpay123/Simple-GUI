import PySimpleGUI as sg
import pandas as pd

# add some color to the window
sg.theme('DarkTeal9')

# path must be directory-specific unless in the same directory
EXCEL_FILE = 'app.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Name', size=(15,1)), sg.InputText(key='Name')],
    [sg.Text('Contact', size=(15,1)), sg.Combo(['Phone', 'Email', 'Text'], key='Contact')],
    [sg.Text('Phone number', size=(15,1)), sg.InputText(key='Phone number')],
    [sg.Text('Email', size=(15,1)), sg.InputText(key='Email')],
    [sg.Text('Preferred Language', size=(15,1)),
        sg.Checkbox('English', key='English'),
        sg.Checkbox('Spanish', key='Spanish'),
        sg.Checkbox('Portuguese', key='Portuguese')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

# pass it to window
window = sg.Window('Simple data entry form', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)    
        sg.popup('Data saved')
        clear_input()
window.close()
