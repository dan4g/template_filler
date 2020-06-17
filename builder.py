import PySimpleGUI as sg
from docxtpl import DocxTemplate
import pandas as pd


def select_menu():
    sg.theme('Dark Blue 3')

    layout = [[sg.Text('Select template and data files')],

              [sg.Text('Choose template file (docx)', size=(20, 1)),
               sg.Input(key='template', default_text='template'), sg.FileBrowse()],

              [sg.Text('Choose data file (xlsx)', size=(20, 1)),
               sg.Input(key='data', default_text='data'), sg.FileBrowse()],

              [sg.Text('Enter resulting file name')],
              [sg.Text('Files will be named "Your_text file_number"')],
              [sg.Text('eg "File 1", "File 2" ...')],

              [sg.Input(key='name', default_text='Document # ')],

              [sg.Ok()],

              [sg.Exit()]]

    window = sg.Window('Main menu', layout)

    while True:
        event, values = window.read()

        if not event or event == 'Exit':
            window.close()
            break
        if event == 'Ok':
            template = values['template']
            data = values['data']
            name = values['name']
            if 'xlsx' or 'XLSX' in data:
                if 'docx' or 'DOCX' in template:
                    if name:
                        df = check_xlsx(data)
                        for i in df.index:
                            make_from_template(i, df, template, name)
                            sg.one_line_progress_meter('Progress', i, len(df), 'key')
                            sg.Popup('Ready!')
                    else:
                        sg.popup('Name error', 'Enter name')
                else:
                    sg.popup('No template selected', 'Select valid docx template')
            else:
                sg.popup('No data file selected', 'Select valid data file')


def make_from_template(number, df, template, name):
    doc = DocxTemplate(template)
    context = {df.columns[m]: df.iloc[number][m] for m in range(len(df.columns))}
    doc.render(context)
    doc.save(name + str(number) + '.docx')


def check_xlsx(data):
    df = pd.read_excel(data, sheet_name=0)
    return df


select_menu()
