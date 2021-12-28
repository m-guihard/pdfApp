import PySimpleGUI as sg
import os.path
from PyPDF2 import PdfFileReader, PdfFileWriter


def change_layout(w, layout, event):
    w[layout].update(visible=False)
    layout = f"-{event.split(' ')[0].upper()}-"
    w[layout].update(visible=True)
    for button in menu_buttons:
        w[button].update(button_color='blue')
    w[event].update(button_color='green')
    w['Exit'].update(button_color='red')
    return layout


def merge(path_1, path_2, name_output):
    #print("Merging PDFs!")
    pdf = PdfFileReader(path_1)
    pdf_writer = PdfFileWriter()
    for page in range(pdf.getNumPages()):
        pdf_writer.addPage(pdf.getPage(page))
    pdf = PdfFileReader(path_2)
    for page in range(pdf.getNumPages()):
        pdf_writer.addPage(pdf.getPage(page))
    output = f'{name_output}.pdf'
    with open(output, 'wb') as output_pdf:
        pdf_writer.write(output_pdf)


def trim(path, name_output, start, end):
    #print("Trimming pages!")
    pdf = PdfFileReader(path)
    pdf_writer = PdfFileWriter()
    for page in range(pdf.getNumPages()):
        if page not in range(start, end+1):
            pdf_writer.addPage(pdf.getPage(page))
    output = f'{name_output}.pdf'
    with open(output, 'wb') as output_pdf:
        pdf_writer.write(output_pdf)


def move(path, name_output, start, end, new_pos):
    #print("Moving pages!")
    pdf = PdfFileReader(path)
    pdf_writer = PdfFileWriter()
    idx = [i for i in range(new_pos + 1) if i not in range(start, end + 1)]
    idx += [i for i in range(start, end + 1)]
    idx += [i for i in range(new_pos + 1, pdf.getNumPages()) if i not in range(start, end + 1)]
    for i in idx:
        pdf_writer.addPage(pdf.getPage(i))
    output = f'{name_output}.pdf'
    with open(output, 'wb') as output_pdf:
        pdf_writer.write(output_pdf)


def rotate(path, name_output, start, end, mode):
    #print("Rotating pages!")
    pdf = PdfFileReader(path)
    pdf_writer = PdfFileWriter()
    for page in range(pdf.getNumPages()):
        if page in range(start, end+1):
            if mode == 'C':
                pdf_writer.addPage(pdf.getPage(page).rotateClockwise(90))
            elif mode == 'CC':
                pdf_writer.addPage(pdf.getPage(page).rotateCounterClockwise(90))
            if mode == 'HALF':
                pdf_writer.addPage(pdf.getPage(page).rotateCounterClockwise(90).rotateCounterClockwise(90))
        else:
            pdf_writer.addPage(pdf.getPage(page))
    output = f'{name_output}.pdf'
    with open(output, 'wb') as output_pdf:
        pdf_writer.write(output_pdf)


def update_files_in_list(in_folder, key, w):
    try:
        file_list = os.listdir(in_folder)
    except:
        file_list = []

    fnames = [
        f for f in file_list
        if os.path.isfile(os.path.join(in_folder, f)) and f.lower().endswith((".pdf"))
    ]
    w[key].update(fnames)


def update_all_lists(fold):
    update_files_in_list(fold, "-FILE LIST ROTATE-", window)
    window['-FOLDER ROTATE-'].update(fold)
    update_files_in_list(fold, "-FILE LIST MOVE-", window)
    window['-FOLDER MOVE-'].update(fold)
    update_files_in_list(fold, "-FILE LIST TRIM-", window)
    window['-FOLDER TRIM-'].update(fold)
    update_files_in_list(fold, "-FILE LIST MERGE 1-", window)
    window['-FOLDER MERGE 1-'].update(fold)
    update_files_in_list(fold, "-FILE LIST MERGE 2-", window)
    window['-FOLDER MERGE 2-'].update(fold)


if __name__ == "__main__":

    # Define the different layouts
    layout_merge = [
        [
            sg.Column([
                [
                    sg.Text("File 1"),
                    sg.In(size=(25, 1), enable_events=True, key="-FOLDER MERGE 1-"),
                    sg.FolderBrowse(),
                ],
                [
                    sg.Listbox(values=[], enable_events=True, size=(40, 10), key="-FILE LIST MERGE 1-")
                ],
                [
                    sg.Text("File 2"),
                    sg.In(size=(25, 1), enable_events=True, key="-FOLDER MERGE 2-"),
                    sg.FolderBrowse(),
                ],
                [
                    sg.Listbox(values=[], enable_events=True, size=(40, 10), key="-FILE LIST MERGE 2-")
                ]
            ]),
            sg.VSeperator(),
            sg.Column([
                [sg.Text("No file selected", key="TXT SELECTED FILE MERGE 1")],
                [sg.Text("No file selected", key="TXT SELECTED FILE MERGE 2")],
                [sg.Text("Output file name:")],
                [sg.In(size=(20, 1), enable_events=False, key="-MERGE OUTPUT NAME-")],
                [sg.Button("Merge")]
            ]),
        ]
    ]
    layout_trim = [
        [
            sg.Column([
                [
                    sg.Text("Folder"),
                    sg.In(size=(25, 1), enable_events=True, key="-FOLDER TRIM-"),
                    sg.FolderBrowse(),
                ],
                [
                    sg.Listbox(values=[], enable_events=True, size=(40, 20), key="-FILE LIST TRIM-")
                ],
                [
                    sg.Text("Enter a range of pages to delete:"),
                    sg.In(size=(5, 1), enable_events=False, key="-TRIM START-"),
                    sg.In(size=(5, 1), enable_events=False, key="-TRIM END-"),
                ]
            ]),
            sg.VSeperator(),
            sg.Column([
                [sg.Text("No file selected", key="TXT SELECTED FILE TRIM")],
                [sg.Text("Output file name:")],
                [sg.In(size=(20, 1), enable_events=False, key="-TRIM OUTPUT NAME-")],
                [sg.Button("Remove pages")]
            ]),
        ]
    ]
    layout_move = [
        [
            sg.Column([
                [
                    sg.Text("Folder"),
                    sg.In(size=(25, 1), enable_events=True, key="-FOLDER MOVE-"),
                    sg.FolderBrowse(),
                ],
                [
                    sg.Listbox(values=[], enable_events=True, size=(40, 20), key="-FILE LIST MOVE-")
                ],
                [
                    sg.Text("Enter a range of pages to move:"),
                    sg.In(size=(5, 1), enable_events=False, key="-MOVE START-"),
                    sg.In(size=(5, 1), enable_events=False, key="-MOVE END-"),
                ],
                [
                    sg.Text("Enter the page after which they should be placed:"),
                    sg.In(size=(5, 1), enable_events=False, key="-MOVE PLACE-")
                ]
            ]),
            sg.VSeperator(),
            sg.Column([
                [sg.Text("No file selected", key="TXT SELECTED FILE MOVE")],
                [sg.Text("Output file name:")],
                [sg.In(size=(20, 1), enable_events=False, key="-MOVE OUTPUT NAME-")],
                [sg.Button("Move selected pages")]
            ]),
        ]
    ]
    layout_rotate = [
        [
            sg.Column([
                [
                    sg.Text("Folder"),
                    sg.In(size=(25, 1), enable_events=True, key="-FOLDER ROTATE-"),
                    sg.FolderBrowse(),
                ],
                [
                    sg.Listbox(values=[], enable_events=True, size=(40, 15), key="-FILE LIST ROTATE-")
                ],
                [
                    sg.Text("Enter a range of pages to rotate:"),
                    sg.In(size=(5, 1), enable_events=False, key="-ROTATE START-"),
                    sg.In(size=(5, 1), enable_events=False, key="-ROTATE END-"),
                ],
                [
                    sg.Text("How should they be rotated?"),
                    sg.Column([
                        [sg.Radio('90 deg clockwise', "RADIO1", default=True, key="-CLOCKWISE-")],
                        [sg.Radio('90 deg clockwise', "RADIO1", default=False, key="-CCLOCKWISE-")],
                        [sg.Radio('180 deg ', "RADIO1", default=False, key="-HALF-")]
                    ])
                ]
            ]),
            sg.VSeperator(),
            sg.Column([
                [sg.Text("No file selected", key="TXT SELECTED FILE ROTATE")],
                [sg.Text("Output file name:")],
                [sg.In(size=(20, 1), enable_events=False, key="-ROTATE OUTPUT NAME-")],
                [sg.Button("Rotate selected pages")]
            ]),
        ]
    ]
    menu_buttons = ['Merge PDFs', 'Trim PDF', 'Move pages', 'Rotate pages', 'Exit']
    layout = [
        [
            sg.Button(menu_buttons[0], button_color='green'),
            sg.Button(menu_buttons[1], button_color='blue'),
            sg.Button(menu_buttons[2], button_color='blue'),
            sg.Button(menu_buttons[3], button_color='blue'),
            sg.Button(menu_buttons[4], button_color='red')
        ],
        [
            sg.Column(layout_merge, visible=True, key='-MERGE-'),
            sg.Column(layout_trim, visible=False, key='-TRIM-'),
            sg.Column(layout_move, visible=False, key='-MOVE-'),
            sg.Column(layout_rotate, visible=False, key='-ROTATE-')
        ]
    ]

    # Create window and run main loop
    window = sg.Window("PDF App", layout, size=(600, 500))
    current_layout = '-MERGE-'

    file_to_trim_pages = [1,2,3,4]
    filename_trim = ""

    while True:
        event, values = window.read()
        if event == "Exit" or event == sg.WIN_CLOSED:
            break
        elif event in menu_buttons:
            current_layout = change_layout(window, current_layout, event)
# MERGE
        elif event == "-FOLDER MERGE 1-":
            #
            update_all_lists(values[event])
        elif event == "-FOLDER MERGE 2-":
            #
            update_files_in_list(values[event], "-FILE LIST MERGE 2-", window)
        elif event == "Merge":

            if values['-FILE LIST MERGE 1-'] != [] and values['-FILE LIST MERGE 2-'] != []:
                merge(os.path.join(values["-FOLDER MERGE 1-"], values["-FILE LIST MERGE 1-"][0]),
                      os.path.join(values["-FOLDER MERGE 2-"], values["-FILE LIST MERGE 2-"][0]),
                      os.path.join(values["-FOLDER MERGE 1-"], values["-MERGE OUTPUT NAME-"]))
                popup = sg.Window("Image Viewer", [
                    [
                        sg.Column([
                            [sg.Text(
                                f"Successfully merged \n{values['-FILE LIST MERGE 1-'][0]} and {values['-FILE LIST MERGE 2-'][0]} \ninto {values['-MERGE OUTPUT NAME-']}.pdf!")],
                            [sg.Button("OK")]
                        ])
                    ]
                ], size=(300, 150))
                while True:
                    ev, _ = popup.read()
                    if ev == "Exit" or ev == sg.WIN_CLOSED or ev == "OK":
                        break
                popup.close()
                update_files_in_list(values["-FOLDER MERGE 1-"], "-FILE LIST MERGE 1-", window)
                update_files_in_list(values["-FOLDER MERGE 2-"], "-FILE LIST MERGE 2-", window)
        elif event == "-FILE LIST MERGE 1-":
            #
            window["TXT SELECTED FILE MERGE 1"].update(f'File 1: {values["-FILE LIST MERGE 1-"][0]}')
        elif event == "-FILE LIST MERGE 2-":
            #
            window["TXT SELECTED FILE MERGE 2"].update(f'File 2: {values["-FILE LIST MERGE 2-"][0]}')
# TRIM
        elif event == "-FOLDER TRIM-":
            #
            update_all_lists(values[event])
        elif event == "Remove pages":
            if filename_trim != "":
                start_trim = values["-TRIM START-"]
                end_trim = values["-TRIM END-"]
                if len(values["-FILE LIST TRIM-"]) == 0:
                    continue
                if len(start_trim) == 0 and len(end_trim) == 0:
                    continue
                if len(start_trim) == 0:
                    start_trim = 1
                if len(end_trim) == 0:
                    end_trim = PdfFileReader(filename_trim).getNumPages()
                start_trim, end_trim = int(start_trim) - 1, int(end_trim) - 1
                if end_trim >= PdfFileReader(filename_trim).getNumPages():
                    end_trim = PdfFileReader(filename_trim).getNumPages()-1
                if start_trim < 0:
                    start_trim = 0
                if start_trim > end_trim:
                    continue
                trim(filename_trim, os.path.join(values["-FOLDER TRIM-"], values["-TRIM OUTPUT NAME-"]), start_trim, end_trim)
                popup = sg.Window("Image Viewer", [
                    [
                        sg.Column([
                            [sg.Text(f"Successfully created file {values['-TRIM OUTPUT NAME-']}.pdf \n with pages {start_trim+1} to {end_trim+1} removed!")],
                            [sg.Button("OK")]
                        ])
                    ]
                ], size=(300, 100))
                while True:
                    ev, _ = popup.read()
                    if ev == "Exit" or ev == sg.WIN_CLOSED or ev == "OK":
                        break
                popup.close()
                update_files_in_list(values["-FOLDER TRIM-"], "-FILE LIST TRIM-", window)
        elif event == "-FILE LIST TRIM-":
            #
            window["TXT SELECTED FILE TRIM"].update(f'File: {values["-FILE LIST TRIM-"][0]}')
            filename_trim = os.path.join(values["-FOLDER TRIM-"], values["-FILE LIST TRIM-"][0])
# MOVE
        elif event == "-FOLDER MOVE-":
            #
            update_all_lists(values[event])
        elif event == "Move selected pages":
            start_move = values["-MOVE START-"]
            end_move = values["-MOVE END-"]
            new_pos = values['-MOVE PLACE-']
            if len(values["-FILE LIST MOVE-"]) == 0:
                continue
            if len(start_move) == 0 and len(end_move) == 0:
                continue
            if len(new_pos) == 0 or int(new_pos) < 0:
                continue
            if len(start_move) == 0:
                start_move = 1
            if len(end_move) == 0:
                end_move = PdfFileReader(os.path.join(values["-FOLDER MOVE-"], values["-FILE LIST MOVE-"][0])).getNumPages()
            start_move, end_move, new_pos = int(start_move) - 1, int(end_move) - 1, int(new_pos) - 1
            if new_pos > PdfFileReader(os.path.join(values["-FOLDER MOVE-"], values["-FILE LIST MOVE-"][0])).getNumPages()-1:
                new_pos = PdfFileReader(os.path.join(values["-FOLDER MOVE-"], values["-FILE LIST MOVE-"][0])).getNumPages()-1
            if end_move >= PdfFileReader(os.path.join(values["-FOLDER MOVE-"], values["-FILE LIST MOVE-"][0])).getNumPages():
                end_move = PdfFileReader(os.path.join(values["-FOLDER MOVE-"], values["-FILE LIST MOVE-"][0])).getNumPages()-1
            if start_move < -1:
                start_move = -1
            if start_move > end_move:
                continue
            move(os.path.join(values["-FOLDER MOVE-"], values["-FILE LIST MOVE-"][0]),
                 os.path.join(values["-FOLDER MOVE-"], values["-MOVE OUTPUT NAME-"]),
                 start_move, end_move, new_pos)
            popup = sg.Window("Image Viewer", [
                [
                    sg.Column([
                        [sg.Text(
                            f"Successfully moved pages!")],
                        [sg.Button("OK")]
                    ])
                ]
            ], size=(200, 100))
            while True:
                ev, _ = popup.read()
                if ev == "Exit" or ev == sg.WIN_CLOSED or ev == "OK":
                    break
            popup.close()
            update_files_in_list(values["-FOLDER MOVE-"], "-FILE LIST MOVE-", window)
        elif event == "-FILE LIST MOVE-":
            #
            window["TXT SELECTED FILE MOVE"].update(f'File: {values["-FILE LIST MOVE-"][0]}')
# ROTATE
        elif event == "-FOLDER ROTATE-":
            #
            update_all_lists(values[event])
        elif event == "Rotate selected pages":
            start_rotate = values["-ROTATE START-"]
            end_rotate = values["-ROTATE END-"]
            if len(values["-FILE LIST ROTATE-"]) == 0:
                continue
            if len(start_rotate) == 0 and len(end_rotate) == 0:
                continue
            if len(start_rotate) == 0:
                start_rotate = 1
            if len(end_rotate) == 0:
                end_rotate = PdfFileReader(
                    os.path.join(values["-FOLDER ROTATE-"], values["-FILE LIST ROTATE-"][0])).getNumPages()
            start_rotate, end_rotate = int(start_rotate) - 1, int(end_rotate) - 1
            if end_rotate >= PdfFileReader(
                    os.path.join(values["-FOLDER ROTATE-"], values["-FILE LIST ROTATE-"][0])).getNumPages():
                end_rotate = PdfFileReader(
                    os.path.join(values["-FOLDER ROTATE-"], values["-FILE LIST ROTATE-"][0])).getNumPages() - 1
            if start_rotate < -1:
                start_rotate = -1
            if start_rotate > end_rotate:
                continue
            if values["-CLOCKWISE-"] == True:
                mode = 'C'
            elif values["-CCLOCKWISE-"] == True:
                mode = 'CC'
            else:
                mode = 'HALF'
            rotate(os.path.join(values["-FOLDER ROTATE-"], values["-FILE LIST ROTATE-"][0]),
                   os.path.join(values["-FOLDER ROTATE-"], values["-ROTATE OUTPUT NAME-"]),
                   start_rotate, end_rotate, mode)
            popup = sg.Window("Image Viewer", [
                [
                    sg.Column([
                        [sg.Text(
                            f"Successfully rotated pages {start_rotate+1} to {end_rotate+1}!")],
                        [sg.Button("OK")]
                    ])
                ]
            ], size=(300, 100))
            while True:
                ev, _ = popup.read()
                if ev == "Exit" or ev == sg.WIN_CLOSED or ev == "OK":
                    break
            popup.close()
            update_files_in_list(values["-FOLDER ROTATE-"], "-FILE LIST ROTATE-", window)
        elif event == "-FILE LIST ROTATE-":
            #
            window["TXT SELECTED FILE ROTATE"].update(f'File: {values["-FILE LIST ROTATE-"][0]}')

    window.close()