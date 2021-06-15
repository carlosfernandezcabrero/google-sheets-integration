from src.exceptions.not_content_returned_exception import NotContentReturnedException
from tkinter import Button, Entry, Label, Tk, messagebox
from tkinter.constants import CENTER, W
from src.services.service import Service
from gspread import SpreadsheetNotFound


def __display_error(text_body):
    messagebox.showerror('ERROR', text_body)


def __is_empty(text):
    return len(text.strip()) == 0


def __download_file(path, sheet, file_name):
    path = path.strip()
    sheet = sheet.strip()
    file_name = file_name.strip()

    try:
        if __is_empty(path):
            __display_error('Output path cannot be empty')
        elif __is_empty(file_name):
            __display_error('File name cannot be empty')
        elif __is_empty(sheet):
            __display_error('Sheet cannot be empty')
        else:
            try:
                sheet = int(sheet)
                if sheet < 0:
                    __display_error('Sheet number cannot be less than 0')
                else:
                    try:
                        service = Service()
                        service.connect()
                        service.load_data(sheet, file_name)
                        service.extract_data_to_file(path)
                        messagebox.showinfo('COMPLETE', 'Download is complete')
                    except NotContentReturnedException as e:
                        messagebox.showwarning('WARNING', e.message)
            except ValueError:
                __display_error('The number of the sheet must be a number')
    except SpreadsheetNotFound:
        __display_error('The file does not exists')


class Setup:
    def __init__(self):
        self.__title = 'Google Sheets Integration'
        self.__window = None

    def load(self):
        self.__window = Tk()
        self.__window.title(self.__title)

    def make_body(self):
        self.__create_label('Path of output file', W, 0, 0)
        entry_path_to_extract = self.__create_entry(80, W, 0, 1)

        self.__create_label('Name of file', W, 1, 0)
        entry_file_name = self.__create_entry(80, W, 1, 1)

        self.__create_label('Sheet', W, 2, 0)
        entry_sheet = Entry(self.__window, width=4, justify=CENTER)
        entry_sheet.grid(row=2, column=1, padx=(30, 10), pady=(15, 0), sticky=W)

        button = Button(self.__window, text='Download',
                        command=lambda: __download_file(entry_path_to_extract.get(),
                                                        entry_sheet.get(),
                                                        entry_file_name.get()))
        button.grid(row=3, column=0, columnspan=2, pady=(30, 10))

    def display(self):
        self.__window.mainloop()

    def __create_label(self, label_text, position, label_row, label_column):
        label = Label(self.__window, text=label_text)
        label.grid(row=label_row, column=label_column, padx=(10, 0), pady=(15, 0), sticky=position)

    def __create_entry(self, label_width, label_position, label_row, label_column):
        entry = Entry(self.__window, width=label_width)
        entry.grid(row=label_row, column=label_column, padx=(30, 10), pady=(15, 0), sticky=label_position)
        return entry
