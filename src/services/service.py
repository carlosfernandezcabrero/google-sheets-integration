import gspread
import os
import csv
from oauth2client.service_account import ServiceAccountCredentials
from src.exceptions.not_content_returned_exception import NotContentReturnedException


class Service:
    def __init__(self):
        self.__client_secret_file = os.path.abspath('.key')
        self.__client = None
        self.__data = None

    def connect(self):
        credentials = ServiceAccountCredentials.from_json_keyfile_name(self.__client_secret_file)
        self.__client = gspread.authorize(credentials)

    def load_data(self, sheet, file_name):
        client = self.__client.open(file_name).get_worksheet(sheet)
        if client is None:
            raise NotContentReturnedException('No data found')
        else:
            self.__data = client.get_all_values('UNFORMATTED_VALUE')

    def extract_data_to_file(self, file_name):
        with open(file_name, 'w', encoding='UTF-8', newline='') as writer:
            csv.writer(writer).writerows(self.__data)
