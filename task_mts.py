import logging
import os
import tkinter as tk
from time import sleep
from tkinter import filedialog

import pandas as pd
import requests
import xlsxwriter
from dotenv import load_dotenv

load_dotenv()

logging.basicConfig(level=logging.DEBUG)


class FsspAPI:
    """
    Base class for access to API ФССП
    docs https://api-ip.fssp.gov.ru/
    """
    FSSP_URL = 'https://fssp.gov.ru/'
    URL_FSSP_API = 'https://api-ip.fssp.gov.ru/api/v1.0/'
    REGIONS = tuple([i for i in range(1, 79)] + [86, 89, 91, 92])

    def __init__(self, strftime_mask: str = '%d.%m.%Y'):
        self.strftime_mask = strftime_mask
        self.fssp_tasks = []
        self.fssp_data = []

    def teimestamp_to_string(self, timestamp, mask: str = None) -> str:  # noqa
        """
        Timestamp pandas to string
        :param timestamp: Timestamp typ of pandas
        :param mask: mask to convert timestamp
        :return str: datetime in string
        """
        timestamp.to_pydatetime()
        return (timestamp.strftime('%d.%m.%Y')
                if mask is None else timestamp.strftime(mask))

    def get_full_fio(self, obj: list) -> str:  # noqa
        """Get ful last name first name patronymic as string"""
        return f'{obj[0]} {obj[1]} {obj[2]}'

    def get_full_fio_and_birthday(self, obj: list, mask_datetime: str = None) -> str:
        """
        Get ful last name first name patronymic and birthday as string
        :param obj: list with attributes
        :return str:
        """
        if len(obj) > 3:
            fio = self.get_full_fio(obj)
            birthday = self.teimestamp_to_string(obj[3], mask_datetime)
            return '{} {}'.format(fio, birthday)
        return self.get_full_fio(obj)

    def api_full_url(self, add_url: str) -> str:
        """Get full url for api source
        :param add_url: additional url value
        :return str: full url to access url
        """
        return self.URL_FSSP_API + add_url

    def get_status_api_fssp(self, task: str = None):  # noqa
        """
        Get task status from API. Call method get_result_api_fssp()
        if code status in allowed
        """
        total_url = self.api_full_url('status')
        payload = {'token': os.getenv('FSSP_TOKEN')}

        for i in range(len(self.fssp_tasks)):

            task = self.fssp_tasks.pop()
            payload.update({'task': task})
            r = requests.get(total_url, params=payload)

            if r.status_code == 200:
                # status == 0 means task is successful
                status = r.json()['code']

                if status == 0 or status == 3:
                    self.get_result_api_fssp(task)

    def get_result_api_fssp(self, task: str) -> None:
        """
        Get result of task from API
        :param task: unique task value
        :return None:
        """
        total_url = self.api_full_url('result')
        payload = {'token': os.getenv('FSSP_TOKEN'),
                   'task': task}
        r = requests.get(total_url, payload)
        if r.status_code == 200:
            result = r.json()['response']['result']

            for element in result:
                element_result = element['result']
                if element_result:
                    for result in element_result:
                        self.fssp_data.append({
                            'ФИО': result['name'],
                            'Приказ': result['exe_production'],
                            'Детально': result['details'],
                            'Цель': result['subject'],
                            'Пристав': result['bailiff'],
                            'Доп': result['ip_end'],

                        })

    def post_search_group(self, excel_data: list) -> None:
        """
        Send POST request to API

        type=1 - Отправить запрос на поиск физического лица;
        type=2 - Отправить запрос на поиск юридического лица;
        type=3 - Отправить запрос на поиск юридического лица;

        :param excel_data: list с элементами на проверку
        :return: None
        """
        if excel_data is None:
            return
        total_url = self.api_full_url('search/group')

        data_params = []
        data = {
            'token': os.getenv('FSSP_TOKEN'),
        }

        for person in excel_data:
            last_name = person[0]
            first_name = person[1]
            second_name = person[2]
            birthdate = self.teimestamp_to_string(person[3])
            for region in self.REGIONS:
                data_params.append({
                    'type': 1,  # see docs to method
                    'params': {
                        'region': region,
                        'firstname': first_name,
                        'lastname': last_name,
                        'secondname': second_name,
                        'birthdate': birthdate,
                    }
                })

                # request can't be greater than 50 in len() or if region last in list
                if len(data_params) >= 49 or region == self.REGIONS[len(self.REGIONS) - 1]:
                    data['request'] = data_params

                    while True:
                        r = requests.post(total_url, json=data)
                        if r.status_code == 200:
                            self.fssp_tasks.append(r.json()['response']['task'])
                            data['request'].clear()
                            logging.info('Successfully request')
                            sleep(5)
                            break
                        elif (r.status_code == 429
                              and r.json()['exception'] == 'Дождитесь результата предыдущего группового запроса'):
                            logging.info('Need wait. r.status_code == 429')
                            sleep(5)

                        # minimal time for waiting between requests is 5 sek
                        sleep(5)


class Application(tk.Frame, FsspAPI):

    def __init__(self, root=None):
        tk.Frame.__init__(self, root)
        FsspAPI.__init__(self)

        self.excel_data = None

        # Tkinter GUI
        self.root = root
        self.canvas = tk.Canvas(self.root, width=800, height=500)
        self.create_widgets()

    def create_widgets(self):
        self.canvas.pack()

        self.label = tk.Label(root, text='Специально для МТС Банка')
        self.label.config(font=('Arial', 20))
        self.canvas.create_window(400, 50, window=self.label)

        # Tkinter GUI

        # Buttons
        # Загрузка excel
        self.button_load_excel = tk.Button(text='Загрузите файл excel', command=self.load_excel,
                                           bg='green', fg='white', font=('helvetica', 12, 'bold'))
        self.canvas.create_window(400, 180, window=self.button_load_excel)

        # Courts
        # https://sudrf.ru/index.php?id=300#sp
        self.button_court = tk.Button(self.root, text='Запрос в суды',
                                      # command=self.run_court,
                                      bg='red',
                                      font=('helvetica', 11, 'bold'))
        self.canvas.create_window(400, 220, window=self.button_court)

        # ФССП
        self.button_fssp = tk.Button(self.root, text='Запрос в ФССП',
                                     command=self.run_api_fssp,
                                     bg='red',
                                     font=('helvetica', 11, 'bold'))
        self.canvas.create_window(400, 260, window=self.button_fssp)

        # exit
        self.button_exit = tk.Button(self.root,
                                     text='Выход',
                                     command=self.root.destroy,
                                     bg='green',
                                     font=('helvetica', 11, 'bold'))
        self.canvas.create_window(400, 300, window=self.button_exit)

    def load_excel(self) -> None:
        """
        Load and parse excel. Launch by button
        """
        import_file_path = filedialog.askopenfilename()
        df = pd.read_excel(import_file_path,
                           engine="odf",
                           usecols=['Фамилия', 'Имя', 'Отчество', 'Дата рождения'])
        self.excel_data = df.values

    def create_excel(self, data: list, file_name: str) -> None:  # noqa
        """
        Create excel file in the same directory, as current file

        example of data to load:
        data = [{'key1': 'value1', 'key2': 'value2'... }]

        :param data: list with dict values to write into excel
        :param file_name: str full file name with extension. Only excel format excepted

        :return: None
        """
        directory_path = os.getcwd()
        workbook = xlsxwriter.Workbook(directory_path + '/' + file_name)
        worksheet = workbook.add_worksheet()

        keys = list(data[0].keys())
        keys.sort()

        # create headers
        for idx, key in enumerate(keys):
            worksheet.write(0, idx, key)

        row = 0
        for element in data:
            row += 1
            for idx, key in enumerate(keys):
                worksheet.write(row, idx, element[key])

        workbook.close()
        logging.info('Excel created!')

    def run_api_fssp(self) -> None:  # noqa
        """
        Send request to API, get data and create excel file.
        Working by push button
        """
        if self.excel_data:
            self.post_search_group(self.excel_data)
            self.get_status_api_fssp()
            self.create_excel(data=self.fssp_data, file_name='result_fssp_api.xlsx')
        logging.debug('Сначала необходио загрузить excel файл')
        return


if __name__ == '__main__':
    root = tk.Tk()
    app = Application(root=root)
    app.mainloop()
