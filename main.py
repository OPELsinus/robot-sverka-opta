import datetime
import os
import shutil
import time
from contextlib import suppress
from copy import copy
from pathlib import Path

import pandas as pd
from time import sleep

import win32com.client as win32
import psycopg2 as psycopg2
from openpyxl import load_workbook
from pywinauto import keyboard

from config import download_path, robot_name, db_host, db_port, db_name, db_user, db_pass, tg_token, chat_id, logger, ecp_paths, mapping_path, template_path, owa_username, owa_password, months, months_normal, saving_path, smtp_host, smtp_author, homebank_login, homebank_password
from core import Sprut, Odines
from tools.app import App
from tools.clipboard import clipboard_get, clipboard_set
from tools.net_use import net_use
from tools.smtp import smtp_send
from tools.tg import tg_send
from tools.web import Web
from utils.homebank import homebank, check_homebank_and_collection
from utils.odines import odines_part, odines_check_with_collection
from utils.ofd import ofd_distributor
from utils.sprut_cashbook import open_cashbook


def create_collection_file(file_path):

    current_month: int = datetime.datetime.now().month
    current_year: int = datetime.datetime.now().year
    current_month_name = months_normal[current_month]

    main_working_file = None

    for item in os.listdir(saving_path):

        if current_month_name in str(item).lower():

            if "~$" in item:
                item = item.replace("~$", "")

            main_working_file = os.path.join(saving_path, item)

            break

    if not main_working_file:
        # * If there is no file related to current month
        file_name = f"Файл сбора {current_month_name.capitalize()} {current_year}.xlsx"
        main_working_file = os.path.join(saving_path, file_name)

        shutil.copy(template_path, main_working_file)

    collection_file = load_workbook(main_working_file)
    collection_sheet = collection_file['Файл сбора']
    print(f'Main Excel File: {main_working_file}')
    cols_dict = {
        'A': 'Компания',
        'B': 'Дата чека',
        'C': 'Дата и время чека',
        'D': 'Сумма с НДС',
        'E': 'Ерау',
        'F': '1с',
        'G': 'офд',
        'H': 'примечание',
        'I': '',
        'J': 'Номер чека',
        'K': 'Серийный № фиск.регистратора',
        'L': 'Клиент',
        'M': 'Дата создания записи',
        'N': 'Состояние розничного чека'
    }

    df = pd.read_excel(file_path)

    if df.columns[0] != 'Компания':
        df = pd.read_excel(file_path, header=1)

    for i, row in df.iterrows():

        last_row = collection_sheet.max_row + 1

        for col_key, col_name in cols_dict.items():

            source_cell = collection_sheet.cell(row=last_row - 1, column=collection_sheet[col_key + '1'].column)
            new_cell = collection_sheet.cell(row=last_row, column=collection_sheet[col_key + '1'].column)

            new_cell._style = copy(source_cell._style)
            new_cell.font = copy(source_cell.font)
            new_cell.border = copy(source_cell.border)
            new_cell.alignment = copy(source_cell.alignment)

            cell = collection_sheet[f'{col_key}{last_row}']
            try:
                print('#1', row[col_name])
                cell.value = row[col_name]
                cell.alignment = copy(source_cell.alignment)
            except:
                cell.value = None

    collection_file.save(main_working_file)

    return main_working_file


if __name__ == '__main__':

    today = datetime.datetime.today().strftime('%d.%m.%Y')
    today1 = datetime.datetime.today().strftime('%d.%m.%y')
    # today = '19.09.2023'
    # today1 = '19.09.23'
    calendar = pd.read_excel(f'\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-sverka-opta\\Производственный календарь 20{today1[-2:]}.xlsx')

    cur_day_index = calendar[calendar['Day'] == today1]['Type'].index[0]
    cur_day_type = calendar[calendar['Day'] == today1]['Type'].iloc[0]
    cur_weekday = calendar[calendar['Day'] == today1]['Weekday'].iloc[0]

    if cur_day_type != 'Holiday':
        # print('Started current date: ', yesterday2)
        _ = f"{today1.split('.')[0]}.{today1.split('.')[1]}.{today1.split('.')[2][-2:]}"
        weekends = [today]

        for i in range(cur_day_index - 1, 0, -1):
            if calendar['Type'].iloc[i] == 'Working':
                yesterday1 = calendar['Day'].iloc[i]
                break

            weekends.append(calendar['Day'].iloc[i][:6] + '20' + calendar['Day'].iloc[i][-2:])

        if len(weekends) > 1 or cur_weekday == 'Вт':
            sleep(12000)
            print('sleeped')
            logger.info(weekends)
        logger.info(weekends)
        print(weekends)
        for day in weekends:

            logger.warning(f'Started processing {day}')
            print(day)
            today = datetime.datetime.now().date()

            day_ = int(day.split('.')[0])
            month_ = int(day.split('.')[1])
            year_ = int(day.split('.')[2])
            today = datetime.date(year_, month_, day_)

            cashbook_day = (today - datetime.timedelta(days=5)).strftime('%d.%m.%Y')

            days = []

            for i in range(7, 1, -1):
                day = (today - datetime.timedelta(days=i)).strftime('%d.%m.%Y')
                days.append(day)

            today = today.strftime('%d.%m.%Y')
            logger.warning(today, cashbook_day)
            print(cashbook_day)
            print(days)
            # days.append('11.08.2023')
            # days = ['12.08.2023']
            logger.info(days)

            net_use(Path(template_path).parent, owa_username, owa_password)
            net_use(ecp_paths, owa_username, owa_password)

            # tg_send(f'Робот запустился - <b>{today}</b>\n\nДата для выгрузки чеков из Спрута - <b>{cashbook_day}</b>\n\nДата проверки в 1С - <b>{days}</b>', bot_token=tg_token, chat_id=chat_id)

            try:

                # * ----- 1 -----
                filepath = open_cashbook(cashbook_day)
                filepath = filepath.replace('Documents', 'Downloads') # If you are compiling for the virtual machines

                # * ----- 2 -----
                main_file = create_collection_file(filepath)
                Path(filepath).unlink()

                # * ----- 3 -----
                filepath = homebank(homebank_login, homebank_password, days[0], days[-1])

                # filepath = r'C:\Users\Abdykarim.D\Downloads\magnumopt_2023-09-07_2023-09-12.xlsx'
                # main_file = r'\\vault.magnum.local\Common\Stuff\_06_Бухгалтерия\Для робота\Процесс Сверка ОПТа\Файл сбора Сентябрь 2023.xlsx'
                check_homebank_and_collection(filepath, main_file)
                Path(filepath).unlink()

                logger.warning('Finished Epay')

                # * ----- 4 -----
                all_days = odines_part(days)

                odines_check_with_collection(all_days, main_file)
                logger.warning('Finished 1C')

                # # * ----- 5 -----
                ofd_distributor(main_file)

                logger.warning('Finished OFD')

                smtp_send(fr"""Добрый день!
                Сверка ОПТа за {today} завершилась успешно, файл сбора лежит в папке {main_file}""",
                          to=['Abdykarim.D@magnum.kz', 'Mukhtarova@magnum.kz', 'Sagimbayeva@magnum.kz', 'Ashirbayeva@magnum.kz'],
                          subject=f'Сверка ОПТа за {today}', username=smtp_author, url=smtp_host)

            except Exception as error:
                # smtp_send(fr"""Добрый день!
                #                Сверка ОПТа за {today} - ОШИБКА!!!""",
                #           to=['Abdykarim.D@magnum.kz', 'Mukhtarova@magnum.kz'],
                #           subject=f'ОШИБКА Сверка ОПТа за {today}', username=smtp_author, url=smtp_host)
                # tg_send(f'Возникла ошибка - {error}', bot_token=tg_token, chat_id=chat_id)
                raise error

    else:
        print(1)
