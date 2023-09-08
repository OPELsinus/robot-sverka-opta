import datetime
import os
import shutil
import time
from contextlib import suppress
from copy import copy

import pandas as pd
from time import sleep

import win32com.client as win32
import psycopg2 as psycopg2
from openpyxl import load_workbook
from pywinauto import keyboard

from config import download_path, robot_name, db_host, db_port, db_name, db_user, db_pass, tg_token, chat_id, logger, ecp_paths, mapping_path
from core import Sprut, Odines
from tools.app import App
from tools.clipboard import clipboard_get, clipboard_set
from tools.tg import tg_send
from tools.web import Web

months = [
    'Январь',
    'Февраль',
    'Март',
    'Апрель',
    'Май',
    'Июнь',
    'Июль',
    'Август',
    'Сентябрь',
    'Октябрь',
    'Ноябрь',
    'Декабрь'
]


def sql_create_table():
    conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)
    table_create_query = f'''
        CREATE TABLE IF NOT EXISTS ROBOT.{robot_name.replace("-", "_")} (
            started_time timestamp,
            ended_time timestamp,
            store_name text,
            status text,
            found_difference text,
            count int,
            error_reason text,
            error_saved_path text,
            execution_time text
            )
        '''
    c = conn.cursor()
    c.execute(table_create_query)

    conn.commit()
    c.close()
    conn.close()


def insert_data_in_db(started_time, store_name, status, found_difference, count, error_reason, error_saved_path, execution_time):
    conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)

    query_delete = f"""
        delete from ROBOT.{robot_name.replace("-", "_")} where store_name = '{store_name}'
    """

    query = f"""
        INSERT INTO ROBOT.{robot_name.replace("-", "_")}
        (started_time, ended_time, store_name, status, found_difference, count, error_reason, error_saved_path, execution_time)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
    """

    values = (
        started_time,
        datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S.%f"),
        store_name,
        status,
        found_difference,
        count,
        error_reason,
        error_saved_path,
        execution_time
    )

    cursor = conn.cursor()

    conn.autocommit = True
    try:
        cursor.execute(query_delete)
    except Exception as e:
        print('GOVNO', e)
        pass

    try:
        cursor.execute(query, values)

    except Exception as e:
        conn.rollback()
        print(f"Error: {e}")

    conn.commit()

    cursor.close()
    conn.close()


def get_all_data():
    conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)
    table_create_query = f'''
            SELECT * FROM ROBOT.{robot_name.replace("-", "_")}
            order by started_time desc
            '''
    cur = conn.cursor()
    cur.execute(table_create_query)

    df1 = pd.DataFrame(cur.fetchall())
    df1.columns = ['started_time', 'ended_time', 'store_name', 'status', 'found_difference', 'count', 'error_reason', 'error_saved_path', 'execution_time']

    cur.close()
    conn.close()

    return df1


def open_cashbook(today):

    sprut = Sprut("REPS")
    sprut.run()

    sprut.open("Кассовая книга", switch=False)

    print('Switching')
    sprut.parent_switch({"title_re": ".Кассовая книга.", "class_name": "Tbo_cashbook_fm_main",
                         "control_type": "Window", "visible_only": True, "enabled_only": True, "found_index": 0}).set_focus()
    print('Switched')
    sprut.find_element({"title": "Приложение", "class_name": "", "control_type": "MenuBar",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click(coords=(50, 17))

    sprut.find_element({"title": "", "class_name": "", "control_type": "MenuItem",
                        "visible_only": True, "enabled_only": True, "found_index": 3}).click()

    sprut.find_element({"title": "", "class_name": "", "control_type": "MenuItem",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.find_element({"title": "Последний использованный фильтр", "class_name": "TvmsToolGridQueryList", "control_type": "Pane",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click(coords=(380, 17))

    print('HERE1')
    sprut.parent_switch({"title": "Выборка по запросу", "class_name": "Tvms_modifier_fm_builder", "control_type": "Window",
                         "visible_only": True, "enabled_only": True, "found_index": 0}).set_focus()
    print('HERE2')

    try:
        sprut.find_element({"title": "Клиент", "class_name": "", "control_type": "ListItem",
                            "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=5).click()
    except:
        sprut.find_element({"title": "", "class_name": "TvmsListBox", "control_type": "Pane",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).click()

        sprut.find_element({"title": "", "class_name": "TvmsListBox", "control_type": "Pane",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(sprut.keys.PAGE_DOWN)
    sprut.find_element({"title": "Клиент", "class_name": "", "control_type": "ListItem",
                        "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=5).click()
    print('kek')
    sprut.find_element({"title": "", "class_name": "TvmsDBTelePusik", "control_type": "Pane",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.parent_switch({"title": "Поиск", "class_name": "Tvms_search_fm_builder", "control_type": "Window",
                         "visible_only": True, "enabled_only": True, "found_index": 0}).set_focus()

    sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys('^N')

    sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(f'%Розничный покупатель ОПТ%', sprut.keys.ENTER, protect_first=True)

    sprut.find_element({"title": "", "class_name": "", "control_type": "Button",
                        "visible_only": True, "enabled_only": True, "found_index": 1}).click()

    sprut.find_element({"title": "Выбрать", "class_name": "TvmsBitBtn", "control_type": "Button",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()
    sprut.parent_back(1)

    while True:
        try:
            sprut.find_element({"title": "", "class_name": "TvmsBitBtn", "control_type": "Button",
                                "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1).click()
            break
        except:
            pass
    print('clicked')

    sprut.find_element({"title": "И", "class_name": "TvmsBitBtn", "control_type": "Button",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    try:
        sprut.find_element({"title": "Дата чека", "class_name": "", "control_type": "ListItem",
                            "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=5).click()
    except:
        sprut.find_element({"title": "", "class_name": "TvmsListBox", "control_type": "Pane",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).click()

        sprut.find_element({"title": "", "class_name": "TvmsListBox", "control_type": "Pane",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(sprut.keys.PAGE_UP)
    sprut.find_element({"title": "Дата чека", "class_name": "", "control_type": "ListItem",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                        "visible_only": True, "enabled_only": True, "found_index": 1}).click()
    sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                        "visible_only": True, "enabled_only": True, "found_index": 1}).type_keys('05.08.2023')

    sprut.find_element({"title_re": ".", "class_name": "TvmsComboBox", "control_type": "Pane",
                        "visible_only": True, "enabled_only": True, "found_index": 1}).click()
    sprut.find_element({"title_re": ".", "class_name": "TvmsComboBox", "control_type": "Pane",
                        "visible_only": True, "enabled_only": True, "found_index": 1}).click()
    sprut.find_element({"title_re": ".", "class_name": "TvmsComboBox", "control_type": "Pane",
                        "visible_only": True, "enabled_only": True, "found_index": 1}).type_keys('^N', sprut.keys.ENTER)
    print('Clicked list')
    # keyboard.send_keys("{UP}" * 4)
    # keyboard.send_keys("{ENTER}")

    print('Clicked item')
    # sprut.find_element({"title": "", "class_name": "TvmsBitBtn", "control_type": "Button",
    #                     "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    while True:
        try:
            sprut.find_element({"title": "", "class_name": "TvmsBitBtn", "control_type": "Button",
                                "visible_only": True, "enabled_only": True, "found_index": 1}, timeout=1).click()
            break
        except:
            pass

    while True:
        try:
            sprut.find_element({"title": "Ввод", "class_name": "TvmsBitBtn", "control_type": "Button",
                                "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1).click()
            break
        except:
            pass
    print('clicked')

    sprut.parent_back(1)

    print('1')
    sprut.find_element({"title": "", "class_name": "TvmsDBToolGrid", "control_type": "Pane",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.find_element({"title": "", "class_name": "TvmsDBToolGrid", "control_type": "Pane",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys('^%E')

    sprut.parent_switch({"title": "Экспортировать данные", "class_name": "Tvms_fm_DBExportExt", "control_type": "Window",
                         "visible_only": True, "enabled_only": True, "found_index": 0})

    sprut.find_element({"title": "", "class_name": "TcxCustomDropDownInnerEdit", "control_type": "Edit",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.find_element({"title": "", "class_name": "TcxCustomDropDownInnerEdit", "control_type": "Edit",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(sprut.keys.DOWN, sprut.keys.ENTER)

    file_path = sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                                    "visible_only": True, "enabled_only": True, "found_index": 1}).get_text()
    print(file_path)

    sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit", "visible_only": True, "enabled_only": True, "found_index": 0}).set_text('')

    right_pane = {"title": "", "class_name": "TvmsListBox", "control_type": "Pane", "visible_only": True, "enabled_only": True, "found_index": 0}

    for i in range(10):
        try:
            sprut.find_element({"title": "Срочное проведения чека?", "class_name": "", "control_type": "ListItem",
                                "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=3).click()

            sprut.find_element({"title": "", "class_name": "", "control_type": "Button",
                                "visible_only": True, "enabled_only": True, "found_index": 1}).click()
            break
        except:
            sprut.find_element(right_pane).click()
            sprut.find_element(right_pane).type_keys(sprut.keys.PAGE_DOWN)

    print()

    sprut.find_element({"title": "Ввод", "class_name": "TvmsFooterButton", "control_type": "Button",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    wait_loading(file_path)

    sprut.quit()

    return file_path


def wait_loading(filepath):
    print('Started loading')
    while True:
        if os.path.isfile(filepath):
            print('downloaded')
            break
    print('Finished loading')
    sleep(3)


def check_if_time_diff_less_than_1_min(first_date, second_date):
    try:
        first_date = datetime.datetime.strptime(first_date, '%d.%m.%Y %H:%M:%S')
    except:
        pass

    try:
        second_date = datetime.datetime.strptime(second_date, '%d.%m.%Y %H:%M')
    except:
        try:
            second_date = datetime.datetime.strptime(second_date, '%d.%m.%Y %H:%M:%S')
        except Exception as e:
            print("GOVNO", e)
            pass
        pass

    # print(first_date, second_date)
    # print((first_date - second_date).total_seconds() // 60)

    return abs((first_date - second_date).total_seconds() // 60)


def create_collection_file(file_path):

    collection_file = load_workbook(r'C:\Users\Abdykarim.D\Documents\Файл сбора.xlsx')
    collection_sheet = collection_file['Файл сбора']

    df = pd.read_excel(file_path, header=1)

    print(df.columns)
    print(df)
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

    for i, row in df.iterrows():

        last_row = collection_sheet.max_row + 1

        for col_key, col_name in cols_dict.items():
            # print(col_key, col_name)
            previous_row = collection_sheet[last_row - 1]
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

    columns = ['Компания', 'Дата чека', 'Дата и время чека', 'Сумма с НДС', 'Ерау', '1с', 'офд ', 'примечание', '', 'Номер чека', 'Серийный № фиск.регистратора', 'Клиент', 'Дата создания записи', 'Состояние розничного чека']
    # collection_file = collection_file[columns]
    # print(columns, len(columns))
    collection_file.save(r'C:\Users\Abdykarim.D\Documents\Файл сбора1.xlsx')


def homebank(email, password):
    web = Web()
    web.run()
    web.get('https://epay.homebank.kz/login')

    web.wait_element('//*[@id="mp-content"]/section/main/div[2]/div/div/div[2]/form/div[1]/div/div/span/div/input')

    sleep(5)

    web.find_element('//*[@id="mp-content"]/section/main/div[2]/div/div/div[2]/form/div[1]/div/div/span/div/input').type_keys(email)
    web.find_element('//*[@id="mp-content"]/section/main/div[2]/div/div/div[2]/form/div[2]/div/div/span/div/span/input').type_keys(password)

    web.find_element('//*[@id="mp-content"]/section/main/div[2]/div/div/div[2]/form/div[3]/div/div/span/button').click()

    web.wait_element("//span[@class='src-layouts-main-header_button hint-section-1-step-3']")

    web.get('https://epay.homebank.kz/statements/payment')

    web.find_element("//span[contains(text(), '427693/14-EC27/07')]").click()

    web.find_element('//*[@id="mp-content"]/div/div/div/div/div[1]/div/div/div[1]/div/div/div/div[2]/button').click()

    web.find_element('//*[@id="period"]').click()

    # ? fix
    web.find_element("//td[@title = '31 августа 2023 г.']").click()


def check_homebank_and_collection():

    collection_file = load_workbook(r'C:\Users\Abdykarim.D\Documents\Файл сбора1.xlsx')

    collection_sheet = collection_file['Файл сбора']

    df = pd.read_excel(r'C:\Users\Abdykarim.D\Downloads\magnumopt_2023-08-09.xlsx')

    df.columns = df.iloc[10]

    for row in range(1, collection_sheet.max_row + 1):
        try:
            new_df = df[df['Дата валютир.'] == collection_sheet[f'B{row}'].value.strftime("%d.%m.%Y")]
        except:
            new_df = df[df['Дата валютир.'] == collection_sheet[f'B{row}'].value]

        filtered_df = new_df[new_df['Оплачено'] == collection_sheet[f'D{row}'].value]  # Отобрал только те записи, которые были произведены за D{row} день из файла сбора

        # print(filtered_df)
        for times in filtered_df['Дата/время транз.']:

            collection_date, homebank_date = collection_sheet[f'C{row}'].value, times

            time_diff = check_if_time_diff_less_than_1_min(collection_date, homebank_date)

            if time_diff <= 1:
                collection_sheet[f'E{row}'].value = 'да'

    collection_file.save(r'C:\Users\Abdykarim.D\Documents\Файл сбора2.xlsx')


def odines_part(days):

    opened_table_selector = {"title": "", "class_name": "", "control_type": "Table", "visible_only": True,
                             "enabled_only": True, "found_index": 0}
    filter_selector = {"title": "Установить отбор и сортировку списка...", "class_name": "",
                       "control_type": "Button",
                       "visible_only": True, "enabled_only": True, "found_index": 0}
    filter_whole_wnd_selector = {"title": "Отбор и сортировка", "class_name": "V8NewLocalFrameBaseWnd",
                                 "control_type": "Window", "visible_only": True, "enabled_only": True,
                                 "found_index": 0}

    app = Odines()
    app.run()

    app.navigate("Банк и касса", "Отчет банка по операциям эквайринга", maximize_innder=True)

    table_element = app.find_element(opened_table_selector)

    app.find_element(filter_selector).click()

    app.parent_switch(filter_whole_wnd_selector, resize=True)

    time.sleep(1)

    app.find_element({"title": "Пометка удаления", "class_name": "", "control_type": "CheckBox",
                      "visible_only": True, "enabled_only": True, "found_index": 0}).click()
    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                      "visible_only": True, "enabled_only": True, "found_index": 3}).type_keys('Нет', app.keys.TAB,
                                                                                               protect_first=True, clear=True,
                                                                                               click=True)

    app.find_element({"title": "Организация", "class_name": "", "control_type": "CheckBox",
                      "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                      "visible_only": True, "enabled_only": True, "found_index": 7}).type_keys('ТОО "Magnum Cash&Carry"', app.keys.TAB,
                                                                                               protect_first=True, clear=True,
                                                                                               click=True)

    app.find_element({"title": "Контрагент", "class_name": "", "control_type": "CheckBox",
                      "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    app.find_element({"title": "", "class_name": "", "control_type": "Edit",
                      "visible_only": True, "enabled_only": True, "found_index": 37}).type_keys('Частное лицо- ОПТ', app.keys.TAB,
                                                                                                protect_first=True, clear=True,
                                                                                                click=True)

    app.find_element({"title": "OK", "class_name": "", "control_type": "Button",
                      "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    app.parent_back(1)

    els = app.find_elements({"title_re": ".* Дата", "class_name": "", "control_type": "Custom",
                             "visible_only": True, "enabled_only": True}, timeout=3)

    all_days = []

    for i in els:

        clipboard_set("")
        i.type_keys("^c", click=True, clear=False)

        get_report_date = clipboard_get()
        get_report_date = str(get_report_date).strip()[:10]
        print(get_report_date)

        if get_report_date in days:

            transaction_dict = dict()

            i.click(double=True)

            sleep(3)

            print()

            app.parent_switch({"title": "", "class_name": "", "control_type": "Pane",
                               "visible_only": True, "enabled_only": True, "found_index": 29}, resize=True, set_focus=True, maximize=True)
            print()
            app.find_element({"title": "Развернуть", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True}, timeout=3).click()

            try:
                transactions = app.find_elements({"title_re": ".* Дата транзакции$", "class_name": "", "control_type": "Custom",
                                                  "visible_only": True, "enabled_only": True}, timeout=10)

                print(transactions)

                print(len(transactions))

            except:
                pass

            summs = app.find_elements({"title_re": ".* Сумма$", "class_name": "", "control_type": "Custom",
                                       "visible_only": True, "enabled_only": True}, timeout=5)

            print(summs)
            print(len(summs))

            for ind, transaction in enumerate(transactions):

                print('-------------------------------------------')
                clipboard_set("")
                transaction.type_keys("^c", click=True, clear=False)
                transaction.type_keys(app.keys.DOWN, click=True, clear=False)

                transaction_date = clipboard_get()
                transaction_date = str(transaction_date).strip()
                print(f'Transaction {transaction}: {transaction_date}')

                clipboard_set("")
                print('Clicking on', ind, summs[ind])
                summs[ind].type_keys("^c", click=True, clear=False)

                summ = clipboard_get()
                summ = round(float(str(summ).replace(' ', '').replace(',', '.').replace(' ', '')))
                print('Sum:', summ)
                print('-------------------------------------------')

                transaction_dict.update({transaction_date: summ})

            app.find_element({"title": "Закрыть", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True}).click()
            print('Finished')
            # exit()
            app.parent_back(1)

            all_days.append(transaction_dict)
    app.quit()

    print(all_days)


def odines_check_with_collection():

    all_days = [{'09.08.2023 10:12:42': 587520, '09.08.2023 10:40:22': 2499680, '09.08.2023 10:42:49': 875840, '09.08.2023 10:46:22': 2499680, '09.08.2023 11:47:15': 504000, '09.08.2023 12:52:40': 201960, '09.08.2023 13:49:30': 2499680, '09.08.2023 13:51:27': 2200480, '09.08.2023 14:17:50': 302080, '09.08.2023 15:43:12': 5572800, '09.08.2023 19:20:37': 2427456, '09.08.2023 19:52:35': 2052060}]

    collection_file = load_workbook(r'C:\Users\Abdykarim.D\Documents\Файл сбора2.xlsx')

    collection_sheet = collection_file['Файл сбора']

    for row in range(2, collection_sheet.max_row + 1):
        collection_sheet[f'F{row}'].value = 'нет'
        for day_ in all_days:
            for single_day in day_:
                # single_day_ = None
                # try:
                #     single_day_ = datetime.datetime.strptime(single_day, '%d.%m.%Y %H:%M:%S')
                # except:
                #     pass

                time_diff = check_if_time_diff_less_than_1_min(collection_sheet[f'C{row}'].value, single_day)
                if time_diff <= 1 and abs(day_.get(single_day) - round(collection_sheet[f'D{row}'].value)) <= 1:
                    print('--------------------------------------------------------------------------')
                    print(single_day, collection_sheet[f'C{row}'].value, day_.get(single_day), collection_sheet[f'D{row}'].value, time_diff, sep=' | ')
                    collection_sheet[f'F{row}'].value = 'да'

    collection_file.save(r'C:\Users\Abdykarim.D\Documents\Файл сбора2.xlsx')
    print('--------------------------------------------------------------------------')


def sign_ecp_kt(ecp):

    app = App('')

    el = {"title": "Открыть файл", "class_name": "SunAwtDialog", "control_type": "Window",
          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": None}

    if app.wait_element(el, timeout=30):

        keyboard.send_keys(ecp.replace('(', '{(}').replace(')', '{)}'), pause=0.01, with_spaces=True)
        sleep(0.05)
        keyboard.send_keys('{ENTER}')

        print('Finished ECP')

        app = None

        return 'signed'

    else:
        print('Quit mazafaka')
        app = None
        return 'broke'


def sign_ecp(ecp):

    app = App('')

    el = {"title": "Открыть файл", "class_name": "SunAwtDialog", "control_type": "Window",
          "visible_only": True, "enabled_only": True, "found_index": 0, "parent": None}

    if app.wait_element(el, timeout=30):

        keyboard.send_keys(ecp.replace('(', '{(}').replace(')', '{)}'), pause=0.01, with_spaces=True)
        sleep(0.05)
        keyboard.send_keys('{ENTER}')

        if app.wait_element({"title_re": "Формирование ЭЦП.*", "class_name": "SunAwtDialog", "control_type": "Window",
                             "visible_only": True, "enabled_only": True, "found_index": 0, "parent": None}, timeout=30):
            app.find_element({"title_re": "Формирование ЭЦП.*", "class_name": "SunAwtDialog", "control_type": "Window",
                              "visible_only": True, "enabled_only": True, "found_index": 0, "parent": None}).type_keys('Aa123456')
            sleep(2)
            keyboard.send_keys('{ENTER}')
            sleep(3)
            keyboard.send_keys('{ENTER}')
            app = None
            return 'signed'
        else:
            print('Quit mazafaka1')
            app = None
            return 'broke'
    else:
        print('Quit mazafaka')
        app = None
        return 'broke'


def ofd_distributor():

    collection_file = load_workbook(r'C:\Users\Abdykarim.D\Documents\Файл сбора2.xlsx')

    collection_sheet = collection_file['Файл сбора']

    mapping_file = pd.read_excel(mapping_path)

    for row in range(2, collection_sheet.max_row + 1):

        seacrh_date = collection_sheet[f'B{row}'].value
        collection_sheet[f'G{row}'].value = 'нет'
        short_name = mapping_file[mapping_file['Наименование в Спруте'] == collection_sheet[f'A{row}'].value]['Филиал'].iloc[0]
        ecp_path = mapping_file[mapping_file['Наименование в Спруте'] == collection_sheet[f'A{row}'].value]['Площадка в Спруте'].iloc[0]

        ecp_auth = ''
        ecp_sign = ''

        for file in os.listdir(os.path.join(ecp_paths, ecp_path)):
            if 'AUTH' in str(file):
                ecp_auth = os.path.join(ecp_paths, ecp_path, file)
            if 'GOST' in str(file):
                ecp_sign = os.path.join(ecp_paths, ecp_path, file)
        # print(ecp_sign, ecp_auth)
        print(short_name)
        if mapping_file[mapping_file['Наименование в Спруте'] == collection_sheet[f'A{row}'].value]['ОФД'].iloc[0] == 'Казахтелеком':
            print('kotaktelekom')
            collection_sheet[f'G{row}'].value = 'нет'
            try:
                open_oofd_kotaktelekom(seacrh_date, collection_sheet, row, ecp_auth, ecp_sign)
            except:
                # error exception
                pass
        elif mapping_file[mapping_file['Наименование в Спруте'] == collection_sheet[f'A{row}'].value]['ОФД'].iloc[0] == 'Транстелеком':
            print('trans')
            collection_sheet[f'G{row}'].value = 'нет'
            try:
                open_oofd_trans(seacrh_date, collection_sheet, row, ecp_auth, ecp_sign)
            except:
                # error exception
                pass

    collection_file.save(r'C:\Users\Abdykarim.D\Documents\Файл сбора2.xlsx')


def open_oofd_trans(seacrh_date, collection_sheet, row, ecp_auth, ecp_sign):

    web = Web()

    web.run()
    web.get('https://ofd1.kz/login')

    web.find_element('//*[@id="login_by_cert_btn"]').click()
    sign_ecp(ecp_auth)

    # if web.wait_element('//*[@id="close_i_modal"]/img', timeout=10):
    #     web.find_element('//*[@id="close_i_modal"]/img').click()
    sleep(5)
    web.get('https://ofd1.kz/cash_register?status_type=registered')

    web.find_element('//*[@id="sample"]').type_keys(collection_sheet[f'K{row}'].value)

    web.find_element('//*[@id="sign_btn"]').click()

    web.find_element('//*[@id="serach_results"]//a').click()

    web.find_element('//*[@id="shift_list_button"]').click()

    day_ = int(seacrh_date.split('.')[0])
    month_ = int(seacrh_date.split('.')[1])
    year_ = int(seacrh_date.split('.')[2])

    seacrh_date = datetime.datetime(year_, month_, day_)

    web.set_elements_value(xpath='//*[@id="start_date"]', value=seacrh_date.strftime('%Y-%m-%d'))
    web.set_elements_value(xpath='//*[@id="end_date"]', value=(seacrh_date + datetime.timedelta(days=1)).strftime('%Y-%m-%d'))
    print(seacrh_date)
    # web.find_element('//*[@id="shift_list_button_list"]').click()
    web.execute_script_click_xpath_selector('//*[@id="shift_list_button_list"]')

    sleep(1.5)

    if web.wait_element('//*[@id="shifts-container"]/tr/td[5]', timeout=5):

        dates = web.find_elements('//*[@id="shifts-container"]/tr/td[5]/preceding-sibling::td[3]')
        summs = web.find_elements('//*[@id="shifts-container"]/tr/td[5]')

        for ind in range(len(dates)):
            print(dates[ind].get_attr('text'))
            print(summs[ind].get_attr('text'))
            summ_ = round(float(summs[ind].get_attr('text').replace(' ', '')))
            if summ_ == int(collection_sheet[f'D{row}'].value) \
               and 1 >= check_if_time_diff_less_than_1_min(collection_sheet[f'C{row}'].value, datetime.datetime.strptime(dates[ind].get_attr('text'), '%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y %H:%M:%S')):
                collection_sheet[f'G{row}'].value = 'да'
            print('----------------------------------------------')
        sleep(1)


def open_oofd_kotaktelekom(seacrh_date, collection_sheet, row, ecp_auth, ecp_sign):

    web = Web()

    web.run()
    web.get('https://org.oofd.kz/#/landing/eds-login')

    if web.wait_element("//button[contains(text(), 'kz')]", timeout=10):
        web.find_element("//button[contains(text(), 'kz')]").click()
        web.execute_script_click_xpath_selector("//div[contains(text(), 'RU')]")

    web.find_element("//button[contains(text(), 'Войти с ЭЦП')]").click()

    web.find_element('//*[@id="storage-password"]').type_keys('Aa123456')
    web.execute_script_click_xpath_selector('//*[@id="storage-type"]/div/div[2]/div/p[2]/span')
    sleep(2)

    sign_ecp_kt(ecp_auth)

    sleep(3)
    if web.wait_element("//button[contains(text(), 'Проверить')]", timeout=5):
        web.find_element("//button[contains(text(), 'Проверить')]").click()
    print('Kek0')
    with suppress(Exception):
        app = App('')

        app.find_element({"title": "Ввод пароля", "class_name": "SunAwtDialog", "control_type": "Window",
                          "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=15).type_keys('Aa123456', app.keys.ENTER)

        sleep(1)

        app.find_element({"title": "Ввод пароля", "class_name": "SunAwtDialog", "control_type": "Window",
                          "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=10).type_keys('Aa123456', app.keys.ENTER)
        app = None
    print('Kek1')
    web.find_element("//input[contains(@placeholder, 'Магазин, касса')]").type_keys(str(collection_sheet[f'K{row}'].value).replace(' ', ''), web.keys.ENTER)  # Filling serial number
    print('Kek2')
    web.find_element("(//a[@class='kkm'])[1]").click()  # Find & click on the first element
    # sleep(100)
    print(seacrh_date, seacrh_date.split('.'))
    if True:
        year = seacrh_date.split('.')[2]
        month = seacrh_date.split('.')[1]
        day = seacrh_date.split('.')[0]
        web.set_elements_innerhtml_or_value('//*[@id="mat-input-0"]', element_type='value', date=f'{year}-{month}-{day}T00:00:00', value=f'{int(day)}.{int(month)}.{year}')
        web.set_elements_innerhtml_or_value('//*[@id="mat-input-1"]', element_type='value', date=f'{year}-{month}-{day}T23:59:59', value=f'{int(day)}.{int(month)}.{year}')

        web.get(f'https://org.oofd.kz/#/main/kkms/757458?startDate={year}-{month}-{day}T00:00:00&endDate={year}-{month}-{day}T23:59:59&page=1')

        transactions = web.find_elements("//div[@class='transaction-wrapper ng-star-inserted']", timeout=40)
        times = web.find_elements("//div[@class='transaction-wrapper ng-star-inserted']/tax-transaction/div/span/span", timeout=1)
        summs = web.find_elements("//div[@class='transaction-wrapper ng-star-inserted']//span[@class='transaction__sum ng-star-inserted']", timeout=1)

        for ind in range(len(transactions)):

            time_ = " " + times[ind].get_attr('text').split()[-1] + ":00"
            summ_ = round(float(summs[ind].get_attr('text').replace('₸', '').replace(' ', '').replace(',', '.')))
            print(seacrh_date + time_, summ_)
            print(check_if_time_diff_less_than_1_min(seacrh_date + time_, collection_sheet[f'C{row}'].value))
            sleep(10)

            if check_if_time_diff_less_than_1_min(seacrh_date + time_, collection_sheet[f'C{row}'].value) <= 1 and summ_ == int(collection_sheet[f'D{row}'].value):
                collection_sheet[f'G{row}'].value = 'да'

    # sleep(10000)
    print('-----------------------------------------------------------------------------')


if __name__ == '__main__':

    if True:

        sql_create_table()

        today = datetime.datetime.now().date()
        today = datetime.date(2023, 8, 4)
        cashbook_day = (today - datetime.timedelta(days=5)).strftime('%d.%m.%Y')

        days = []

        for i in range(7, 1, -1):
            day = (today - datetime.timedelta(days=i)).strftime('%d.%m.%Y')
            days.append(day)

        today = today.strftime('%d.%m.%Y')
        print(today, cashbook_day)
        days.append('11.08.2023')
        days = ['12.08.2023']
        print(days)

        # filepath = open_cashbook(today)
        # filepath = r'C:\Users\Abdykarim.D\Documents\Export_230908_124925.xlsx'
        # create_collection_file(filepath)
        #
        # homebank('mukhtarova@magnum.kz', 'Aa123456!')
        #
        check_homebank_and_collection()
        #
        odines_part(days)
        #
        odines_check_with_collection()

        ofd_distributor()

        # # open_oofd_kotaktelekom()

    # except Exception as error:
    #     print('GOVNO', error)
    #     sleep(2000)
