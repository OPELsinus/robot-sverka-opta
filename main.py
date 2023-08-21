import datetime
import os
import shutil
import time
from copy import copy

import pandas as pd
from time import sleep

import win32com.client as win32
import psycopg2 as psycopg2
from openpyxl import load_workbook
from pywinauto import keyboard

from config import download_path, robot_name, db_host, db_port, db_name, db_user, db_pass, tg_token, chat_id, logger
from core import Sprut, Odines
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


def wait_loading(filepath):
    print('Started loading')
    logger.info('Started loading')
    while True:
        if os.path.isfile(filepath):
            print('LOOOL NASHEL')
            break
    print('Finished loading')
    logger.info('Finished loading')
    sleep(3)


def homebank(email, password):
    web = Web()
    web.run()
    web.get('https://epay.homebank.kz/login')

    web.find_element('//*[@id="mp-content"]/section/main/div[2]/div/div/div[2]/form/div[1]/div/div/span/div/input').type_keys(email)
    web.find_element('//*[@id="mp-content"]/section/main/div[2]/div/div/div[2]/form/div[2]/div/div/span/div/span/input').type_keys(password)

    web.find_element('//*[@id="mp-content"]/section/main/div[2]/div/div/div[2]/form/div[3]/div/div/span/button').click()
    print()
    web.get('https://epay.homebank.kz/statements/payment')

    web.find_element("//span[contains(text(), '427693/14-EC27/07')]").click()

    web.find_element('//*[@id="mp-content"]/div/div/div/div/div[1]/div/div/div[1]/div/div/div/div[2]/button').click()

    web.find_element('//*[@id="period"]').click()

    web.find_element("//td[@title = '31 августа 2023 г.']").click()


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

    for i in els:
        # print(i)
        clipboard_set("")
        i.type_keys("^c", click=True, clear=False)

        get_report_date = clipboard_get()
        get_report_date = str(get_report_date).strip()[:10]
        print(get_report_date)

        if get_report_date in days:

            i.click(double=True)

            sleep(3)

            print()

            app.parent_switch({"title": "", "class_name": "", "control_type": "Pane",
                               "visible_only": True, "enabled_only": True, "found_index": 29}, resize=True, set_focus=True)

            try:
                transactions = app.find_elements({"title_re": ".* Дата транзакции$", "class_name": "", "control_type": "Custom",
                                                  "visible_only": True, "enabled_only": True}, timeout=10)
            except:
                pass

            summs = app.find_elements({"title_re": ".* Сумма$", "class_name": "", "control_type": "Custom",
                                       "visible_only": True, "enabled_only": True}, timeout=5)
            print(summs)
            for ind, transaction in enumerate(transactions):

                print('-------------------------------------------')
                clipboard_set("")
                transaction.type_keys("^c", click=True, clear=False)

                transaction_date = clipboard_get()
                transaction_date = str(transaction_date).strip()[:10]
                print(f'Transaction {i}: {transaction_date}')

                clipboard_set("")
                print('Clicking on', summs[ind])
                summs[ind].type_keys("^c", click=True, clear=False)

                summ = clipboard_get()
                summ = str(summ)
                print('Sum:', summ)
                print('-------------------------------------------')

            app.find_element({"title": "Закрыть", "class_name": "", "control_type": "Button",
                              "visible_only": True, "enabled_only": True}).click()
            print('Finished')
            # exit()
            app.parent_back(1)


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
        print(days)

        # open_cashbook(today)

        collection_file = load_workbook(r'C:\Users\Abdykarim.D\Documents\Файл сбора.xlsx')
        collection_sheet = collection_file['Файл сбора']

        df = pd.read_excel(r'C:\Users\Abdykarim.D\Documents\Export_230821_121856.xlsx')

        print(df.columns)
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
                print(col_key, col_name)
                previous_row = collection_sheet[last_row - 1]
                source_cell = collection_sheet.cell(row=last_row - 1, column=collection_sheet[col_key + '1'].column)
                new_cell = collection_sheet.cell(row=last_row, column=collection_sheet[col_key + '1'].column)

                new_cell._style = copy(source_cell._style)
                new_cell.font = copy(source_cell.font)
                new_cell.border = copy(source_cell.border)
                new_cell.alignment = copy(source_cell.alignment)

                cell = collection_sheet[f'{col_key}{last_row}']
                try:
                    cell.value = row[col_name]
                    cell.alignment = copy(source_cell.alignment)
                except:
                    print('kek')
                    cell.value = None

        columns = ['Компания', 'Дата чека', 'Дата и время чека', 'Сумма с НДС', 'Ерау', '1с', 'офд ', 'примечание', '', 'Номер чека', 'Серийный № фиск.регистратора', 'Клиент', 'Дата создания записи', 'Состояние розничного чека']
        # collection_file = collection_file[columns]
        print(columns, len(columns))
        collection_file.save(r'C:\Users\Abdykarim.D\Documents\Файл сбора1.xlsx')
        # homebank('mukhtarova@magnum.kz', 'Aa123456!')

        # odines_part()

        # odines_part(days)

    # except Exception as error:
    #     print('GOVNO', error)
    #     sleep(2000)
