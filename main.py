import datetime
import os
import shutil
import time

import pandas as pd
from time import sleep

import win32com.client as win32
import psycopg2 as psycopg2
from pywinauto import keyboard

from config import download_path, robot_name, db_host, db_port, db_name, db_user, db_pass, tg_token, chat_id
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


def open_cashbook(sprut, today):
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

    a = sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                            "visible_only": True, "enabled_only": True, "found_index": 1}).get_text()
    print(a)
    print()
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

    # sprut.find_element({"title": " ", "class_name": "TPanel", "control_type": "Pane",
    #                     "visible_only": True, "enabled_only": True, "found_index": 0}).click(coords=(20, 17))
    #
    # sprut.find_element({"title": "Журналы", "class_name": "", "control_type": "MenuItem",
    #                     "visible_only": True, "enabled_only": True, "found_index": 0}).click()
    #
    # sprut.find_element({"title": "", "class_name": "", "control_type": "MenuItem",
    #                     "visible_only": True, "enabled_only": True, "found_index": 2}).click()
    #
    # sprut.find_element({"title": "", "class_name": "", "control_type": "MenuItem",
    #                     "visible_only": True, "enabled_only": True, "found_index": 1}).click()
    #
    # sprut.find_element({"title": "Закрыть", "class_name": "", "control_type": "Button",
    #                     "visible_only": True, "enabled_only": True, "found_index": 0}).click()
    #
    # sprut.parent_back(1)


def create_z_reports(sprut, branches, start_date, end_date):
    sprut.open("Отчеты")

    keyboard.send_keys("{F5}")

    sprut.parent_switch({"title": "Поиск", "class_name": "Tvms_search_fm_builder", "control_type": "Window",
                         "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.find_element({"title": "Название отчета", "class_name": "TvmsComboBox", "control_type": "Pane",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    keyboard.send_keys("{UP}" * 4)
    keyboard.send_keys("{ENTER}")
    # sprut.find_element({"title": "", "class_name": "", "control_type": "ListItem",
    #                     "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys('3303')

    keyboard.send_keys("{ENTER}")

    sprut.find_element({"title": "Перейти", "class_name": "TvmsBitBtn", "control_type": "Button",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    # ? ---------------------------------------------------------
    # sprut.get_pane(1).type_keys(sprut.Keys.F9)

    sprut.parent_back(1)

    for branch in branches[::]:
        print('Started', branch)

        sprut.find_element({"title": "", "class_name": "", "control_type": "SplitButton",
                            "visible_only": True, "enabled_only": True, "found_index": 4}).click()

        sprut.parent_switch({"title": "N100912-Сверка Z отчётов и оборота Спрут", "class_name": "TfrmParams", "control_type": "Window",
                             "visible_only": True, "enabled_only": True, "found_index": 0})

        sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                            "visible_only": True, "enabled_only": True, "found_index": 1}).click(double=True)
        keyboard.send_keys("{BACKSPACE}" * 20)
        sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                            "visible_only": True, "enabled_only": True, "found_index": 1}).type_keys(start_date)

        sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).click()
        sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).click(double=True)
        keyboard.send_keys("{BACKSPACE}" * 20)
        sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(end_date)

        # ? Search for 1 branch
        sprut.find_element({"title": "", "class_name": "TcxCustomDropDownInnerEdit", "control_type": "Edit",
                            "visible_only": True, "enabled_only": True, "found_index": 1}).click()
        sprut.find_element({"title": "", "class_name": "TcxCustomDropDownInnerEdit", "control_type": "Edit",
                            "visible_only": True, "enabled_only": True, "found_index": 1}).click(double=True)

        if sprut.wait_element({"title": "Отчеты", "class_name": "#32770", "control_type": "Window",
                               "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=8):
            sprut.find_element({"title": "ОК", "class_name": "Button", "control_type": "Button", "visible_only": True, "enabled_only": True, "found_index": 0}).click()
        sprut.find_element({"title": "", "class_name": "TcxCustomDropDownInnerEdit", "control_type": "Edit",
                            "visible_only": True, "enabled_only": True, "found_index": 1}).type_keys("{F5}")

        sprut.parent_switch({"title": "Поиск", "class_name": "Tvms_search_fm_builder", "control_type": "Window",
                             "visible_only": True, "enabled_only": True, "found_index": 0}).set_focus()

        try:
            sprut.find_element({"title": "", "class_name": "", "control_type": "Button",
                                "visible_only": True, "enabled_only": True, "found_index": 1}, timeout=10).click()
        except:
            pass

        sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).click()

        sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys('^N')

        # branch = 'Алматинский филиал №1 ТОО "Magnum Cash&Carry"'

        sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(f'%{branch}%', sprut.keys.ENTER, protect_first=True)

        sprut.find_element({"title": "", "class_name": "", "control_type": "Button",
                            "visible_only": True, "enabled_only": True, "found_index": 1}).click()

        sprut.find_element({"title": "Выбрать", "class_name": "TvmsBitBtn", "control_type": "Button",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).click()
        sprut.parent_back(1)

        sprut.find_element({"title": "Ввод", "class_name": "TvmsFooterButton", "control_type": "Button",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).click()

        wait_loading(branch)

        sprut.parent_back(1).set_focus()


def wait_loading(branch):
    branch = branch.replace('.', '').replace('"', '')
    found = False
    while True:
        for file in os.listdir(download_path):
            sleep(.1)
            creation_time = os.path.getctime(os.path.join(download_path, file))
            current_time = datetime.datetime.now().timestamp()
            time_difference = current_time - creation_time
            days_since_creation = time_difference / (60 * 60 * 24)

            if int(days_since_creation) <= 1 and file[0] != '$' and '.' in file and 'xl' in file:
                print(file)
                type = '.' + file.split('.')[1]
                shutil.move(os.path.join(download_path, file), os.path.join(os.path.join(download_path, 'reports'), branch + type))
                found = True
                break
        if found:
            break


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

            app.parent_switch({"title": "", "class_name": "", "control_type": "Pane",
                               "visible_only": True, "enabled_only": True, "found_index": 34}, resize=True)

            transactions = app.find_elements({"title_re": ".* Дата транзакции", "class_name": "", "control_type": "Custom",
                                              "visible_only": True, "enabled_only": True}, timeout=5)

            summs = app.find_elements({"title_re": ".* Сумма", "class_name": "", "control_type": "Custom",
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
                              "visible_only": True, "enabled_only": True, "found_index": 5}).click()
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

        # sprut = Sprut("REPS")
        # sprut.run()

        # open_cashbook(sprut, today)

        # homebank('mukhtarova@magnum.kz', 'Aa123456!')

        # odines_part()

        odines_part(days)

    # except Exception as error:
    #     print('GOVNO', error)
    #     sleep(2000)
