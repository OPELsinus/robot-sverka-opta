import datetime

# months = [
#     'января',
#     'февраля',
#     'марта',
#     'апреля',
#     'мая',
#     'июня',
#     'июля',
#     'августа',
#     'сентября',
#     'октября',
#     'ноября',
#     'декабря'
# ]
#
# today = datetime.date(2024, 1, 1)
#
# cashbook_day = (today - datetime.timedelta(days=5)).strftime('%d.%m.%Y')
# today = today.strftime('%d.%m.%Y')
# print(today, cashbook_day)
#
# day = int(cashbook_day.split('.')[0])
# month = int(cashbook_day.split('.')[1])
# year = int(cashbook_day.split('.')[2])
#
# print(f'{day} {months[month - 1]} {year} г.')
import os
import uuid
from math import ceil

import psycopg2

from config import db_host, db_port, db_name, db_user, db_pass


def a(file):

    import openpyxl

    all_days = []

    monthes = ['', 'январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь',
               'декабрь']
    template_path = "\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-sverka-beznala\\Шаблон.xlsx"
    main_excel_file = file # r'C:\Users\Abdykarim.D\Documents\ШФ4 Сверка по безналичной выручке 2023 год.xlsx'

    wb = openpyxl.load_workbook(main_excel_file, data_only=False)
    print(f"sheet names of {main_excel_file}")

    for int_process_month in range(9, 10): # * Месяцы, которые вносить в базу для отработки
        # int_process_month = 8
        int_process_year = datetime.datetime.today().year

        month_name_rus = monthes[int_process_month]
        needed_sheet_name = None
        for sheet_name in wb.sheetnames:
            if f"{month_name_rus}{int_process_year}" in sheet_name or f"{month_name_rus} {int_process_year}" in sheet_name or f"{month_name_rus}{str(int_process_year)[2:]}" in sheet_name:
                needed_sheet_name = sheet_name
                break
        max_days = datetime.date(int_process_year, int_process_month + 1, 1) - datetime.timedelta(days=1)

        for day in range(1, max_days.day):
            index_of_process_date: int = day + 2

            value = wb[needed_sheet_name].cell(index_of_process_date, 3).value

            if value is None:

                today = datetime.datetime.today()
                if (today - datetime.datetime(int_process_year, int_process_month, day)).days >= 0:
                    # print(f'{day}.0{int_process_month}')
                    if day < 10:
                        if int_process_month < 10:
                            all_days.append(f'0{day}.0{int_process_month}.{int_process_year}')
                        else:
                            all_days.append(f'0{day}.{int_process_month}.{int_process_year}')
                    else:
                        if int_process_month < 10:
                            all_days.append(f'{day}.0{int_process_month}.{int_process_year}')
                        else:
                            all_days.append(f'{day}.{int_process_month}.{int_process_year}')

        """ Безнал по Z - in column C
            Безнал по БД СПРУТ D
            Расхождение между Z и БД Спрут E
            Безнал по Выписке банка F
            Расхождение между ВБ и БД G
        """
    # wb.save(main_excel_file)
    wb.close()

    return all_days


def table_create():
    conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)

    table_create = '''CREATE TABLE IF NOT EXISTS ROBOT.ROBOT_SVERKA_BEZNALA_TEST (
    id text PRIMARY KEY,
    process_date text,
    branch_name text,
    odines_name text,
    sprut_name text,
    store_names text,
    main_excel_file text,
    status text,
    retry_count INTEGER,
    error_message text,
    comments text,
    execution_time text,
    finish_date text,
    date_created text,
    executor_name text) '''
    c = conn.cursor()
    c.execute(table_create)
    conn.commit()
    c.close()
    conn.close()


def read_mapping_excel_file(path):
    import pandas as pd
    df = pd.read_excel(path, sheet_name='Свод')

    return df


def define_executors():

    executors_name = ['10.70.2.12', '10.70.2.23', '10.70.2.11', '10.70.2.10', '10.70.2.3']
    executors = dict()

    branches = ['ШФ33', 'ШФ7', 'АФ8', 'ППФ4', 'АСФ9', 'АСФ6', 'АФ22', 'АФ36', 'ШФ25', 'ШФ10', 'АФ77', 'ШФ34', 'АСФ47', 'АСФ60', 'АФ31', 'АСФ74', 'АФ4', 'АФ30', 'АСФ39', 'АСФ71', 'АФ56', 'ШФ12', 'АФ60', 'ШФ4', 'АФ68', 'АФ82', 'АФ40', 'АФ17', 'АСФ10', 'АСФ24', 'ШФ8', 'АСФ32', 'АСФ69', 'АСФ31', 'АСФ14', 'АСФ35', 'АФ29', 'АФ63', 'АФ84', 'АСФ55', 'АСФ66', 'ШФ26', 'ШФ19', 'АСФ2', 'АФ71', 'АФ61', 'ШФ27', 'ШФ24', 'АСФ21', 'АСФ81', 'АСФ73', 'АФ46', 'АФ19', 'ППФ20', 'АФ12', 'АФ44', 'ППФ7', 'АСФ27', 'АСФ4', 'АФ2', 'АФ39', 'АСФ56', 'АФ80', 'ТФ1', 'АСФ45', 'АФ58', 'АФ50', 'ФКС2', 'АФ65', 'АСФ48', 'ППФ22', 'АФ6', 'АФ76', 'ТФ2', 'АСФ41', 'ППФ16', 'АСФ25', 'ТЗФ2', 'АФ70', 'ППФ2', 'АСФ57', 'АСФ67', 'АСФ1', 'АФ73', 'АФ38', 'АСФ63', 'ППФ9', 'АСФ61', 'АСФ16', 'ШФ6', 'АСФ34', 'АСФ65', 'АФ25', 'АСФ51', 'ППФ11', 'АСФ52', 'АФ7', 'АСФ36', 'АСФ28', 'КФ2', 'ППФ3', 'ТКФ1', 'КФ5', 'АСФ53', 'АФ42', 'АФ49', 'АФ51', 'КФ1', 'АФ59', 'АФ32', 'АФ26', 'ШФ23', 'АСФ20', 'АСФ58', 'АСФ75', 'АФ9', 'АСФ23', 'ШФ1', 'ШФ21',
                'АСФ80', 'АФ78', 'АСФ46', 'ППФ6', 'АФ20', 'ШФ32', 'ШФ30', 'КФ6', 'АФ48', 'АФ35', 'УКФ2', 'АФ64', 'ШФ35', 'АФ54', 'АФ14', 'АФ66', 'АФ72', 'ППФ18', 'АСФ7', 'АСФ64', 'ППФ1', 'ШФ22', 'ШФ28', 'АСФ72', 'АСФ42', 'АСФ59', 'АФ52', 'АФ62', 'АФ16', 'АФ41', 'УКФ3', 'АФ34', 'КЗФ1', 'ППФ19', 'АСФ29', 'АФ57', 'АФ24', 'АСФ54', 'АСФ5', 'АСФ3', 'АФ45', 'ШФ3', 'АСФ15', 'АСФ17', 'АСФ82', 'АСФ83', 'АФ43', 'АФ21', 'АСФ77', 'АСФ38', 'АФ3', 'АСФ12', 'АСФ70', 'АФ33', 'АСФ50', 'ТЗФ3', 'ППФ8', 'АСФ33', 'ЕКФ1', 'ТФ3', 'ППФ17', 'ППФ5', 'ППФ10', 'АСФ40', 'ШФ9', 'ШФ13', 'АФ69', 'АСФ26', 'АФ75', 'УКФ1', 'АФ15', 'АСФ68', 'АФ47', 'АСФ11', 'АСФ62', 'ШФ20', 'ШФ14', 'АФ23', 'ППФ21', 'ФКС1', 'АСФ13', 'АСФ8', 'ППФ15', 'АСФ18', 'АФ10', 'АФ11', 'АФ37', 'ШФ15', 'АФ53', 'АФ83', 'ШФ2', 'АСФ19', 'ТКФ2', 'АФ18', 'ППФ13', 'АФ67', 'АФ28', 'КФ7', 'ШФ17', 'ШФ5', 'ШФ29', 'АСФ30', 'ШФ18']

    l = ceil(len(branches) / 5)

    print(len(branches), l)

    for i in range(5):

        br = []

        for j in range(l):

            try:
                br.append(branches[0])
                branches.remove(branches[0])
            except:
                pass

        executors.update({executors_name[i]: br})

    return executors


def dispatcher():

    print("Dispatcher starts")

    table_create()

    executors = define_executors()

    # process_date = datetime.datetime.strptime("26.07.2023", "%d.%m.%Y")

    # str_process_date = process_date.strftime("%d.%m.%Y")

    main_directory_folder = "\\\\vault.magnum.local\\Common\\Stuff\\_06_Бухгалтерия\\Для робота\\Процесс безнала\\отчеты 2023"

    files = os.listdir(main_directory_folder)

    df = read_mapping_excel_file("\\\\172.16.8.87\\d\\.rpa\\.agent\\robot-sverka-beznala\\маппинг для сверки безнала.xlsx")

    list_of_lost_records = []
    tr_count = 0
    for i, row in df.iterrows():
        record_found = False
        file_name = None
        for file in files:
            name_no_spaces = str(file).replace(" ", "")
            search = f"{str(row[1]).replace(' ', '')}Сверкапобезналичной"
            # Mistake if КФ1 and ЕКФ1 exists

            if name_no_spaces.startswith(search):

                if "~$" in file:
                    continue
                file_name = file
                record_found = True

                break
        if record_found:
            # find_query = f"Select process_date from ROBOT.ROBOT_SVERKA_BEZNALA_TEST where process_date='{str_process_date}' AND branch_name='{row[1]}' AND odines_name= '{row[2]}' AND sprut_name= '{row[3]}'"
            conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)
            c = conn.cursor()
            # c.execute(find_query)
            # result = c.fetchone()
            try:
                result = a(os.path.join(main_directory_folder, file_name))
            except:
                print(f'BRANCH {file_name} IS CLOSED!!!')
                result = None
            str_now = datetime.datetime.now().strftime('%d.%m.%Y %H:%M:%S')
            if result is not None:
                for day in result:
                    for executor in executors:
                        if row[1] in executors.get(executor):
                            insert_q = f"Insert Into ROBOT.ROBOT_SVERKA_BEZNALA (id, process_date, branch_name, odines_name, sprut_name, store_names, main_excel_file, status, retry_count, date_created, executor_name) values ('{uuid.uuid4()}', '{day}','{row[1]}','{row[2]}', '{row[3]}','{row[4]}','{os.path.join(main_directory_folder, file_name)}', 'New', 0, '{str_now}', '{executor}')"
                            c.execute(insert_q)
                            conn.commit()
                c.close()
                conn.close()
                tr_count += 1
        else:
            list_of_lost_records.append(str(row[1]))
    print(f"Добавили {tr_count} в db")


if __name__ == '__main__':
    dispatcher()
