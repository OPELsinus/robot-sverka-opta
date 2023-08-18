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

    for int_process_month in range(6, 9):
        # int_process_month = 8
        int_process_year = 2023

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
                if (today - datetime.datetime(int_process_year, int_process_month, day)).days > 4:
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


def dispatcher():
    print("Dispatcher starts")
    table_create()

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
                    insert_q = f"Insert Into ROBOT.ROBOT_SVERKA_BEZNALA_TEST (id, process_date, branch_name, odines_name, sprut_name, store_names, main_excel_file, status, retry_count, date_created) values ('{uuid.uuid4()}', '{day}','{row[1]}','{row[2]}', '{row[3]}','{row[4]}','{os.path.join(main_directory_folder, file_name)}', 'New', 0, '{str_now}' )"
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
