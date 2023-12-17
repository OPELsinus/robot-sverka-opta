import os
from contextlib import suppress
from time import sleep

from pandas.io.clipboard import clipboard_get
from pywinauto import keyboard

from config import logger
from core import Sprut
from tools import clipboard


def open_cashbook(today):

    sprut = Sprut("MAGNUM")
    sprut.run()

    try:

        sprut.open("Кассовая книга", switch=False)

        sprut.parent_switch({"title_re": ".Кассовая книга.", "class_name": "Tbo_cashbook_fm_main",
                             "control_type": "Window", "visible_only": True, "enabled_only": True, "found_index": 0}).set_focus()
        sprut.find_element({"title": "Приложение", "class_name": "", "control_type": "MenuBar",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).click(coords=(50, 17))

        sprut.find_element({"title": "", "class_name": "", "control_type": "MenuItem",
                            "visible_only": True, "enabled_only": True, "found_index": 3}).click()

        sprut.find_element({"title": "", "class_name": "", "control_type": "MenuItem",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).click()
        print()

        try:
            sprut.find_element({"title": "Последний использованный фильтр", "class_name": "TvmsToolGridQueryList", "control_type": "Pane",
                                "visible_only": True, "enabled_only": True, "found_index": 0}).click(coords=(290, 15))

            sprut.parent_switch({"title": "Выборка по запросу", "class_name": "Tvms_modifier_fm_builder", "control_type": "Window",
                                 "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=30).set_focus()

        except:

            sprut.find_element({"title": "Последний использованный фильтр", "class_name": "TvmsToolGridQueryList", "control_type": "Pane",
                                "visible_only": True, "enabled_only": True, "found_index": 0}).click(coords=(380, 17))

            sprut.parent_switch({"title": "Выборка по запросу", "class_name": "Tvms_modifier_fm_builder", "control_type": "Window",
                                 "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=30).set_focus()

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
            with suppress(Exception):
                sprut.find_element({"title": "", "class_name": "TvmsBitBtn", "control_type": "Button",
                                    "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1).click()
                break

        sprut.find_element({"title": "И", "class_name": "TvmsBitBtn", "control_type": "Button",
                            "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=30).click()

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
                            "visible_only": True, "enabled_only": True, "found_index": 1}).type_keys(today)

        sprut.find_element({"title_re": ".", "class_name": "TvmsComboBox", "control_type": "Pane",
                            "visible_only": True, "enabled_only": True, "found_index": 1}).click()
        sprut.find_element({"title_re": ".", "class_name": "TvmsComboBox", "control_type": "Pane",
                            "visible_only": True, "enabled_only": True, "found_index": 1}).click()
        # print('started waiting')
        # sleep(60)
        # print('finished waiting')
        # sleep(60)
        sprut.find_element({"title_re": ".", "class_name": "TvmsComboBox", "control_type": "Pane",
                            "visible_only": True, "enabled_only": True, "found_index": 1}).type_keys('^N', sprut.keys.ENTER)

        while True:

            with suppress(Exception):
                sprut.find_element({"title": "", "class_name": "TvmsBitBtn", "control_type": "Button",
                                    "visible_only": True, "enabled_only": True, "found_index": 1}, timeout=1).click()
                break

        while True:
            with suppress(Exception):
                sprut.find_element({"title": "Ввод", "class_name": "TvmsBitBtn", "control_type": "Button",
                                    "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=1).click()
                break

        sprut.parent_back(1)

        sprut.find_element({"title": "", "class_name": "TvmsDBToolGrid", "control_type": "Pane",
                            "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=360).click()

        # * Берёт бонусы

        keyboard.send_keys('^A')
        keyboard.send_keys('^{INSERT}')

        bonuses = []

        sprut.find_element({"title": "", "class_name": "TvmsDBToolGrid", "control_type": "Pane",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(sprut.keys.UP * 30)

        for i in range(len(clipboard_get().split(','))):

            sprut.find_element({"title": "Оплаты по чеку", "class_name": "", "control_type": "TabItem",
                                "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            sprut.find_element({"title": "", "class_name": "TvmsDBToolGrid", "control_type": "Pane",
                                "visible_only": True, "enabled_only": True, "found_index": 0}).click()

            sprut.find_element({"title": "", "class_name": "TvmsDBToolGrid", "control_type": "Pane",
                                "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(sprut.keys.UP * 5)
            sprut.find_element({"title": "", "class_name": "TvmsDBToolGrid", "control_type": "Pane",
                                "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(sprut.keys.LEFT * 15)
            bonuse = 0
            for rows in range(2):
                is_bonuse = False
                keyboard.send_keys('{LEFT}' * 10)
                for cols in range(5):

                    keyboard.send_keys('^{INSERT}')

                    val = clipboard.clipboard_get()

                    if 'бонус' in val:
                        is_bonuse = True

                    if is_bonuse and cols == 4:
                        bonuse = int(val)

                    keyboard.send_keys('{RIGHT}')
                keyboard.send_keys('{DOWN}')
            print('BONUS', bonuse)
            bonuses.append(bonuse)
            sprut.find_element({"title": "Чеки", "class_name": "", "control_type": "TabItem",
                                "visible_only": True, "enabled_only": True, "found_index": 0}).click()
            sprut.find_element({"title": "", "class_name": "TvmsDBToolGrid", "control_type": "Pane",
                                "visible_only": True, "enabled_only": True, "found_index": 0}).click()
            keyboard.send_keys('{DOWN}')
        # * ---

        sprut.find_element({"title": "", "class_name": "TvmsDBToolGrid", "control_type": "Pane",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys('^%E')

        sprut.parent_switch({"title": "Экспортировать данные", "class_name": "Tvms_fm_DBExportExt", "control_type": "Window",
                             "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=10)

        sprut.find_element({"title": "", "class_name": "TcxCustomDropDownInnerEdit", "control_type": "Edit",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).click()

        sprut.find_element({"title": "", "class_name": "TcxCustomDropDownInnerEdit", "control_type": "Edit",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).type_keys(sprut.keys.DOWN, sprut.keys.ENTER)

        file_path = sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit",
                                        "visible_only": True, "enabled_only": True, "found_index": 1}).get_text()
        logger.info(file_path)

        sprut.find_element({"title": "", "class_name": "TcxCustomInnerTextEdit", "control_type": "Edit", "visible_only": True, "enabled_only": True, "found_index": 0}).set_text('')

        right_pane = {"title": "", "class_name": "TvmsListBox", "control_type": "Pane", "visible_only": True, "enabled_only": True, "found_index": 0}

        sprut.find_element(right_pane).click()

        for i in range(8):
            try:
                sprut.find_element({"title": "Срочное проведения чека?", "class_name": "", "control_type": "ListItem",
                                    "visible_only": True, "enabled_only": True, "found_index": 0}, timeout=3).click()

                sprut.find_element({"title": "", "class_name": "", "control_type": "Button",
                                    "visible_only": True, "enabled_only": True, "found_index": 1}, timeout=1).click()
                break

            except:
                sprut.find_element(right_pane).type_keys(sprut.keys.PAGE_DOWN)

        sprut.find_element({"title": "Ввод", "class_name": "TvmsFooterButton", "control_type": "Button",
                            "visible_only": True, "enabled_only": True, "found_index": 0}).click()

        wait_loading(file_path)

        os.system('taskkill /im excel.exe /f')

        sprut.quit()

        return file_path, bonuses

    except Exception as error:

        logger.warning(f'SPRUT ERROR: {error}')

        sprut.quit()

        return ['', '']


def wait_loading(filepath):
    print('Started loading')
    while True:
        if os.path.isfile(filepath):
            print('downloaded')
            break
    print('Finished loading')
    sleep(3)
