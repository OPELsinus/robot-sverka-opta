import os
from contextlib import suppress
from time import sleep

from config import logger
from core import Sprut


def open_cashbook(today):

    sprut = Sprut("REPS")
    sprut.run()

    sprut.open("Кассовая книга", switch=False)

    sprut.parent_switch({"title_re": ".Кассовая книга.", "class_name": "Tbo_cashbook_fm_main",
                         "control_type": "Window", "visible_only": True, "enabled_only": True, "found_index": 0}).set_focus()
    sprut.find_element({"title": "Приложение", "class_name": "", "control_type": "MenuBar",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click(coords=(50, 17))

    sprut.find_element({"title": "", "class_name": "", "control_type": "MenuItem",
                        "visible_only": True, "enabled_only": True, "found_index": 3}).click()

    sprut.find_element({"title": "", "class_name": "", "control_type": "MenuItem",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click()

    sprut.find_element({"title": "Последний использованный фильтр", "class_name": "TvmsToolGridQueryList", "control_type": "Pane",
                        "visible_only": True, "enabled_only": True, "found_index": 0}).click(coords=(380, 17))

    sprut.parent_switch({"title": "Выборка по запросу", "class_name": "Tvms_modifier_fm_builder", "control_type": "Window",
                         "visible_only": True, "enabled_only": True, "found_index": 0}).set_focus()

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
                        "visible_only": True, "enabled_only": True, "found_index": 1}).type_keys(today)

    sprut.find_element({"title_re": ".", "class_name": "TvmsComboBox", "control_type": "Pane",
                        "visible_only": True, "enabled_only": True, "found_index": 1}).click()
    sprut.find_element({"title_re": ".", "class_name": "TvmsComboBox", "control_type": "Pane",
                        "visible_only": True, "enabled_only": True, "found_index": 1}).click()
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
    logger.info(file_path)

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
