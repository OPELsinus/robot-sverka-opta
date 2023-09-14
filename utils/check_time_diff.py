import datetime

from config import logger


def check_time_diff(first_date, second_date, mins):

    try:
        first_date = datetime.datetime.strptime(first_date, '%d.%m.%Y %H:%M:%S')
    except:
        ...

    try:
        second_date = datetime.datetime.strptime(second_date, '%d.%m.%Y %H:%M')
    except:
        try:
            second_date = datetime.datetime.strptime(second_date, '%d.%m.%Y %H:%M:%S')
        except Exception as e:
            logger.info("GOVNO", e)

    max_ = max(first_date, second_date)
    min_ = min(first_date, second_date)

    return True if abs((max_ - min_).total_seconds() // 60) <= mins else False
