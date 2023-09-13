import datetime

from config import logger


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
            logger.info("GOVNO", e)
            pass
        pass

    # logger.info(first_date, second_date)
    # logger.info((first_date - second_date).total_seconds() // 60)

    return abs((first_date - second_date).total_seconds() // 60)