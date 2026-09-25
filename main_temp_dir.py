import os
import pandas as pd
import warnings
warnings.filterwarnings('ignore') # игнорируем предупреждения
# блок импортов для обновления сводных
import pythoncom
pythoncom.CoInitializeEx(0)
import win32com.client
import sqlalchemy
import numpy as np
import copy

# блок импорта отправки почты
import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

import socket
import logging



def dir_link():
    """возвращает абсолютный путь
    """
    import os
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        return script_dir
    except:
        script_dir_2 = os.getcwd()
        return script_dir_2
DIR = dir_link()

from functools import wraps
import time
# декоратор для times-повторного выполнения функции при неудачном выполнении 
def retry(times, sec_):
    """_summary_

    Args:
        times (_type_): попыток
        sec_ (_type_): секунд между попытками
    """
    def wrapper_fn(f):
        @wraps(f)
        def new_wrapper(*args,**kwargs):
            for i in range(times):
                try:
                    print ('---ПОПЫТКА ЧТЕНИЯ ФАЙЛА ---- %s' % (i + 1))
                    return f(*args,**kwargs)
                except Exception as e:
                    error = e
                    print(time.sleep(sec_))
            raise error
        return new_wrapper
    return wrapper_fn


# === НАСТРОЙКА ЛОГИРОВАНИЯ С ПЕРЕЗАПИСЬЮ ===
LOG_FILE = os.path.join(DIR, "log_main_temp_dir.log")

# Удаляем старый лог, если он существует
if os.path.exists(LOG_FILE):
    try:
        os.remove(LOG_FILE)
        print(f" Старый лог удалён: {LOG_FILE}")
    except PermissionError:
        print(f"Не удалось удалить старый лог: возможно, файл открыт в другом приложении")
    except Exception as e:
        print(f"Ошибка при удалении лога: {e}")

# Теперь настраиваем логирование — файл будет создан заново
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),  # Создастся как новый
        logging.StreamHandler()  # вывод в консоль
    ]
)

logger = logging.getLogger(__name__)
logger.info("Запуск скрипта log_main_temp_dir - дополнения темпов")

@retry(10, 5)
def links_main(name_file, key):
    """функция для работы с путями, ссылки, вводные данные хранятся в блокноте

    Args:
        name_file (_type_): имя файла
        key (_type_): имя ключа

    Returns:
        _type_: _description_
    """
    try:
        file = pd.read_csv(name_file, sep=';')
        result = list(file[file['ключ']==key]['значение'])[0]
        return result
    except Exception as ex_:
        logger.error(f'ошибка функции {links_main.__name__} не удалось считать файл {name_file} или данные в нем {key} ошибка {ex_}')

#links_main('file_links.txt', 'read_file_main')

def my_pass(link):
    """функция считывания пароля

    Returns:
        _type_: _description_
    """
    
    try:
        with open(link, 'r') as actual_pass:
            return actual_pass.read()
        
    except Exception as ex_:
        logger.error(f'ошибка функции {my_pass.__name__} {ex_}')


def get_data_user_param(socket_name=True):
    """возврат данных пользователя по соккету (имени ПК) для отправки почты в формате кортежа
    name_pc server port username send_from password

    Args:
        socket_name (bool, optional): _description_. Defaults to True.

    Returns:
        _type_: _description_
    """
    try:
        df = pd.read_excel(links_main(fr"{DIR}\main_links.txt", "pass"),sheet_name='PC')
        if len(df) > 0:
            logger.info(f"датафрейм обнаружен")
            if socket_name == True: 
                socket_name = socket.gethostname()
                logger.info(f'ищем данные по сокету {socket_name}')
                records_dict = df.to_dict('records')
                index_result = [index for index, dct in enumerate(records_dict) if dct['name_pc'] == socket_name]
                if len(index_result)>0:
                    try:
                        index_data=index_result[0]
                        result_all = records_dict[index_data]
                        name_pc = result_all['name_pc']
                        server = result_all['server']
                        port = result_all['port']
                        username = result_all['username']
                        send_from = result_all['send_from']
                        password = result_all['password']
                        logger.info(f'данные получены и возвращены кортежем в следующем порядке : name_pc server port username send_from password')
                        return name_pc, server, port, username, send_from, password
                    except Exception as e:
                        logger.error(f'Ошибка функции {get_data_user_param.__name__} в блоке получения данных пользователя по имени пк {socket_name} ошибка {e}')
            else :
                logger.info(f"Поиск данных с учетом соккета не включен, будет возвращен весь фрейм с данными")
                return df
        else: 
            logger.info(f"датафрейм не найден")
            return None
    except Exception as e:
        logger.error(f'Ошибка функции {get_data_user_param.__name__} ошибка {e}')



get_data_user_param()

def test_connect_SQL_driver():
    """функция проверки доступных драйверов для работы с SQL
    от этого зависит какой драйвер будем использовать
    на серверном пк s-kao3 новый ODBC Driver 18 for SQL Server
    на рабочих пк ODBC Driver 13 for SQL Server
    так же на серверном варинате в настройке нужно отключение 
    проверки сертификатов SSL - это описывается в самих функциях подключения к SQL
    """
    try:
        import pyodbc
        wite_list_driver = ['ODBC Driver 13 for SQL Server', 'ODBC Driver 18 for SQL Server']
        logger.info(f'🔄 Проверка достутных драйверов')

        list_driver = []
        for driver in pyodbc.drivers():
            if 'SQL' in driver:
                list_driver.append(driver)

        logger.info(f'🛠️ Доступные драйвера SQL: {list_driver}')

        if len(list_driver)>0:
            found_drivers = []
            for i in wite_list_driver:
                if i in list_driver:
                    found_drivers.append(i)
            if len(found_drivers)>0:
                logger.info(f'✅ Найден драйвер: {found_drivers[0]} соответсвующий белому списку: {wite_list_driver}')
                return found_drivers[0]
            else:
                logger.error(f'⚠️ Нет драйверов соответсвующих белому списку: {wite_list_driver}')
        else:
            logger.error(f'⚠️ Доступные драйвера SQL не обнаружены')
    except Exception as e:
        logger.error(f'ошибка функции {test_connect_SQL_driver.__name__} {e}')

test_connect_SQL_driver()


def connect_bd_SQL(server, database, username, password):
    """функция подключения к бд SQL
    Returns:
        _type_: _description_
    """
    # import sqlalchemy
    # Замените значения на ваши
    # если серит ошибками - проверить данные по столбцам ибо в наценке оборотки было inf и из-за этого серило ошибки
    server = server
    database = database
    username = username
    password = password
    driver = test_connect_SQL_driver() # 'ODBC Driver 13 for SQL Server' или 'ODBC Driver 18 for SQL Server' # моддет работать просто 'SQL Server'

    try:
        # connection_string = f'mssql+pyodbc://{username}:{password}@{server}/{database}?driver={driver}' # или f'mssql://{username}:{password}@{server}/{database}?driver={driver}'
        # исправленный вариант
        connection_string = (
            f'mssql+pyodbc://{username}:{password}@{server}/{database}'
            f'?driver={driver}&Encrypt=no'  # эта строка отключает проверку SSL сертификата
        )
        engine = sqlalchemy.create_engine(
                    connection_string,
                    echo=False, 
                    pool_pre_ping=True,
                    fast_executemany=True) 
        # fast_executemany=True - можно удалить эту строчку / Она оптимизирует процедуру массовых вставок, 
        # значительно сокращая количество запросов к базе данных, ускоряя запись
        logger.info(f'соединение с bd SQL - установлено')
        return engine
    except Exception as ex_:
        logger.error(f'Проблемы с подключением к SQL')
        logger.error(f'ошибка функции {connect_bd_SQL.__name__} ошибка {ex_}')


def exception_column_SQL(df, server, database, username, password):
    """проверяет наличие ошибок в df при записи в SQL
    перебирает каждый столбец и записывает в тестовую таблицу БД SQL
    при возникновении ошибки в записи - возвращает имя столбца и ошибку
    чаще такие ошибки вызван ытипом данных inf 
    Args:
        df (_type_): _description_
    """
    import sqlalchemy
    server = server
    database = database
    username = username
    password = password
    driver = test_connect_SQL_driver() # driver = 'ODBC Driver 13 for SQL Server' or # driver = 'ODBC Driver 18 for SQL Server'

    # connection_string = f'mssql+pyodbc://{username}:{password}@{server}/{database}?driver={driver}' # или f'mssql://{username}:{password}@{server}/{database}?driver={driver}'
    # исправленный вариант
    connection_string = (
        f'mssql+pyodbc://{username}:{password}@{server}/{database}'
        f'?driver={driver}&Encrypt=no'  # эта строка отключает проверку SSL сертификата
    )
    engine = sqlalchemy.create_engine(
                connection_string,
                echo=False, 
                pool_pre_ping=True,
                fast_executemany=True) # fast_executemany=True - можно удалить эту строчку / Она оптимизирует процедуру массовых вставок, значительно сокращая количество запросов к базе данных

    # перебираем каждую колонку df и пытаемся записать
    for i in df.columns:
        try:
            df[[i]].to_sql('df_ttt', con=engine, if_exists='replace', index=False)
            logger.info(f'✅ столбец {i} - ок')
        except Exception as ex_:
            logger.error(f'❌ {exception_column_SQL.__name__} ОШИБКА в столбце {i}  возможно порблемы  стипом данных {ex_}')


def yesterday(days:int=1):
    """возвращает дату на вчера - по цморлчанию минус 1 день
    Args:
        days (int, optional): на сколько дней назад откатываемся по дате. Defaults to 1.
    Returns:
        _type_: _description_
    """
    try:
        from datetime import datetime, timedelta
        date = datetime.now()
        new_date = date - timedelta(days=days)# вычитание одного дня
        return new_date
    except Exception as ex_:
        logger.error(f'ошибка функции {yesterday.__name__}  {ex_}')

# yesterday().strftime("%Y-%m-%d")

# отправка email в том числе со скрытыми получателями

def send_mail(send_to:list, 
              send_cc:list, 
              send_bcc:list, 
              topic_text:str, 
              body_text:str, 
              file_link:str, 
              file_name:str, 
              SEND_FROM:str, 
              SERVER:str, 
              PORT:int, 
              USER_NAME:str, 
              PASSWORD:str):
    """рассылка почты с вложением, пользователям в т.ч. добавление в копию и скрытую копию 

    Args:
        send_to (list):  список адресов для рассылки
        send_cc (list):  список адресов для рассылки - копии (можно не заполнять - проставить пустой список)
        send_bcc (list): список адресов для рассылки - скриыте копии (можно не заполнять - проставить пустой список)

        topic_text(str): текст темы писма
        body_text(str):  текст тела писма

        file_link(str):  ссылка на файл / если файла нет то оставить пустым file_name(str) = ''
        file_name(str):  имя файла в данном варианте нужно указывать с расширением 'BAIC_MSK.xlsx' (имя должно быть на латинице иначе придет в кодировке bin)
                         если файла нет то оставить пустым file_name(str) = ''

        SEND_FROM (str):  email пользователя от кого будет отправлено сообщение
        SERVER (str):     имя сервера
        PORT (int):       порт
        USER_NAME (str):  имя пользователя в сети
        PASSWORD (str):   пароль пользователя (учетной записи в сети)

        Пример:
        send_mail(['skrqqqo@siml-auto.ru'],             # send_to
          [],                                           # send_cc
          ['krutkosergey11111@yandex.ru'],              # send_bcc
          f'Привет привет',                             # topic_text
          f'Здравствуйте \nВо вложении файлик',         # body_text
          "//local/data/BAIC_MSK.xlsx",                 # file_link
          'BAIC_MSK.xlsx',                              # file_name
          'skrqqqo@siml-auto.ru',                       # SEND_FROM
          'server-vm23.LOCAL',                          # SERVER
          555,                                          # PORT
          'skrutko',                                    # USER_NAME
          'ZZZZZZZxxxxxx1111a')                         # PASSWORD

    """
    try:
        send_from = SEND_FROM                                                             
        subject = topic_text                                                               
        text = body_text                                                                 
        files = fr'{file_link.strip()}'
        server = SERVER
        port = PORT
        username=USER_NAME
        password = PASSWORD
        isTls=True

        msg = MIMEMultipart()
        msg['From'] = send_from
        msg['To'] = ','.join(send_to)
        msg["Cc"] = ','.join(send_cc)
        msg["Bcc"] = ','.join(send_bcc)
        msg['Date'] = formatdate(localtime = True)
        msg['Subject'] = subject
        msg.attach(MIMEText(text))

        part = MIMEBase('application', "octet-stream")
        if len(files) > 0: # если есть вложения то прикрепляем их к письму
            part.set_payload(open(files, "rb").read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={file_name.strip()}') # имя файла должно быть на латинице иначе придет в кодировке bin
            msg.attach(part)

        smtp = smtplib.SMTP(server, port)
        if isTls:
            smtp.starttls()
        smtp.login(username, password)
        smtp.sendmail(send_from, send_to+send_cc+send_bcc, msg.as_string())
        smtp.quit()
    except Exception as ex_:
        logger.error(f'ошибка функции {send_mail.__name__} входне параметры {send_to, send_cc, send_bcc, topic_text, body_text, file_link, file_name, SEND_FROM, SERVER, PORT, USER_NAME, PASSWORD} {ex_}')


def update_file(link):
    """обновление сводной таблицы Excel
    # блок импортов для обновления сводных
    import pythoncom
    pythoncom.CoInitializeEx(0)
    import win32com.client
    Args:
        link (_type_): ссылка на файл - который нужно обновить
    !!!!!!
        Настроить эксель:
        Данные - Получить данные - Параметры запроса - Конфеденциальность - Всегда Игнорировать параметры уровней конфеденциальности
    """
    try:
        xlapp = win32com.client.DispatchEx("Excel.Application")
        wb = xlapp.Workbooks.Open(link)
        wb.Application.AskToUpdateLinks = False   # разрешает автоматическое  обновление связей (файл - парметры - дополнительно - общие - убирает галку запрашивать об обновлениях связей)
        wb.Application.DisplayAlerts = True  # отображает панель обновления иногда из-за перекрестного открытия предлагает ручной выбор обновления True - показать панель
        wb.RefreshAll()
        # xlapp.CalculateUntilAsyncQueriesDone() # удержит программу и дождется завершения обновления. было прописано time.sleep(30)
        time.sleep(40) # задержка 60 секунд, чтоб уж точно обновились сводные wb.RefreshAll() - иначе будет ошибка 
        wb.Application.AskToUpdateLinks = True   # запрещает автоматическое  обновление связей / то есть в настройках экселя (ставим галку обратно)
        wb.Save()
        wb.Close()
        xlapp.Quit()
        wb = None # обнуляем сслыки переменных иначе процесс эксель не завершается и висит в дистпетчере
        xlapp = None # обнуляем сслыки переменных иначе процесс эксел ь не завершается и висит в дистпетчере
        del wb # удаляем сслыки переменных иначе процесс эксель не завершается и висит в дистпетчере
        del xlapp # удаляем сслыки переменных иначе процесс эксель не завершается и висит в дистпетчере
    except Exception as ex_:
        wb.Close()
        xlapp.Quit()
        del wb # удаляем сслыки переменных иначе процесс эксель не завершается и висит в дистпетчере
        del xlapp # удаляем сслыки переменных иначе процесс эксель не завершается и висит в дистпетчере
        print(f'ошибка функции {update_file.__name__} {ex_} не удалось обновить файл по ссылке {link}')
        logger.error(f'ошибка функции {update_file.__name__} {ex_} не удалось обновить файл по ссылке {link}')



def agentskie_auto(kommetary:str):
    """ищем в комментарии признаки слов из списка и если да - True
    Args:
        kommetary (str): _description_
    """
    try:
        kommetary = str(kommetary).lower()
        lst_name = ['агентский договор', 'пиагент']
        result = any([name in kommetary for name in lst_name])
        if result: return True
        else: return False
    except Exception as ex_:
        logger.error(f'ошибка функции {agentskie_auto.__name__} входне параметры {kommetary} {ex_}')

def marka_iz_modeli_par_importa(model, marka_2):
    """если в column_marka_2 есть PAR_IMP то из столбца модель - вытаскиваем марку
    Args:
        model (_type_): столбец модель
        marka_2 (_type_): столбец марка_2
    """
    try:
        if 'PAR_IMP' in marka_2 and model!=None:
            model = str(model)
            marka_2 = str(marka_2)
            result_model = model.split()
            if len(result_model) > 0:
                result = result_model[0].strip().upper()
                return result
            else:
                return 'нет данных'
        else: return 'НЕ PAR_IMP'
    except Exception as ex_:
        logger.error(f'ошибка функции {marka_iz_modeli_par_importa.__name__} входне параметры {model, marka_2} {ex_}')

def pravka_kia_import(marka, marka_pi):
    """правка вспомогательного столбца marka_pi для KIA
    если в столбце marka есть KIAimp, то заменить на KIA в столбец marka_pi
    Args:
        marka (_type_): _description_
        marka_pi (_type_): _description_
    Returns:
        _type_: _description_
    """
    try:
        if 'KIAimp' in marka:
            return 'KIA'
        else: return marka_pi
    except Exception as ex_:
            logger.error(f'ошибка функции {pravka_kia_import.__name__} входне параметры {marka, marka_pi} {ex_}')
            return marka_pi

def zamena_marki_par_importa(marka, marka_2, marka_pi):
    # если марка_2 PAR_IMP - тогда вытаскиваем марку из столбца marka_pi и ставим в марку(столбец) в противном счлучае оставляем как есть
    try:
        if 'PAR_IMP' in marka_2:
            return marka_pi
        else: return marka
    except Exception as ex_:
            logger.error(f'ошибка функции {zamena_marki_par_importa.__name__} входне параметры {marka, marka_2, marka_pi} {ex_}')
            return marka_pi

def read_email_users(sheet_name_, name_return_col:str):
    """возвращает значение из фрейма констант по имени константы
    Args:
        sheet_name_ (str): имя листа с которого берем данные
        name_return_col (str): имя столбца из которого возвращаем данные
    Returns:
        _type_: _description_
    """
    try:
        df = pd.read_excel(links_main(fr'{DIR}\main_links.txt','email_users'), sheet_name=sheet_name_)
        df = [str(i).strip() for i in list(df[name_return_col]) if str(i) != 'nan' and '@' in str(i)]
        if len(df)>0:
            return df
        else: 
            logger.info(f"email адресов нет для -  {name_return_col}")
            return None
    except Exception as ex_:
        logger.error(f'ошибка функции {read_email_users.__name__}  {ex_} входные арг {sheet_name_, name_return_col}')
        return 0

def pravka_marki_OVP(marka, marka_2, marka_auto_ovp, model):
    """только для ОВП чистого ОВП без некст хуекст
    если есть совпадение по столбцам марки и марки 2, то возвращаем марку авто ОВП

    Args:
        marka (_type_): _description_
        marka_2 (_type_): _description_
        marka_auto_ovp (_type_): _description_

    Returns:
        _type_: _description_
    """
    try:
        if marka == "OVP" and marka_2 == "OVP" and marka_auto_ovp!=None:
            result = str(marka_auto_ovp).strip().upper()
            return result
        else: return model
    except Exception as ex_:
        logger.error(f'ошибка функции {pravka_marki_OVP.__name__}  {ex_} входные арг {marka, marka_2, marka_auto_ovp}')
        return model
    
def rename_OVP_names(marka_2):
    """правка ОВП в столбце макка_2
    Args:
        marka_2 (_type_): _description_
    """
    try:
        lst_names_OVP = ['OVP_JETOUR', 'OVP_MAZDA']
        res = any([i in marka_2 for i in lst_names_OVP])
        if res:
            return 'OVP'
        else: return marka_2
    except Exception as ex_:
        logger.error(f'ошибка функции {rename_OVP_names.__name__}  {ex_} входные арг {marka_2}')
        return marka_2
    
def pravka_huyauka(marka, marka_2, plan):
    """справка столбца марка Планов так как они где-то выше наебнулись функциями

    Args:
        marka (_type_): _description_
        marka_2 (_type_): _description_
        plan (_type_): _description_
    """
    try:
        if marka == 'НЕ PAR_IMP' and marka_2 == 'PAR_IMP' and plan != None:
            return 'PAR_IMP'
        else: return marka
    except Exception as ex_:
        logger.error(f'ошибка функции {pravka_huyauka.__name__}  {ex_} входные арг {marka, marka_2, plan}')
        return marka

def podpis_modeli_plan(model, plan):
    try:
        if plan != None and model ==None:
            return "! ПЛАН"
        else: return model
    except Exception as ex_:
        logger.error(f'ошибка функции {podpis_modeli_plan.__name__}  {ex_} входные арг {model, plan}')
        return model   


def yesterday_new(days:int=1, simbol:str='-' or '+'):
    """возвращает дату на вчера - по уморлчанию минус 1 день

    Args:
        days (int, optional): кол-во дней от текущей. Defaults to 1.
        simbol (str, optional): прибавляем или отнимаем. Defaults to '-'or'+'.

    Returns:
        _type_: datetime
    """
    from datetime import datetime, timedelta
    try:
        if simbol == '+':
            date = datetime.now()
            new_date = date + timedelta(days=days)# вычитание одного дня
            return new_date
        else:
            date = datetime.now()
            new_date = date - timedelta(days=days)# вычитание одного дня
            return new_date
    except Exception as ex_:
        print(f'ошибка функции {yesterday_new.__name__}  {ex_}')
        
        
logger.info(f'считываем email-ы пользователей для отправки тем_ДИР')  
try: 
    KUM_USER_EMAIL = read_email_users('user_temp_dir', 'email') if read_email_users('user_temp_dir', 'email') != None else list()
    KUM_USER_EMAIL_CC = read_email_users('user_temp_dir', 'email_cc') if read_email_users('user_temp_dir', 'email_cc') != None else list()
    KUM_USER_EMAIL_BCC = read_email_users('user_temp_dir', 'email_bcc') if read_email_users('user_temp_dir', 'email_bcc') != None else list() 
    print("считываем емэйлы пользователей")

except Exception as ex_:
        logger.error(f'не удалось считать email-ы {ex_}')


logger.info("Получаем парметры пользователя для работы с почтой")
try:
    NAME_PC, SERVER, PORT, USER_NAME, SEND_FROM, PASSWORD = get_data_user_param()
except Exception as ex_:
    logger.error(f"Ошибка получения данных: {ex_}")

logger.info("Проверка полученных параметров")

resul_mail_param = all([i is not None for i in [NAME_PC, SERVER, PORT, USER_NAME, SEND_FROM, PASSWORD]])

if resul_mail_param:
    logger.info(f"Все данные для работы с почтой присутствуют {resul_mail_param}")
else:
    logger.error(f"Некоторые данные для работы с почтой отсутствуют {resul_mail_param} {get_data_user_param()}")


logger.info(f'считываем данные подключения SQL')
try:
    SERVER_SQL = links_main(fr"{DIR}\main_links.txt", "server_sql")
    DATABASE_SQL = links_main(fr"{DIR}\main_links.txt", "database_sql")
    USERNAME_SQL = links_main(fr"{DIR}\main_links.txt", "username_sql")
    PASSWORD_SQL = links_main(fr"{DIR}\main_links.txt", "password_sql")
except Exception as ex:
    logger.error(f'не удалось получить данные подключения SQL')


logger.info(f"считываем result_temp_sql из SQL")
try:
    df_result_temp_sql = pd.read_sql('select * from result_temp_sql', connect_bd_SQL(SERVER_SQL, DATABASE_SQL, USERNAME_SQL, PASSWORD_SQL))
except Exception as e:
    logger.error(f"ошибка - при считывании result_temp_sql из SQL {e}")


logger.info(f'определяем агентские авто - добавляем столбец - dogovor_agent') 
try:
    df_result_temp_sql['dogovor_agent'] = df_result_temp_sql['комментарий'].apply(lambda x: agentskie_auto(x))
except Exception as ex_:
        logger.error(f'не удалось добавить агентские авто {ex_}')


logger.info(f'определяем марку для PAR_IMP - добавляем столбец - marka_pi') 
try:
    df_result_temp_sql['marka_pi']= df_result_temp_sql[['модель','марка_2']].apply(lambda x: marka_iz_modeli_par_importa(x.модель, x.марка_2), axis=1)
except Exception as ex_:
        logger.error(f'не удалось определить марку для PAR_IMP {ex_}')


logger.info(f'правим KIAPI') 
try:
    df_result_temp_sql['marka_pi']= df_result_temp_sql[['марка','marka_pi']].apply(lambda x: pravka_kia_import(x.марка, x.marka_pi), axis=1)
except Exception as ex_:
        logger.error(f'не удалось исправить KIAPI {ex_}')


logger.info(f'правим столбец марка по тем у кого в марка_2 признак PAR_IMP') 
try:
    df_result_temp_sql['марка'] = df_result_temp_sql[['марка','марка_2','marka_pi']].apply(lambda x: zamena_marki_par_importa(x.марка, x.марка_2,x.marka_pi), axis=1)
except Exception as ex_:
        logger.error(f'не удалось исправить столбец марка по тем у кого в марка_2 признак PAR_IMP {ex_}')


logger.info(f'правим столбец марка по ОВП') 
try:
    df_result_temp_sql['модель'] = df_result_temp_sql[['марка','марка_2','marka_auto','модель']].apply(lambda x: pravka_marki_OVP(x.марка, x.марка_2,x.marka_auto, x.модель), axis=1)
except Exception as ex_:
        logger.error(f'не удалось исправить столбец марка по ОВП {ex_}')

logger.info(f'правим столбец марка_2 по ОВП') 
try:
    df_result_temp_sql['марка_2'] = df_result_temp_sql['марка_2'].apply(lambda x: rename_OVP_names(x))
except Exception as ex_:
        logger.error(f'не удалось исправить столбец марка_2 по ОВП {ex_}')


logger.info(f'правим столбец марка с наебнувшимися планами') 
try:
    df_result_temp_sql['марка'] = df_result_temp_sql[['марка', 'марка_2', 'ПЛН']].apply(lambda x: pravka_huyauka(x.марка, x.марка_2, x.ПЛН), axis=1)
except Exception as ex_:
        logger.error(f'не удалось исправить столбец марка с наебнувшимися планами {ex_}')


logger.info(f'правим столбец модель') 
try:
    df_result_temp_sql['модель'] = df_result_temp_sql[['модель', 'ПЛН']].apply(lambda x: podpis_modeli_plan(x.модель, x.ПЛН), axis=1)
except Exception as ex_:
        logger.error(f'не удалось исправить столбец модель  {ex_}')


YEAR_COSTRATION = 2025
logger.info(f'обрезаем базу с {YEAR_COSTRATION}') 
try:
    df_result_temp_sql = df_result_temp_sql[df_result_temp_sql['дата'].dt.year >=YEAR_COSTRATION]
except Exception as ex_:
        logger.error(f'❌ ошибка обрезания базы  {ex_}')


logger.info(f'удаляем лишние столбцы') 
try:
    cols_to_drop = ["maneger_buy", "dopy", "podarok", "pp_remont", "oformlenie", "gai", "rezina", "itog_dohod_ovp", "prejniy_vladelec", "istochnik_am", "день"]
    df_result_temp_sql = df_result_temp_sql.drop(columns=cols_to_drop)
except Exception as ex_:
        logger.error(f'❌ удаления столбцов  {ex_}')


logger.info(f'обрезаем даты по вчера') 
try:
    df_result_temp_sql = df_result_temp_sql[df_result_temp_sql['дата']<=yesterday_new(1,'-')]
except Exception as ex_:
        logger.error(f'❌ обрезание дат {ex_}')

df_result_temp_dir = copy.deepcopy(df_result_temp_sql)

df_result_temp_dir.to_excel(links_main(fr"{DIR}\main_links.txt", 'save_result'), index=False)


logger.info(f'тестирование столбцов на запись в SQL') 
try:
    exception_column_SQL(df_result_temp_dir, SERVER_SQL, DATABASE_SQL, USERNAME_SQL, PASSWORD_SQL)
except Exception as ex_:
        logger.error(f'❌ ошибка тестирования столбцов  {ex_}')


name_save_file_sql = 'df_result_temp_dir'
logger.info(f'запись темпа ДИР - в SQL под именем {name_save_file_sql}') 
try:
    df_result_temp_dir.to_sql(name_save_file_sql, con=connect_bd_SQL(SERVER_SQL, DATABASE_SQL, USERNAME_SQL, PASSWORD_SQL), if_exists='replace', index=False)
    logger.info(f'✅ {name_save_file_sql} запиcан в SQL') 
except Exception as ex_:
        logger.error(f'❌ ошибка записи темпа ДИР - в SQL {ex_}')


logger.info(f"запускаем обновление темп_ДИР")
try:
    update_file(links_main(fr'{DIR}\main_links.txt','update_dashboard'))   # добавить ссылку ан обновляемый файл
except Exception as ex_:
        logger.error(f'не удалось обновить темп_ДИР {ex_}')


logger.info(f"рассылаем Темпы_ИД")
try:
    TOPIC_TEXT = fr"""Темпы_ИД на {yesterday().strftime('%d-%m-%Y')}  service_message"""
    BODY_TEXT = f"Здравствуйте \nВо вложении Темпы_ИД на {yesterday().strftime('%d-%m-%Y')}"
    send_mail(KUM_USER_EMAIL,
            KUM_USER_EMAIL_CC,
            KUM_USER_EMAIL_BCC,
            TOPIC_TEXT,
            BODY_TEXT,
            links_main(fr'{DIR}\main_links.txt','update_dashboard'),
            os.path.basename(links_main(fr'{DIR}\main_links.txt','update_dashboard')),
            SEND_FROM,
            SERVER,
            PORT,
            USER_NAME,
            PASSWORD)
    logger.info(f"рассылка Темпы_ИД прошла успешно для следующих пользователей {KUM_USER_EMAIL, KUM_USER_EMAIL_CC, KUM_USER_EMAIL_BCC}")
except Exception as ex_:
        logger.error(f'не удалось разослать Темпы_ИД {ex_}')
