import pandas as pd
import numpy as np
import copy
# блок логирования
import logging
from functools import wraps

# блок импортов для обновления сводных
import pythoncom
pythoncom.CoInitializeEx(0)
import win32com.client, time, os, time
import warnings
warnings.filterwarnings('ignore')
# блок импорта отправки почты
import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

import sqlalchemy
import socket

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
LOG_FILE = os.path.join(DIR, "log_temp.log")

# Удаляем старый лог, если он существует
if os.path.exists(LOG_FILE):
    try:
        os.remove(LOG_FILE)
        print(f"🗑️ Старый лог удалён: {LOG_FILE}")
    except PermissionError:
        print(f"❌ Не удалось удалить старый лог: возможно, файл открыт в другом приложении")
    except Exception as e:
        print(f"❌ Ошибка при удалении лога: {e}")

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
logger.info("Запуск скрипта ТЕМП обычный")


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

def get_data_user_param(socket_name=True):
    """возврат данных пользователя по соккету (имени ПК) для отправки почты в формате кортежа
    name_pc server port username send_from password

    Args:
        socket_name (bool, optional): _description_. Defaults to True.

    Returns:
        _type_: _description_
    """
    try:
        df = pd.read_excel(links_main(fr"{DIR}\file_links.txt", "pass"),sheet_name='PC')
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

def sort_columns(pattern, word_search, word):
    """фильтрует данные

    Args:
        pattern (_type_): слво для первоначального поиска
        word_search (_type_): слово точного совпадения
        word (_type_): значение где ищем

    Returns:
        _type_: _description_
    """
    if pattern in word:
        if word_search==word:
            return word
    else:
        return word


def OVP_YAR(marka, region, ploshchadka):
    """разделяет ОВП ЯР на Яр и РЫБ

    Args:
        marka (_type_): _description_
        region (_type_): _description_
        ploshchadka (_type_): _description_

    Returns:
        _type_: _description_
    """
    try:
        if marka == 'OVP' and region == 'YAR' and ('Ярославль' in ploshchadka or 'Рыбинск' in ploshchadka):
            if ploshchadka == 'Ярославль':
                return 'YAR'
            elif ploshchadka == 'Рыбинск':
                return 'RYB'
        else:
            return region
    except Exception as ex_:
        logger.error(f'{OVP_YAR.__name__} {ex_}')

def saratov_marka(region, marka, model):
    """разделение саратова на OMODA JAECOO

    Args:
        region (_type_): _description_
        marka (_type_): _description_
        model (_type_): _description_

    Returns:
        _type_: _description_
    """
    try:
        if region=='SAR' and marka=='OMODA':
            if 'OMODA' in str(model).strip().upper():
                return 'OMODA'
            elif 'JAECOO' in str(model).strip().upper():
                return 'JAECOO'
        else:
            return marka
    except Exception as ex_:
        logger.error(f'{saratov_marka.__name__} {ex_}')


def autocentr_rename_all(word):
    try:
        if 'HYUNDAIpi'== word: return 'PAR_IMP'
        elif 'KIAPI'== word: return 'KIAimp'
        elif 'MAZDApi' == word: return f'MAZDAimp'
        elif 'VOLKSWAGENpi' == word: return f'VOLKSWAGENimp'
        elif 'PARimp' == word: return f'PAR_IMP'
        else:
            return word
    except Exception as ex_:
        logger.error(f'{autocentr_rename_all.__name__} {ex_}')


def jetoor_msk(marka, region, komment ):
    """ делит JETOUR МСК на JETOUR OVP_JETOU

    Args:
        marka (_type_): _description_
        region (_type_): _description_
        komment (_type_): _description_

    Returns:
        _type_: _description_
    """
    try:
        if marka == 'JETOUR' and region=='MSK' and str(komment).strip() == 'б/у':
            return 'OVP_JETOUR'
        else:
            return marka
    except Exception as ex_:
        logger.error(f'{jetoor_msk.__name__} {ex_}')


def mazda_msk_next(marka, region, komment, statys_original):
    """ делит JETOUR МСК на JETOUR OVP_JETOU

    Args:
        marka (_type_): _description_
        region (_type_): _description_
        komment (_type_): _description_

    Returns:
        _type_: _description_
    """
    try:
        if marka == 'MAZDA' and region=='MSK' and ('б/у' in str(komment).strip() or 'б/у' in str(statys_original).strip()):
            return 'OVP_MAZDA'
        else:
            return marka
    except Exception as ex_:
        logger.error(f'{mazda_msk_next.__name__} {ex_}')


def individ_date_plan(year, month):
    try:
        year = str(year)
        month = str(month)
        month = month if len(month)==2 else '0'+month
        day = '01'
        return f'{year}-{month}-{day}'
    except Exception as ex_:
        logger.error(f'Ошибка функции {individ_date_plan.__name__} {ex_} не удалось преобразовать {year}{month}')


def kre_nal(vidacha, vid_opl):
    """приводит вид оплаты в кре нал т.е. bool 1 0

    Args:
        vidacha (_type_): _description_
        vid_opl (_type_): _description_

    Returns:
        _type_: _description_
    """
    spisok_kre = ['кре', 'банк', 'фин', 'лиз', 'fin', 'liz', 'bank']
    if vidacha == 1:
        if any([i in str(vid_opl).strip() for i in spisok_kre]):
            return 1
        else:
            return 0
    else:
        return 0
    

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
        logger.error(f'ошибка функции {yesterday_new.__name__}  {ex_}')
        
        
def korp_rozn(klient):
    """разделяет клиентов на корп и розницу
    если корп возвращает 1

    Args:
        klient (str): _description_

    Returns:
        int: 1 or 0
    """
    try:
        list_sort = set(['ООО', 'ПАО', 'ЗАО', 'ОАО', 'АО', 'ВТБ', 'ГПБЛ', 'ИП', 'ПАО', 'САО', 'ФБУ', 'ФГУП'])
        res = any([i in str(klient) for i in list_sort])
        if res:
            return 1
        else:
            return 0
    except Exception as ex_:
        logger.error(f'ошибка функции {korp_rozn.__name__}  {ex_}')


def kre_nal_ovp(marka, kredit, kommentariy, vidacha):
    """распределение на кре нал для ОВП по примечанию

    Args:
        marka (_type_): _description_
        kredit (_type_): _description_
        klient (_type_): _description_

    Returns:
        _type_: _description_
    """
    try:
        if str(marka) == 'OVP' and vidacha==1.0:
            if 'кредит' in str(kommentariy).lower():
                return 1
            else:
                return 0
        else:
            return kredit
    except Exception as ex_:
        logger.error(f'ошибка функции {kre_nal_ovp.__name__}  {ex_}')       


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
        #xlapp.CalculateUntilAsyncQueriesDone() # удержит программу и дождется завершения обновления. было прописано time.sleep(30)
        time.sleep(30) # задержка 60 секунд, чтоб уж точно обновились сводные wb.RefreshAll() - иначе будет ошибка 
        wb.Application.AskToUpdateLinks = True   # запрещает автоматическое  обновление связей / то есть в настройках экселя (ставим галку обратно)
        wb.Save()
        wb.Close()
        xlapp.Quit()
        wb = None # обнуляем сслыки переменных иначе процесс эксель не завершается и висит в дистпетчере
        xlapp = None # обнуляем сслыки переменных иначе процесс эксел ь не завершается и висит в дистпетчере
        del wb # удаляем сслыки переменных иначе процесс эксель не завершается и висит в дистпетчере
        del xlapp # удаляем сслыки переменных иначе процесс эксель не завершается и висит в дистпетчере
    except Exception as ex_:
        logger.error(f'ошибка функции {update_file.__name__} {ex_} не удалось обновить файл по ссылке {link}')


def read_email(link):
    try:
        df_email = pd.read_excel(link)
        res = list(df_email['email'])
        return res
    except Exception as ex_:
        logger.error(f'ошибка функции {read_email.__name__} {ex_} входне параметры {link}')


def read_email_adress(link, name_columns = 'email'):
    """Функция считывания адресатов для рассылки

    Args:
        link (_type_, str):  ссылка на файл с адресами
        name_columns (str): имя столбца с адресами

    Returns:
        _type_: возфращает строку со списком email
    """
    try:
        em_list = pd.read_excel(link)
        lst = list(em_list[name_columns])
        except_ = all('nan' in str(i) for i in lst) # проверка списка на nan 
        return [] if except_ == True else [i for i in lst if 'nan' not in str(i)]       # если список с адресами с nan - заменяем на пустой список (иначе будет ошибк апри отправке письма)
    except Exception as ex_:
        logger.error(f'Ошибка функции {read_email_adress.__name__} {name_columns} {ex_}')


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

        file_link(str):  ссылка на файл
        file_name(str):  имя файла в данном варианте нужно указывать с расширением 'BAIC_MSK.xlsx' (имя должно быть на латинице иначе придет в кодировке bin)

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
        logger.error(f'ошибка функции {send_mail.__name__} {ex_} входне параметры {send_to, send_cc, send_bcc, topic_text, body_text, file_link, file_name, SEND_FROM, SERVER, PORT, USER_NAME, PASSWORD}')


# правка 17.11.25
def proverka_daty(data_):
    """проверяет типы дат,
    если тип не является датой возвращает None

    Args:
        data_ (_type_): _description_

    Returns:
        _type_: _description_
    """
    try:
        if 'timestamp' in str(type(data_)).lower() or 'datetime' in str(type(data_)).lower():
            return data_
        else: return None
    except Exception as ex_:
        logger.error(f'ошибка функции {proverka_daty.__name__} {ex_} входные параметры {data_}')


# правка 17.11.25
def poluchit_day_iz_daty(data_):
    """получает день из даты

    Args:
        data_ (_type_): _description_

    Returns:
        _type_: _description_
    """
    try:
        return data_.day
    except Exception as ex_:
        logger.error(f'ошибка функции {poluchit_day_iz_daty.__name__} {ex_} входные данные {data_}')
        return None

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

def exception_column_SQL(df, server:str, 
                         database:str, 
                         username:str, 
                         password:str):
    
    """проверяет наличие ошибок в df при записи в SQL
    перебирает каждый столбец и записывает в тестовую таблицу БД SQL
    при возникновении ошибки в записи - возвращает имя столбца и ошибку
    чаще такие ошибки вызван ытипом данных inf 
    Args:
        df (_type_): _description_
    """
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
            logger.info(f'столбец {i} - ок')
        except Exception as ex_:
            logger.error(f'{exception_column_SQL.__name__} ОШИБКА в столбце {i}  возможно порблемы  стипом данных {ex_}')



def connect_bd_SQL(server:str, 
                   database:str, 
                   username:str, 
                   password:str):
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

def rename_autocenry_st_2(df, search_col:str, search_name:str, result_col:str)->str:
    """_summary_

    Args:
        df (_type_): фрейм по которому ищем автоцентры для переименования
        search_col (_type_): колока в которой ищем
        search_name (_type_): имя которое ищем
        result_col (_type_): колонка по которой получем результат

    Returns:
        str: новое_имя 
    """
    try:
        res_list = list(df[df[search_col]==search_name][result_col])
        if len(res_list)>0:
            return res_list[0]
        else:
            return search_name
    except Exception as ex_:
        logger.error(f'ошибка функции {rename_autocenry_st_2.__name__} {ex_}')

def clean_string(text):
    try:
        import re
        text = str(text).upper()
        cleaned = re.sub(r'[^a-zA-Zа-яА-ЯёЁ0-9\s]', ' ', str(text))
        cleaned = re.sub(r'\s+', ' ', cleaned)
        return cleaned.strip()
    except Exception as ex_:
        logger.error(f'ошибка функции {clean_string.__name__} {ex_}')   


def korrektirovka_vidach_HYUNDAIpi_MSK(klient, data_, key, vidacha):
    """корректировка видач параллельного импорта HYUNDAIpi_MSK который передают в ЯР
    Args:
        klient (_type_): клиент
        data_ (_type_): дата
        key (_type_): ключ 
        vidacha (_type_): выдача столбец с готовыми значениями
    Returns:
        _type_: _description_
    """
    from datetime import datetime
    date_kostracii = "2026-03-01"
    date_kostracii = datetime.strptime(date_kostracii, "%Y-%m-%d")

    if key == 'HYUNDAIpi_MSK':
        if 'ООО СИМ' in clean_string(klient) and data_ >= date_kostracii: return 0
        return vidacha
    else: return vidacha

logger.info(f"Получаем данны для подключенияServer SQL")
try:
    SERVER_SQL = links_main(fr"{DIR}\file_links.txt", "server_sql")
    DATABASE_SQL = links_main(fr"{DIR}\file_links.txt", "database_sql")
    USERNAME_SQL = links_main(fr"{DIR}\file_links.txt", "username_sql")
    PASSWORD_SQL = links_main(fr"{DIR}\file_links.txt", "password_sql")
    df_rename_au = pd.read_excel(links_main(fr'{DIR}\file_links.txt','rename_au'), sheet_name='autocentry')
except Exception as e:
    logger.error(f"Ошибка получения данных для подключения к SQL {e}")


logger.info(f"Получаем фреймы с ПЛАНОМ и фрейм для ОБЪЕДИНЕНИЯ по брендам, фрейм из SQL с данными для темпа")
try:
    CONNECTION_BRAND_PLAN_AUTO = pd.read_excel(links_main(fr'{DIR}\file_links.txt', 'connection_brand'), sheet_name='PLAN_AUTO')
    PLAN_AUTO = pd.read_excel(links_main(fr'{DIR}\file_links.txt', 'plan_auto'), sheet_name='auto')
    KOSTRACIA = '2023-01-01'

    sql_query = """  
    SELECT *
    FROM first_class_oborotka_for_temp;
    """
    df = pd.read_sql(sql_query, con=connect_bd_SQL(SERVER_SQL, DATABASE_SQL, USERNAME_SQL, PASSWORD_SQL))
except Exception as e:
    logger.error(f"Ошибка получения данных для подключения к SQL или иного фрейма {e}")


logger.info(f"Наводим порядок со столбцами")
try:
    black_list = ['id', 'vin_novogo', 'model_novogo', 'дата_прихода_на_склад', 
                'дата_полной_оплаты_факт', 'дата_справки_счет_факт','с_листа', 'ссылка', 
                'сотрудник_продал', 'склад_заказ',	'в_ар_хив', 'получено_за_ам_руб']
    # df = df[df['принадлежность']!='SCLAD_OMODA_SAR.xlsx'] # какого-то лешего затесался склад SCLAD_OMODA_SAR.xlsx
    df = df[[i for i in df.columns if i not in black_list]]
    df = df.rename(columns={'дата_изм':'дата_отказа'})
except Exception as e:
    logger.error(f"Ошибка при наведении порядка со столбцами {e}")

logger.info(f"Меняем тип данных to_datetime")
try:
    df['дата_отказа'] = pd.to_datetime(df['дата_отказа'], errors='coerce') # errors='ignore'
except Exception as e:
    logger.error(f"Ошибка приведения к типу to_datetime {e}")


logger.info(f"Создаем копии фреймов для выдачи заказ отказ")
try:
    df_vidacha = copy.deepcopy(df)
    df_zakaz = copy.deepcopy(df)
    df_otkaz = copy.deepcopy(df)
except Exception as e:
    logger.error(f"Ошибка приведения к типу to_datetime {e}")


logger.info(f"Обрабатываем выдачу")
try:
    column_name_vidacha = 'дата_выдачи_факт'
    df_vidacha = df_vidacha[[sort_columns('дата', column_name_vidacha, i) for i in df_vidacha.columns if sort_columns('дата', column_name_vidacha, i)!=None]]
    df_vidacha['выдача'] = df_vidacha[column_name_vidacha].apply(lambda x: 1 if len(str(x))>5 else 0)
    df_vidacha = df_vidacha.rename(columns={column_name_vidacha:'дата'})
except Exception as e:
    logger.error(f"Ошибка при обработке выдачи {e}")


logger.info(f"Обрабатываем заказы")
try:
    column_name_zakaz = 'дата_заказа'
    df_zakaz = df_zakaz[[sort_columns('дата', column_name_zakaz, i) for i in df_zakaz.columns if sort_columns('дата', column_name_zakaz, i)!=None]]
    df_zakaz['заказ'] = df_zakaz[column_name_zakaz].apply(lambda x: 1 if len(str(x))>5 else 0)
    df_zakaz = df_zakaz.rename(columns={column_name_zakaz:'дата'})
except Exception as e:
    logger.error(f"Ошибка при обработке заказы {e}")


logger.info(f"Обрабатываем отказы")
try:
    column_name_otkaz = 'дата_отказа'
    df_otkaz = df_otkaz[[sort_columns('дата', column_name_otkaz, i) for i in df_otkaz.columns if sort_columns('дата', column_name_otkaz, i)!=None]]
    df_otkaz['отказ'] = df_otkaz[column_name_otkaz].apply(lambda x: 1 if len(str(x))>5 else 0)
    df_otkaz = df_otkaz.rename(columns={column_name_otkaz:'дата'})
except Exception as e:
    logger.error(f"Ошибка при обработке отказы {e}")


logger.info(f"Обрабатываем тип данных to_datetime по заказу и выдаче")
try:
    df_vidacha['дата'] = pd.to_datetime(df_vidacha['дата'], format='mixed', errors='coerce')
    df_zakaz['дата'] = pd.to_datetime(df_zakaz['дата'], format='mixed', errors='coerce')
except Exception as e:
    logger.error(f"Ошибка при обработке типа данных to_datetime по заказу и выдаче {e}")


logger.info(f"Объединяем фреймы")
try:
    result = pd.concat([df_vidacha, df_zakaz])
except Exception as e:
    logger.error(f"Ошибка объединения {e}")


logger.info(f"Правим дату")
try:
    result['дата'] = result['дата'].apply(proverka_daty)
    result['день'] = result['дата'].apply(poluchit_day_iz_daty)
except Exception as e:
    logger.error(f"Ошибка правки даты {e}")


logger.info(f"Удаляем лишнее")
try:
    result = result.dropna(subset='дата')
except Exception as e:
    logger.error(f"Ошибка удаления {e}")

logger.info(f"Обрабатываем столбцы марка и регион")
try:
    result['марка'] = result['принадлежность'].apply(lambda x: str(x).split('_')[1] )
    result['регион'] = result['принадлежность'].apply(lambda x: str(x).split('_')[-1].split('.')[0] )
    result['регион'] = result.apply(lambda x: (OVP_YAR(x.марка, x.регион, x.площадка)), axis=1)
    result['марка'] = result.apply(lambda x: (saratov_marka(x.регион, x.марка, x.модель)), axis=1)
except Exception as e:
    logger.error(f"Ошибка обработки марки и региона {e}")


logger.info(f"Обрабатываем марку")
try:
    result['марка'] = result.apply(lambda x: (autocentr_rename_all(x.марка)), axis=1)
    result['марка'] = result.apply(lambda x: (jetoor_msk(x.марка, x.регион, x.комментарий)), axis=1)
    result['марка'] = result.apply(lambda x: (mazda_msk_next(x.марка, x.регион, x.комментарий, x.статус_оригинал)), axis=1)
except Exception as e:
    logger.error(f"Ошибка обработки марки{e}")

logger.info(f"Собираем планы")
try:
    PLAN_AUTO_2 = copy.deepcopy(PLAN_AUTO)
    PLAN_AUTO_2['календарь'] = PLAN_AUTO_2.apply(lambda x: (individ_date_plan(x.year, x.mnth)), axis=1)
    PLAN_AUTO_2['календарь'] = pd.to_datetime(PLAN_AUTO_2['календарь'])
    PLAN_AUTO_2 = PLAN_AUTO_2[PLAN_AUTO_2['type_ind'] == 'Авто'][['календарь' , 'reg', 'item_ind', 'zone','ПЛН']]
    PLAN_AUTO_2 = PLAN_AUTO_2[abs(PLAN_AUTO_2['ПЛН']) > 0]
    PLAN_AUTO_2 = PLAN_AUTO_2.merge(CONNECTION_BRAND_PLAN_AUTO, how='left')[['календарь','марка_фильтр',  'регион_фильтр', 'ПЛН']]
    PLAN_AUTO_2 = PLAN_AUTO_2.rename(columns={'марка_фильтр':'марка', 'регион_фильтр':'регион', 'календарь':'дата'})
except Exception as e:
    logger.error(f"Ошибка сбора планов{e}")

logger.info(f"Объединяем основй фрейм и планы")
try:
    result_svod = pd.concat([result, PLAN_AUTO_2])
except Exception as e:
    logger.error(f"Ошибка объединения{e}")

logger.info(f"Обрабатываем кредиты")
try:
    result_svod['кредит'] = result_svod.apply(lambda x: kre_nal(x.выдача, x.форма_оплаты), axis=1)
    result_svod['кредит'] = result_svod.apply(lambda x: kre_nal_ovp(x.марка, x.кредит, x.комментарий, x.выдача), axis=1)
    result_svod['корпоратив'] = result_svod.apply(lambda x: korp_rozn(x.клиент), axis=1)
except Exception as e:
    logger.error(f"Ошибка обратоки кредитов {e}")


logger.info(f"Правка даты")
try:
    result_svod['дата'] = result_svod['дата'].apply(proverka_daty)
except Exception as e:
    logger.error(f"Ошибка правки даты {e}")

logger.info(f"Кастрируем фрейм")
try:
    result_svod = result_svod[result_svod['дата']>=KOSTRACIA]
    result_svod = result_svod[result_svod['дата']<=yesterday_new(31, '+')]
except Exception as e:
    logger.error(f"Ошибка кастрации {e}")

logger.info(f"Переименовываем все pi imp в паралельный импорт")
try:
    result_svod['марка_2'] = result_svod['марка'].apply(lambda x: (rename_autocenry_st_2(df_rename_au, 'name', str(x), 'new_name')))
except Exception as e:
    logger.error(f"Ошибка переименования все pi imp в паралельный импорт {e}")


logger.info(f'вносим корректировки в выдачи по HYUNDAIpi_MSK авто которые передавались в ЯР')
try:
    result_svod['выдача'] = result_svod.apply(lambda x: korrektirovka_vidach_HYUNDAIpi_MSK(x.клиент, x.дата, x.ключ, x.выдача), axis=1)
    result_svod['заказ'] = result_svod.apply(lambda x: korrektirovka_vidach_HYUNDAIpi_MSK(x.клиент, x.дата, x.ключ, x.заказ), axis=1)
except Exception as ex_:
    logger.error(f'в данных result_temp_sql не обнаружено значений inf - пробуем запистаь данные в SQL')


# предобрабатываем для сохранения в SQL
result_temp_sql = result_svod.copy()
try:
    result_temp_sql = result_temp_sql.replace([np.inf, -np.inf], np.nan)
    result_temp_sql.to_sql('result_temp_sql', con=connect_bd_SQL(SERVER_SQL, DATABASE_SQL, USERNAME_SQL, PASSWORD_SQL), if_exists='replace', index=False)
except Exception as ex_:
    logger.info(f'в данных result_temp_sql не обнаружено значений inf - пробуем запистаь данные в SQL')
    try:
        result_temp_sql.to_sql('result_temp_sql', con=connect_bd_SQL(SERVER_SQL, DATABASE_SQL, USERNAME_SQL, PASSWORD_SQL), if_exists='replace', index=False)
    except Exception as ex_:
        logger.error(f'не удалось записать result_temp_sql в SQL запущена процедура поиска ошибок')
        exception_column_SQL(result_temp_sql)


logger.info(f'обновляем дашборд ТЕМП')
try:
    update_file(links_main(fr'{DIR}\file_links.txt', 'uptate_dashboard'))
except Exception as e:
    logger.error(f"Ошибка обновления дашборда ТЕМП{e}")


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

# # отправляем почту 
logger.info(f'отправялем почту')

try:
    DATA_ = yesterday().strftime('%d-%m-%Y')

    TEXT_BODY_0 = """Здравствуйте"""
    TEXT_BODY_1 = """
Во вложении темпы на """
    TEXT_BODY_2 =""" """

    send_mail(read_email_adress(links_main(fr'{DIR}\file_links.txt', 'email_adress'), 'email'), 
            read_email_adress(links_main(fr'{DIR}\file_links.txt', 'email_adress'), 'email_cс'), 
            read_email_adress(links_main(fr'{DIR}\file_links.txt', 'email_adress'), 'email_bcc'), 
            f'Темпы на {DATA_} service_message', 
            f'{TEXT_BODY_0} \n{TEXT_BODY_1} {DATA_} \n\n{TEXT_BODY_2}', 
            links_main(fr'{DIR}\file_links.txt', 'uptate_dashboard'), 
            'temp.xlsx', 
            SEND_FROM, 
            SERVER, 
            PORT, 
            USER_NAME, 
            PASSWORD)
    
except Exception as ex_:
    logger.error(f'Не удалось отправить почту')
