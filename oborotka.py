import os
import time
script_dir = os.path.dirname(os.path.abspath(__file__)) # привет пути )))
print(script_dir)
# time.sleep(3)

# блок логирования
import logging
logging.basicConfig(level=logging.INFO, filename=fr"{script_dir}\py_log.log",filemode="w", format="%(asctime)s %(levelname)s %(message)s")

import copy
import pandas as pd
# pd.options.display.max_colwidth = 100 # увеличить максимальную ширину столбца
# pd.set_option('display.max_columns', None) # макс кол-во отображ столбц
import datetime as DT
from datetime import timedelta
import xlrd


import shutil

# блок импортов для обновления сводных
import pythoncom
pythoncom.CoInitializeEx(0)
import win32com.client
import time

import warnings
warnings.filterwarnings('ignore')

# блок импорта отправки почты
import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders


def LOG_inf(name, type_='INFO' or 'ERROR', *args):
    try:
        if type_ == 'INFO': logging.info(f"{name} {args}")
        elif type_ == 'ERROR': logging.error(f"{name} {args}")
    except Exception as ex_:
        print(f'Х_НЯ с логированием {ex_}')
      
        
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
        file = pd.read_csv(name_file, sep=':')
        result = list(file[file['ключ']==key]['значение'])[0]
        return result
    except Exception as ex_:
        print(f'ошибка функции {links_main.__name__} не удалось считать файл {name_file} или данные в нем {key} ошибка {ex_}')
        
        
def file_update(link):
    
    "возвращает дату последнего обновления файла"
    try:
        from datetime import datetime, date, timedelta
        res = datetime.fromtimestamp(os.path.getmtime(link))
        return res
    except Exception as ex_:
        print(f'ошибка функции {file_update.__name__} не удалось считать метаданные файла {link} ошибка {ex_}')
        
        
KOSTRACIVA = '2023-01-01' # дата до которой все обрезаем
LOG_inf(f'Дата по которую все обрезаем ', 'INFO', KOSTRACIVA)
PARK = pd.read_excel(links_main(fr"{script_dir}/file_links.txt", "park"), sheet_name='БД')
PARK['мес'] = PARK['мес'].fillna(0).apply(lambda x: xlrd.xldate_as_datetime(x, 0))
PARK = PARK[['мес', 'Подразделение/площадка', 'ТИП', 'Марка', 'Бонус', 'Доход, руб.']]
CONNECTION_BRAND_PARK = pd.read_excel(links_main(fr"{script_dir}/file_links.txt", "conenection_brand"), sheet_name='PARK')
CONNECTION_BRAND_PLAN_AUTO = pd.read_excel(links_main(fr"{script_dir}/file_links.txt", "conenection_brand"), sheet_name='PLAN_AUTO')
PLAN_AUTO = pd.read_excel(links_main(fr"{script_dir}/file_links.txt", "plan_auto"), sheet_name='auto')
SERVER = links_main(fr"{script_dir}/file_links.txt", "server")
PORT = int(links_main(fr"{script_dir}/file_links.txt", "port"))
USER_NAME = links_main(fr"{script_dir}/file_links.txt", "username")
SEND_FROM = links_main(fr"{script_dir}/file_links.txt", "send_from")


df_main = pd.read_excel(links_main(fr"{script_dir}/file_links.txt", 'read_file_main'), sheet_name='Sheet1')
df_main['date_update'] = df_main['ссылка'].apply(lambda x: file_update(fr'{x}'))
df_main['date_update'] = df_main['date_update'].apply(lambda x: str(x).split(' ')[0])
df_main['date_update'] = pd.to_datetime(df_main['date_update'] )

for i in df_main.ключ.unique():
    print(i)
    try:
        print(df_main[df_main['ключ']==i])
    except Exception as ex_:
        print(f'{i} ошибка {ex_}')


def Shapka (table, text='VIN'):
    """ функция ищет в какой строке находится шапка таблицы, путём поиска "VIN" в ограниченной таблице table.iloc[0:15,0:15]
                    и вырезает лишние куски до найденной шапки и вместе с ней
                    если ошибка, то оставляем входную таблицу оставляем без именений

    Args:
        table (_type_): таблица
        text (str, optional): поиск по ключевому слову. Defaults to 'VIN'.

    Returns:
        table: таблица только с заголовками без верхних вспомогательных строк
    """
    
    try:
        table = table.T.reset_index().T.reset_index(drop=True) # опускает имена столбцов в первую строку
        qqq = table.iloc[0:20,0:20] # кординаты поиска 20х20
    
        # находим в какой строке находится шапка таблицы
        qqq = qqq[qqq.astype(str).apply(lambda x: x.astype(str).str.upper().str.contains(text, case=False)).any(axis=1)]
        q = qqq.index[0]
        table.columns = table.loc[q]
        table = table.iloc[q+1:,:]
        table = table.reset_index(drop = True)
        table = table[table.columns.dropna()]
        return(table)
    except:
        return(table)
    
    
def head_registr_low_strip(df):
    """переводит названия столбцов в нижний регистр удаляет пробелы слева и справа и добавляет _ ниж подч

    Args:
        df (_type_): _description_

    Returns:
        _type_: _description_
    """
    try:
        return df.rename(columns={f'{i}' : f'{str(i).lower().strip().replace(" ","_")}' for i in df.columns})
    except Exception as ex_:
                print(f'ошибка {ex_} функция - {head_registr_low_strip.__name__}')
                
                
def list_date_work(YEAR=2023, MONTH=1, DAY=1):
    """функция формируюящая список дат с и по
    """
    try:
        start_date = DT.datetime(YEAR, MONTH, DAY) # начальная дата
        cur =  (DT.date.today()-timedelta(days=1)).strftime('%Y-%m-%d').split('-') # текущая дата минус 1 день
        end_date = DT.datetime(int(cur[0]), int(cur[1]), int(cur[-1]))             # текущая дата минус 1 день

        res = pd.date_range(
            min(start_date, end_date),
            max(start_date, end_date)
        ).tolist() #.strftime('%d/%m/%Y').tolist()
        return res
    except Exception as ex_:
                print(f'ошибка {ex_} функция - {list_date_work.__name__}')
                
                
def read_datafarme(link):
    """считывает ссылку на книгу и получает имена всех листов и датафреймы
    пепеводит все имена листов в верхний регистр

    Args:
        link (_type_): сслыка на файл !!! ВАЖНО перед ссылкой fr 

    Returns:
        получаем все датафреймы и имена листов 
    """
    try:
        df_ = pd.read_excel(link, sheet_name=None)
        df_ = {key.upper(): value for key, value in df_.items()}  # привели названия листов в единый регистр
        df_names_lists = df_.keys()                               # получили все названия листов книги
        return df_, df_names_lists
    except Exception as ex_:
                print(f'ошибка {ex_} функция - {read_datafarme.__name__}')
                
                
def datetime_columns_convertor(df, name_colums=['дата']):
    """конвертирует столбцы с датами в формат даты

    Args:
        df (_type_): df для обработки
        name_colums (list, optional): имя которое будем искать в столбцах

    Returns:
        _type_: _description_
    """
    try:
        for i in name_colums:
            for j in df.columns:
                if i in j:
                    df[j] = pd.to_datetime(df[j], errors='coerce')

        return df
    except Exception as ex_:
                print(f'ошибка {ex_} функция - {datetime_columns_convertor.__name__}')


def numeric_columns_convertor(df, name_colums: list = ['внесено_в_рублях', 'получено_за_ам_руб', 'себестоимость', 'цена']):
    """конвертирует нужные столбцы в float

    Args:
        df (_type_): df для обработки
        name_colums (list, optional): имя которое будем искать в столбцах

    Returns:
        _type_: _description_
    """
    try:
        for i in name_colums:
            for j in df.columns:
                if i in j:
                    df[j] = pd.to_numeric(df[j], errors='coerce')

        return df
    except Exception as ex_:
                print(f'ошибка {ex_} функция - {numeric_columns_convertor.__name__} запнулось на {i, j}')
                
                
def zakazy(df, date_search, date_search_in_column, vid_zakaza:str, vid_zakaza_search_in_column:str):
    """считает количество заказов 

    Args:
        df (_type_): df по которому обрабатываем данные
        date_search (datetime64[ns]): входящая дата для поиска / 2025-03-01
        date_search_in_column (_type_): имя столца по которому фильтруем входящую дату
        vid_zakaza (_type_): вид заказа - кре/нал
        vid_zakaza_search_in_column (_type_):  имя столца по которому фильтруем вид заказа

    Returns:
        _type_: кол-во заказаов удовлетворяющее заданным параметрам
    """
    try:
        res = df[(df[date_search_in_column] == date_search) 
                & (df[vid_zakaza_search_in_column] == vid_zakaza)]['vin'].count()
        return res
    except Exception as ex_:
                print(f'ошибка {ex_} функция - {zakazy.__name__} входяные параметры {date_search, date_search_in_column, vid_zakaza, vid_zakaza_search_in_column}')
                
                
def zakazy_vid_oplaty(df, date_search, date_search_in_column, vid_zakaza:str, vid_zakaza_search_in_column:str, vid_oplaty_search_in_column:str, vid_oplaty:str):
    """считает количество заказов 

    Args:
        df (_type_): df по которому обрабатываем данные
        date_search (datetime64[ns]): входящая дата для поиска / 2025-03-01
        date_search_in_column (_type_): имя столца по которому фильтруем входящую дату
        vid_zakaza (_type_): вид заказа - кре/нал
        vid_zakaza_search_in_column (_type_):  имя столца по которому фильтруем вид заказа

    Returns:
        _type_: кол-во заказаов удовлетворяющее заданным параметрам
    """
    try:
        res = df[(df[date_search_in_column] == date_search) 
                & (df[vid_zakaza_search_in_column] == vid_zakaza)
                & (df[vid_oplaty_search_in_column] == vid_oplaty)]['vin'].count()
        return res
    except Exception as ex_:
                print(f'ошибка {ex_} функция - {zakazy_vid_oplaty.__name__} входяные параметры {date_search, date_search_in_column, vid_zakaza, vid_zakaza_search_in_column, vid_oplaty_search_in_column, vid_oplaty}')
                
                
def otkazy(df, date_search, date_search_in_column, vid_zakaza:str, vid_zakaza_search_in_column:str, ststus_arhiv:str, arhiv_search_in_column:str):
    """считает количество отказов 

    Args:
        df (_type_): df по которому обрабатываем данные
        date_search (datetime64[ns]): входящая дата для поиска / 2025-03-01
        date_search_in_column (_type_): имя столца по которому фильтруем входящую дату
        vid_zakaza (_type_): вид заказа - кре/нал
        vid_zakaza_search_in_column (_type_):  имя столца по которому фильтруем вид заказа
        status_zakaza (_type_): стутус заказа - на складе / в пути
        status_zakaza_search_in_column (_type_): имя столца по которому фильтруем стутус заказа
        stsus_arhiv:str - указывается да
        arhiv_search_in_column:str - столбец по которому фильтруется признак stsus_arhiv

    Returns:
        _type_: кол-во заказаов удовлетворяющее заданным параметрам
    """
    try:
        res = df[(df[date_search_in_column] == date_search) 
                & (df[vid_zakaza_search_in_column] == vid_zakaza)
                & (df[arhiv_search_in_column] == ststus_arhiv)]['vin'].count()
        return res
    except Exception as ex_:
                print(f'ошибка {ex_} функция - {otkazy.__name__} входяные параметры {date_search, date_search_in_column, vid_zakaza, vid_zakaza_search_in_column, ststus_arhiv, arhiv_search_in_column}')
                
                
def forma_pay(pay:str):
    """функция преобразования видов оплат в два парметра кре/нал


    Args:
        pay (_type_): поступающий аргумент

    Returns:
        _type_: _description_
    """
    kredit = ['кре', 'банк', 'лиз', 'finance', 'direct', 'втб', 'альфа'] # кре
    
    try:
        result = any([i in pay for i in kredit])
        if result == True:
            return 'кре'
        else:
            return 'нал'
    except Exception as ex_:
                print(f'ошибка {ex_} функция - {forma_pay.__name__} входяные параметры {pay}')
                
                
def kolichestyo_vidach(df, date_serch, column_date_serch, forma_oplaty, colmn_forma_pay_serch):
    """кол-во фактических выдач по форме оплаты кре/нал

    Args:
        df (_type_): df по которому производится поиск
        date_serch (_type_): искомая дата 
        column_date_serch (_type_): столбец с датами выдачи
        forma_pay (_type_): форма оплаты кре/нал
        colmn_forma_pay_serch (_type_): столбец с формой оплаты

    Returns:
        _type_: _description_
    """
    try:
        res = df[(df[column_date_serch]==date_serch) 
                            & (df[colmn_forma_pay_serch]==forma_oplaty)]['vin'].count()
        return res
    except Exception as ex_:
                print(f'ошибка {ex_} функция - {kolichestyo_vidach.__name__} входяные параметры {date_serch, column_date_serch, forma_oplaty, colmn_forma_pay_serch}')
                
                
def zakazy_s_vchetom_okazov_i_vidach(df, date_serch, column_date, column_sum_1, column_sum_2, column_sum_3):
    """кол-во заказаов с учетом отказов и выдач накопительно

    Args:
        df (_type_): df по которому идет поиск
        date_serch (_type_): дата поиска
        column_date (_type_): столбец с датами по которым идет поиск
        column_sum_1 (_type_): столбец с заказами (склад или в пути) кре/нал
        column_sum_2 (_type_): столбец с заказами (склад или в пути) кре/нал
        column_sum_3 (_type_): столбец с отказами (склад или в пути) кре/нал
        column_sum_4 (_type_): столбец с отказами (склад или в пути) кре/нал
        column_sum_5 (_type_): выдачи кре/нал

    Returns:
        _type_: _description_
    """
    try:
        res = (df[df[column_date]<=date_serch][column_sum_1].sum() 
        - df[df[column_date]<=date_serch][column_sum_2].sum()
        - df[df[column_date]<=date_serch][column_sum_3].sum())
        return res
    
    except Exception as ex_:
                print(f'ошибка {ex_} функция - {zakazy_s_vchetom_okazov_i_vidach.__name__} входяные параметры {date_serch, column_date, column_sum_1, column_sum_2, column_sum_3}')
                
                
def sum_finance_day(df, date_serch, column_date_serch, sum_column):
    """_summary_

    Args:
        df по которому идет фильтрация и поиск
        date_serch (_type_): дата по которой ищем
        column_date_serch (_type_): столбец с датой по которому идет фильтрация
        sum_column (_type_): столбец суммирования после фильтрации

    Returns:
        _type_: _description_
    """
    try:
        res = df[df[column_date_serch]==date_serch][sum_column].sum()
        return res
    except Exception as ex_:
                print(f'ошибка {ex_} функция - {sum_finance_day.__name__} входяные параметры {date_serch, column_date_serch, sum_column}')
                
                
def sum_finance_day_nakopitelno(df, date_serch, column_date, column_sum):
    """суммирование показателей накопительно с начала месяца

    Args:
        df (_type_): df по которому идет фильтрация
        date_serch (_type_): дата поиска
        column_date (_type_): столбец с датой по которому ищем
        column_sum (_type_): столбец по которому суммируем данные

    Returns:
        _type_: _description_
    """
    
    try:
        date_serch = pd.to_datetime(date_serch)
    except:
        print(f'ошибка с датой {date_serch} - не удалось преобразовать')
        return f'ошибка с датой {date_serch} - не удалось преобразовать'
    date_start = pd.to_datetime(f'{date_serch.year}-{date_serch.month}-01')
    res = df[(df[column_date]>=date_start) 
                          & (df[column_date]<=date_serch)][column_sum].sum()
    return res


def nacenka(df, date_serch, column_date_serch, colimn_sum_revenue, colimn_sum_cost_price):
    """рассчет наценки

    Args:
        df (_type_): по которому фильтрация
        date_serch (_type_): дата поиска
        column_date_serch (_type_): столбец по котрому фильтруем дату 
        colimn_sum_revenue (_type_): выручка
        colimn_sum_cost_price (_type_): себестоимость

    Returns:
        _type_: _description_
    """
    
    vir = df[df[column_date_serch]==date_serch][colimn_sum_revenue].sum()
    seb = df[df[column_date_serch]==date_serch][colimn_sum_cost_price].sum()
    try:
        res = (vir-seb)/seb if vir > 0 else 0
        return round(res, 2)
    
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {nacenka.__name__} входяные параметры {date_serch, column_date_serch, colimn_sum_revenue, colimn_sum_cost_price}')
        return 0
    
    
def prihod_auto(df, date_serch, column_date_na_sclad, column_date_zakaza, svobod_klient='svobod'or'klient'):
    """_summary_

    Args:
        df (_type_): фрейм по которму фильтруем
        date_serch (_type_): дата поиска
        column_date_na_sclad (_type_): столбцец даты прихода авто на склад
        column_date_zakaza (_type_): столббец даты заказа / контракта с клиентом
        svobod_klient (str, optional): _description_. Defaults to 'svobod'or'klient'.

    Returns:
        _type_: _description_
    """
    
    try:
        klient = df[(df[column_date_na_sclad]==date_serch) 
                        & (df[column_date_zakaza]<date_serch)]['vin'].count()
        svobod = df[(df[column_date_na_sclad]==date_serch)]['vin'].count()
        res = svobod-klient if svobod_klient =='svobod' else klient
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {prihod_auto.__name__} входяные параметры {date_serch, column_date_na_sclad, column_date_zakaza, svobod_klient}')
        return f'Проверьте дату {date_serch}'
    
    return res


def auto_na_sclade(df, date_serch, column_date_prihoda_na_sclad, column_date_realizacii, column_date_zakaza, pokazatel= 'all' or 'klient' or 'sclad' or 'demo'):
    """_summary_

    Args:
        df (_type_): фрейм по которому фильтруем
        date_serch (_type_): дата поиска
        column_date_prihoda_na_sclad (_type_): колонка с датой прихода авто на склад
        column_date_realizacii (_type_): колонка с датой продажи/реализации
        column_date_zakaza (_type_): колонка с датой заказа
        pokazatel (str, optional): _description_. Defaults to 'all'or'klient'or'sclad'.

    Returns:
        _type_: _description_
    """
    dem = ['DEMO', 'ДЕМО']
    try:
        all_ = df[(df[column_date_prihoda_na_sclad]<=date_serch) 
                    & ((df[column_date_realizacii].isna())
                            |(df[column_date_realizacii]==0)
                            |(df[column_date_realizacii]>date_serch))]['vin'].count()
        only_demo = df[(df[column_date_prihoda_na_sclad]<=date_serch)                                   # тоолько демо
                    & ((df[column_date_realizacii].isna())
                            |(df[column_date_realizacii]==0)
                            |(df[column_date_realizacii]>date_serch))
                    & (df['с_листа'].str.contains('|'.join(dem), case=False, na=False)) ]['vin'].count()  
        klient = df[(df[column_date_prihoda_na_sclad]<=date_serch) & ((df[column_date_realizacii].isna())
                            |(df[column_date_realizacii]==0)
                            |(df[column_date_realizacii]>date_serch)) & (df[column_date_zakaza]<=date_serch)]['vin'].count()
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {auto_na_sclade.__name__} входяные параметры {date_serch, column_date_prihoda_na_sclad, column_date_realizacii, column_date_zakaza, pokazatel}')
    
    if pokazatel == 'klient': res = klient
    elif pokazatel == 'sclad': res = all_- klient
    elif pokazatel == 'demo': res = only_demo
    else: res = all_
    
    return res


def korrekt_forma_oplaty(df, vin, text_forma_oplaty,column_date_realizcii, column_date_oplaty):
    """корреткирует форму оплаты в СКЛАД по NP

    Args:
        df (_type_): фрейм по которому идет фильтрация
        vin (_type_): vin для поиска
        text_forma_oplaty (_type_): входная форма оплаты для сравнения с источника откуда берем vin
        column_date_realizcii (_type_): столбец с датой фактической выдачи проверяем чтоб была заполнена
        column_date_oplaty (_type_): столбец с данными формы оплаты

    Returns:
        _type_: _description_
    """
    try:
        res = list(df[(df['vin']==vin) & ~(df[column_date_realizcii].isna())][column_date_oplaty])[0]
        if res==text_forma_oplaty: res = text_forma_oplaty
        elif res!=text_forma_oplaty: res = res
        else: res = text_forma_oplaty
    except:
        res = text_forma_oplaty
    return res


def all_letters():
    'все буквы ru_eng алфавита'
    
    import string
    try:
        eng = string.ascii_letters
        ru = ''.join(map(chr, range(ord('А'), ord('я')+1))) + 'Ёё'
        res = eng+ru
        return res
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {all_letters.__name__}')
        
        
def del_letters_date(word:str):
    """удаляет все буквы и порбелы лев прав

    Args:
        word (_type_): _description_

    Returns:
        _type_: _description_
    """
    try:
        word = str(word)
        for i in word:
            if i in all_letters():
                word=word.replace(i, '')
        return word.strip()
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {del_letters_date.__name__} входяные параметры {word}')
        
        
def shablon_date_test(date:str):
    """проверяет формат даты по шаблону - если есть дата, возвращает очищенное знаечние даты
    если нет - возвращает входящее значение без изменений

    Args:
        date (str): стркоа с датой

    Returns:
        _type_: _description_
    """
    import re
    match = re.search(r'(\d{4}.\d{2}.\d{2})', str(date)) # r'^(\d{4}.\d{2}.\d{2})$' - точное совпадение
    try:
        return match.group(1)
    except:
        return date
    
    
def shablon_date_test_2(d):
    import datetime
    # if d not in ['ok', 'nan', 'None', 'NaT', '-']:
    if '00:00:00' in d:
        d = d.split(' ')[0]

    if len(d.split('-')) == 3:
        try:
            datetime.datetime.strptime(d, '%Y-%m-%d')
            return 'ok' 
        except Exception:
            return None if d=='0' else d
    if len(d.split('.')) == 3:
        try:
            datetime.datetime.strptime(d, '%Y.%m.%d')
            return 'ok'
        except Exception:
            return None if d=='0' or 0 else d
        
        
def shablon_date_test_pravka(date:str):
    """ функция для сбора ошибок дат
    проверяет формат даты по шаблону - если есть дата, возвращает 'ok'
    если нет - возвращает входящее значение без изменений

    Args:
        date (str): стркоа с датой

    Returns:
        _type_: _description_
    """
    import re
    match = re.search(r'(\d{4}.\d{2}.\d{2})', str(date)) # r'^(\d{4}.\d{2}.\d{2})$' - точное совпадение
    try:
        match.group(0)
        if shablon_date_test_2(date)=='ok':
            return 'ok'
        else:
            return shablon_date_test_2(date)
    except:
        return None if date=='0' else date
    
    
def auto_na_sclade_consignacia(df, date_serch):
    """консигнационные авто на складе
    есть на складе но счет поним не оплачен нами

    Args:
        df (_type_): df по которому ищем
        date_serch (_type_): дата поиска

    Returns:
        _type_: _description_
    """
    try:
        res = df[(df['дата_прихода_на_склад']<=date_serch) & (df['дата_оплаты_счета'].isna())]['vin'].count()
        return int(res)
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {auto_na_sclade_consignacia.__name__} входяные параметры {date_serch}')
        None
        
        
def auto_u_puti_vikuplenie(df, date_serch):
    """авто в пути выкупленные
    есть дата оплаты счета но нет даты прихода на склад

    Args:
        df (_type_): df по которому ищем
        date_serch (_type_): дата поиска

    Returns:
        _type_: _description_
    """
    try:
        res = df[(df['дата_оплаты_счета']<=date_serch) & ((df['дата_прихода_на_склад'].isna()) | (df['дата_прихода_на_склад']>date_serch))]['vin'].count()
        return int(res)
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {auto_u_puti_vikuplenie.__name__} входяные параметры {date_serch}')
        None
        
        
def oplaty(df, date_serch):
    """считает сумму оплат на текущий день за авто

    Args:
        df (_type_): df по которому ведем поиск/фильтрацию
        date_serch (_type_): дата поиска

    Returns:
        _type_: _description_
    """
    try:
        res = df[(df['за_что']=='а/м') & (df['дата_оплаты']==date_serch)]['внесено_в_рублях'].sum()
        return res
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {oplaty.__name__} входяные параметры {date_serch}')
        None
        
        
def tek_day():
    try:
        import datetime
        current_date = datetime.date.today().isoformat()
        return current_date
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {tek_day.__name__}')
        
        
def df_oborotka_shablon(year:int, month:int, day:int):
    """щаблон для заполнения оборотки

    Args:
        year (int): год
        month (int): месяц
        day (int): число

    Returns:
        _type_: _description_
    """
    try:
        kalendar = list_date_work(year, month, day)
        df_oborotka = pd.DataFrame({'календарь': kalendar, 
                                    # 'зкз_скл_кред': [0 for i in range(len(kalendar))],
                                    # 'зкз_скл_нал': [0 for i in range(len(kalendar))],
                                    # 'зкз_путь_кред': [0 for i in range(len(kalendar))],
                                    # 'зкз_путь_нал': [0 for i in range(len(kalendar))],
                                    # 'откз_скл_кред': [0 for i in range(len(kalendar))],
                                    # 'откз_скл_нал': [0 for i in range(len(kalendar))],
                                    # 'откз_путь_кред': [0 for i in range(len(kalendar))],
                                    # 'откз_путь_нал': [0 for i in range(len(kalendar))],
                                    'зкз_путь_кред': [0 for i in range(len(kalendar))], # в процессе
                                    'зкз_склад_кред': [0 for i in range(len(kalendar))],# в процессе
                                    'зкз_путь_нал': [0 for i in range(len(kalendar))], # в процессе
                                    'зкз_склад_нал': [0 for i in range(len(kalendar))],# в процессе
                                    
                                    'зкз_кред': [0 for i in range(len(kalendar))],
                                    'зкз_нал': [0 for i in range(len(kalendar))],
                                    'откз_кред': [0 for i in range(len(kalendar))],
                                    'откз_нал': [0 for i in range(len(kalendar))],
                                    
                                    'всего_зкз_с_уч_откз_и_выд_кред': [0 for i in range(len(kalendar))],
                                    'всего_зкз_с_уч_откз_и_выд_нал': [0 for i in range(len(kalendar))],
                                    'всего_зкз_с_уч_откз_и_выд_всего': [0 for i in range(len(kalendar))],
                                    'выдачи_кред': [0 for i in range(len(kalendar))],
                                    'выдачи_нал': [0 for i in range(len(kalendar))],
                                    'выдачи_всего': [0 for i in range(len(kalendar))],
                                    'выдачи_выручка': [0 for i in range(len(kalendar))],
                                    'выдачи_себестоимость': [0 for i in range(len(kalendar))],
                                    'продано_ам_накоп': [0 for i in range(len(kalendar))],
                                    'выручка_накоп': [0 for i in range(len(kalendar))],
                                    'себестоимость_накоп': [0 for i in range(len(kalendar))],
                                    'наценка': [0 for i in range(len(kalendar))],
                                    'приход_ам_своб': [0 for i in range(len(kalendar))],
                                    'приход_ам_клиент': [0 for i in range(len(kalendar))],
                                    'ам_на_складе_своб': [0 for i in range(len(kalendar))],
                                    'ам_на_складе_клиент': [0 for i in range(len(kalendar))],
                                    'склад_всего_ам': [0 for i in range(len(kalendar))],
                                    'склад_в_тч_демо_ам': [0 for i in range(len(kalendar))],
                                    'склад_конс_ам': [0 for i in range(len(kalendar))],
                                    # 'ам_в_пути_конс': [0 for i in range(len(kalendar))],
                                    'ам_в_пути_выкуп': [0 for i in range(len(kalendar))],
                                    # 'ам_в_пути_всего': [0 for i in range(len(kalendar))],
                                    'оплаты': [0 for i in range(len(kalendar))],
                                    'платежи_ам_клиент_шт': [0 for i in range(len(kalendar))],
                                    'платежи_ам_клиент_руб': [0 for i in range(len(kalendar))],
                                    'платежи_ам_свободн_шт': [0 for i in range(len(kalendar))],
                                    'платежи_ам_свободн_руб': [0 for i in range(len(kalendar))],
                                    'платежи_ам_всего_шт': [0 for i in range(len(kalendar))],
                                    'платежи_ам_всего_руб': [0 for i in range(len(kalendar))],
                                    #'остаток_денеж_средств': [0 for i in range(len(kalendar))],
                                    'оборот_средства_без_демо': [0 for i in range(len(kalendar))],
                                    'оборот_средства_без_демо_на_скл': [0 for i in range(len(kalendar))],
                                    'оборот_средства_демо': [0 for i in range(len(kalendar))],
                                    'проверка': [0 for i in range(len(kalendar))],
                                    })
        return df_oborotka
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {df_oborotka_shablon.__name__} входяные параметры {year, month, day} не удалось создать календарь')
        
        
def min_date_column(df, column):
    """ считывает MIN дату по столбцу"""
    try:
        res = df[df[column].notna()][column].min()
        return res
    except:
        None
          
        
def min_date_test(df):
    """поиск минимального значения дат
    используется для передачи данных функции заполнения шаблона df_oborotka_shablon
    
    Args:
        df (_type_): _description_

    Returns:
        _type_: выаодит год, месяц, число
    """
    try:
        res = []
        res.append(min_date_column(df.df_sclad, 'дата_оплаты_счета'))
        res.append(min_date_column(df.df_sclad, 'дата_прихода_на_склад'))
        res.append(min_date_column(df.df_sclad, 'дата_контракта_заказа'))
        res.append(min_date_column(df.df_np_auto, 'дата_заказа'))
        res.append(min_date_column(df.df_np_oplata, 'дата_оплаты'))
        res = min(res)
        
        # возникают ошибки ввода 1900 год, тогда берем максимально дупустимую дату в ручном режиме 2005 год,
        # что займет около 5 минут времени на заполнение шаблона, все лучше бесконечности
        return 2005 if int(res.year)<2005 else int(res.year), int(res.month), int(res.day)
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {min_date_test.__name__}')
        return 2025, 1, 1
    
    
def min_year_date_column(df, column):
    """считывает минимальный год по столбцу

    Args:
        df (_type_): _description_
        column (_type_): _description_

    Returns:
        _type_: _description_
    """
    
    try:
        res = df[df[column].notna()][column].min()
        return res.year
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {min_year_date_column.__name__} входяные параметры {column}')
        
        
def mean_year_date_column(df, column):
    """считывает средний год по столбцу

    Args:
        df (_type_): _description_
        column (_type_): _description_

    Returns:
        _type_: _description_
    """
    try:
        res = df[df[column].notna()][column].mean()
        return res.year
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {mean_year_date_column.__name__} входяные параметры {column}')
        
        
def platejy(df, date_serch, sum_or_count: str = 'sum_' or 'count_', pokazatel:str = 'klient' or 'sclad'):
    """_summary_

    Args:
        df (_type_): df по которому идет фильтрация
        date_serch (_type_): дата поиска
        sum_or_count (str, optional): _description_. Defaults to 'sum_'or'count_'.
        pokazatel (str, optional): _description_. Defaults to 'klient'or'sclad'.

    Returns:
        _type_: _description_
    """
    try:
        if pokazatel=='klient':
            if sum_or_count == 'count_':
                res = df[(df['дата_оплаты_счета']==date_serch) & (df['дата_контракта_заказа']<date_serch)]['себестоимость_ам'].count()
                return res
            elif sum_or_count == 'sum_':
                res = df[(df['дата_оплаты_счета']==date_serch) & (df['дата_контракта_заказа']<date_serch)]['себестоимость_ам'].sum()
                return res
            else:
                print(f"неизвестный показатель [{sum_or_count}] допустимые значения 'sum_' or 'count_'")
                

        elif pokazatel=='sclad':
            if sum_or_count == 'count_':
                res = df[(df['дата_оплаты_счета']==date_serch)]['себестоимость_ам'].count()
                return res
            elif sum_or_count == 'sum_':
                res = df[(df['дата_оплаты_счета']==date_serch)]['себестоимость_ам'].sum()
                return res
            else:
                print(f"неизвестный показатель [{sum_or_count}] допустимые значения 'sum_' or 'count_'")
                
        else:
            print(f"неизвестный показатель [{pokazatel}] допустимые значения 'klient' or 'sclad'")
    
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {platejy.__name__} входяные параметры {date_serch, sum_or_count, pokazatel}')
        
        
def unique_name_list_demo():
    """сбор уник имен листов с демо - используется в Manufacturing_df_oborotka

    Returns:
        _type_: _description_
    """
    try:
        names_list = set()
        for i in catalog_df_predobrabotka.keys(): # !!! данные появляются в Manufacturing_df_oborotka
            # собираем уникальные названия из объектов класса предобратки sclad с именами листов
            names_list.update(catalog_df_predobrabotka[i].df_sclad.с_листа.unique())

        filtr_ = ['ТЕСТ', 'DEM', 'ДЕМ'] # фильтры для отсеивания демо авто
        filtr_demo_names_list = []      # названий листов имеющих отношение к демо
        for i in names_list:
            if any([p.upper() in i.upper() for p in filtr_ ]):
                filtr_demo_names_list.append(i)
        return filtr_demo_names_list
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {unique_name_list_demo.__name__} входяные параметры')
        
        
def oborotnie_sredstya(df, date_serch, result:str = 'not_demo' or 'not_demo_na_sclade' or 'demo'):
    """оборотные средства - возвращает результат в зависимости от запрашиваемого типа данных

    Args:
        df (_type_): df по которому фильтруем
        date_serch (_type_): дата поиска 
        result (str, optional): _description_. Defaults to 'not_demo' or 'not_demo_na_sclade' or 'demo'.

    Returns:
        _type_: _description_
    """
    try:
        if result == 'not_demo':
            res1 = df[(~df['с_листа'].str.contains('|'.join(unique_name_list_demo()))) & 
                      (df['дата_оплаты_счета']<=date_serch) & 
                      (df['дата_продажи_факт']>date_serch)]['себестоимость_ам'].sum()
            res2 = df[(~df['с_листа'].str.contains('|'.join(unique_name_list_demo()))) & 
                    (df['дата_оплаты_счета']<=date_serch) & 
                    (~df['дата_продажи_факт'].notna())]['себестоимость_ам'].sum()
            return res1+res2

        elif result == 'not_demo_na_sclade':
            res1 = df[(~df['с_листа'].str.contains('|'.join(unique_name_list_demo()))) & 
                      (df['дата_оплаты_счета']<=date_serch) & 
                      (df['дата_продажи_факт']>date_serch)& (~df['дата_продажи_факт'].notna())]['себестоимость_ам'].sum()
            res2 = df[(~df['с_листа'].str.contains('|'.join(unique_name_list_demo()))) & 
                    (df['дата_оплаты_счета']<=date_serch) & 
                    (~df['дата_продажи_факт'].notna()) & (~df['дата_продажи_факт'].notna())]['себестоимость_ам'].sum()
            return res1+res2            
            
        elif result == 'demo':
            res1 = df[(df['с_листа'].str.contains('|'.join(unique_name_list_demo()))) & 
                      (df['дата_оплаты_счета']<=date_serch) & 
                      (df['дата_продажи_факт']>date_serch)]['себестоимость_ам'].sum()
            res2 = df[(df['с_листа'].str.contains('|'.join(unique_name_list_demo()))) & 
                    (df['дата_оплаты_счета']<=date_serch) & 
                    (~df['дата_продажи_факт'].notna())]['себестоимость_ам'].sum()
            return res1+res2

        else:
            print(f'!!! проверьте поданный параметр "{result}"')
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {oborotnie_sredstya.__name__} входяные параметры {date_serch, result}')
        
        
def proverka_oborotnih_sredsty(df, date_serch):
    """для проверки оборотных средств

    Args:
        df (_type_): df по которому фильтруем
        date_serch (_type_): дата поиска

    Returns:
        _type_: _description_
    """
    try:
        res1 = df[(~df['с_листа'].str.contains('|'.join(unique_name_list_demo()))) & 
                (df['дата_оплаты_счета']<=date_serch) & 
                (df['дата_продажи_факт']>date_serch)]['себестоимость_ам'].sum()
        res2 = df[(~df['с_листа'].str.contains('|'.join(unique_name_list_demo()))) & 
                (df['дата_оплаты_счета']<=date_serch) & 
                (~df['дата_продажи_факт'].notna())]['себестоимость_ам'].sum()
        return res1+res2
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {proverka_oborotnih_sredsty.__name__} входяные параметры {date_serch}')
        
        
def status_zakaza_VARSH_BAIK_UKA_HYUNDAI(znach):
    """индивидуально только для _VARSH_BAIK_UKA_HYUNDAI так как у них нет NP
    и статусы склада свои, подгоняем под общий стандарт

    Args:
        znach (_type_): _description_

    Returns:
        _type_: _description_
    """
    try:
        znach = str(znach).lower().strip()
        if znach == 'москва':
            return 'на складе'
        elif znach == 'выдан':
            return 'на складе'
        else:
            return 'в пути'
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {status_zakaza_VARSH_BAIK_UKA_HYUNDAI.__name__} входяные параметры {znach}')
        
        
def pravka_statysa_KIA_(word):
    """индивидуальная правка статуса склада np

    Args:
        word (_type_): _description_

    Returns:
        _type_: _description_
    """
    try:
        if str(word).strip().lower() == 'demo':
            return 'на складе'
        elif 'кмр' in str(word).strip().lower():
            return 'в пути'
        elif str(word).strip().lower() == 'склад':
            return 'на складе'
        elif 'овп' in str(word).strip().lower():
            return 'на складе'
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {pravka_statysa_KIA_.__name__} входяные параметры {word}')
        
        
def kostraciya(df_np, df_sklad, df_oplata, date_kastracii):
    """функция обрезания данных NP СКЛАДА и ОПЛАТ

    Args:
        df_sklad (_type_): склад
        df_np (_type_): np
        date_kistracii (_type_): дата '2023-01-01'
        new_np, new_sklad = kostraciya(catalog_df_predobrabotka['HYUNDAI_YAR'].df_sclad, catalog_df_predobrabotka['HYUNDAI_YAR'].df_np_auto, '2023-01-01')
    Returns:
        _type_: 2 df
    """
    try:
        sklad_do_obrezaniya = df_sklad[df_sklad['дата_прихода_на_склад']<date_kastracii]            # кострируем склад
        sklad_do_obrezaniya = sklad_do_obrezaniya.groupby(['vin','цена_продажи', 'дата_прихода_на_склад']).count().reset_index()[['vin','цена_продажи', 'дата_прихода_на_склад']] # группуируем
        sklad_do_obrezaniya = sklad_do_obrezaniya[['vin', 'дата_прихода_на_склад','цена_продажи']]
        sklad_do_obrezaniya['q'] = 1
        inner_joined_df = pd.merge(df_np, sklad_do_obrezaniya, left_on=['vin', 'получено_за_ам_руб'], right_on=['vin', 'цена_продажи'], how='outer', suffixes=('','_y'))
        inner_joined_df = inner_joined_df[inner_joined_df['q'].isna()]
        lst_del_col = ['дата_прихода_на_склад_y', 'цена_продажи', 'q']
        new_np = inner_joined_df[[i for i in inner_joined_df.columns if i not in lst_del_col]]
        
        new_np = new_np[(new_np['дата_изм']>=date_kastracii)|(new_np['дата_изм'].isna()) ]
        
        new_np = new_np[((new_np['дата_прихода_на_склад']>=date_kastracii) | (new_np['дата_прихода_на_склад'].isna()))
                                        & (new_np['дата_заказа']>=date_kastracii) | (new_np['дата_заказа'].isna())
                                        & ((new_np['дата_полной_оплаты_факт']>=date_kastracii) | (new_np['дата_полной_оплаты_факт'].isna())) 
                                        & ((new_np['дата_выдачи_факт']>=date_kastracii) | (new_np['дата_выдачи_факт'].isna()))]
        
        
        df_oplata = df_oplata[(df_oplata['дата_оплаты']>=date_kastracii)]
        
        
        sklad_posle_obrezaniya = df_sklad[((df_sklad['дата_прихода_на_склад']>=date_kastracii) | (df_sklad['дата_прихода_на_склад'].isna()))
                                        & ((df_sklad['дата_контракта_заказа']>=date_kastracii) | (df_sklad['дата_контракта_заказа'].isna()))
                                        & ((df_sklad['дата_оплаты_счета']>=date_kastracii) | (df_sklad['дата_оплаты_счета'].isna()))
                                        & ((df_sklad['дата_продажи_факт']>=date_kastracii) | (df_sklad['дата_продажи_факт'].isna()))]
        
        return new_np, sklad_posle_obrezaniya, df_oplata
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {kostraciya.__name__} входяные параметры {date_kastracii}')
        
        
def kostraciya_2(df_np, df_sklad, df_oplata, date_kastracii):
    """функция обрезания данных NP СКЛАДА все тож как и в kostraciya 
    только берем авто пришедшее на склад до даты кострации и у которых дата продажи 2025 год
    это могут быть демо авто или подменные

    Args:
        df_sklad (_type_): склад
        df_np (_type_): np
        date_kistracii (_type_): дата '2023-01-01'
        new_np, new_sklad = kostraciya(catalog_df_predobrabotka['HYUNDAI_YAR'].df_sclad, catalog_df_predobrabotka['HYUNDAI_YAR'].df_np_auto, '2023-01-01')
    Returns:
        _type_: 2 df
    """
    try:
        sklad_do_obrezaniya = df_sklad[(df_sklad['дата_прихода_на_склад']<date_kastracii) 
                           & (df_sklad['дата_продажи_факт']>='2025-01-01')]
        
        # sklad_do_obrezaniya = df_sklad[df_sklad['дата_прихода_на_склад']<date_kastracii]            # кострируем склад
        sklad_do_obrezaniya = sklad_do_obrezaniya.groupby(['vin','цена_продажи', 'дата_прихода_на_склад']).count().reset_index()[['vin','цена_продажи', 'дата_прихода_на_склад']] # группуируем
        sklad_do_obrezaniya = sklad_do_obrezaniya[['vin', 'дата_прихода_на_склад','цена_продажи']]
        sklad_do_obrezaniya['q'] = 1
        inner_joined_df = pd.merge(df_np, sklad_do_obrezaniya, left_on=['vin', 'получено_за_ам_руб'], right_on=['vin', 'цена_продажи'], how='outer', suffixes=('','_y'))
        inner_joined_df = inner_joined_df[inner_joined_df['q'].isna()]
        lst_del_col = ['дата_прихода_на_склад_y', 'цена_продажи', 'q']
        new_np = inner_joined_df[[i for i in inner_joined_df.columns if i not in lst_del_col]]
        
        new_np = new_np[(new_np['дата_изм']>=date_kastracii)|(new_np['дата_изм'].isna()) ]
        
        new_np = new_np[((new_np['дата_прихода_на_склад']<date_kastracii))
                                        # & (new_np['дата_заказа']>=date_kastracii) | (new_np['дата_заказа'].isna())
                                        # & ((new_np['дата_полной_оплаты_факт']>=date_kastracii) | (new_np['дата_полной_оплаты_факт'].isna())) 
                                        & ((new_np['дата_выдачи_факт']>='2025-01-01') )]
        
        
        # df_oplata = df_oplata[(df_oplata['дата_оплаты']>=date_kastracii)]
        
        
        sklad_posle_obrezaniya = df_sklad[(df_sklad['дата_прихода_на_склад']<date_kastracii) 
                           & (df_sklad['дата_продажи_факт']>='2025-01-01')]
        
        
        return new_np, sklad_posle_obrezaniya
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {kostraciya.__name__} входяные параметры {date_kastracii}')
        
        
def status_zakaza_po_date(date_zakaza, date_prih):
    """определяет логистический статус авто 
    по дате заказа и прихода

    Args:
        date_zakaza (_type_): _description_
        date_prih (_type_): _description_

    Returns:
        _type_: _description_
    """
    try:
        if date_zakaza<date_prih:
            return 'в пути'
        else:
            return 'на складе'
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {status_zakaza_po_date.__name__} входяные параметры {date_zakaza, date_prih}')
        
        
def one_unique_pokazatel(df, vin, column):
    """возвращает резуьтат искомого данного по vin

    Args:
        df (_type_): вdf
        vin (_type_): vin
        column (_type_): столбец откуда получаем результат

    Returns:
        _type_: _description_
    """
    try:
        res = list(df[df['vin']==vin][column])[0]
        return res 
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {one_unique_pokazatel.__name__} входяные параметры {vin, column}')
        
        
def convertor_brands_in_PARK(marka, region):
    """ подается марка и регион из оборотки
    возвращает конвертированные данные для поиска бонусов в ПАРКЕ

    Args:
        marka (_type_): марка из оборотки
        region (_type_): регион из оборотки

    Returns:
        _type_: повзращает пару для поиска в ПАРКЕ
    """
    try:
 
        mark = list(CONNECTION_BRAND_PARK[(CONNECTION_BRAND_PARK['марка_фильтр']==marka) 
                                     & (CONNECTION_BRAND_PARK['регион_фильтр']==region)]['марка'])
        reg = list(CONNECTION_BRAND_PARK[(CONNECTION_BRAND_PARK['марка_фильтр']==marka) 
                                    & (CONNECTION_BRAND_PARK['регион_фильтр']==region)]['подразделение'])
        return mark, reg
 
    except Exception as ex_:
        print(f'ошибка {ex_} функция - {convertor_brands_in_PARK.__name__} входяные параметры {marka, region}')
        
        
def bonus_park(date_serch, list_marka:list, list_reg:list, result_column):
    """вытаскивает данные из ПАРКА

    Args:
        date_serch (_type_): дата поиска
        list_marka (list): список марок
        list_reg (list): список регионов
        result_column (_type_): по какому столбцу ищем

    Returns:
        _type_: _description_
    """
    try:
        if len(list_marka)!=0 and len(list_reg)!=0:
            res = PARK[(PARK['мес']==date_serch) 
                & (PARK['Марка'].str.contains("|".join(list_marka))) 
                & (PARK['Подразделение/площадка'].str.contains("|".join(list_reg))) 
                & (PARK['ТИП']=='Бонус')][result_column].sum()
            return res
        else:
            return 0
    except Exception as ex_:
        print(f'ошибка функции {bonus_park.__name__} {ex_} не удалось определеить входные параметры {date_serch, list_marka, list_reg, result_column}')
        
        
def individ_date_plan(year, month):
    try:
        year = str(year)
        month = str(month)
        month = month if len(month)==2 else '0'+month
        day = '01'
        return f'{year}-{month}-{day}'
    except Exception as ex_:
        print(f'Ошибка функции {individ_date_plan.__name__} {ex_} не удалось преобразовать {year}{month}')
        
        
def result_date_update(key, serch_column):
    'вовращае дату обновления файла'
    try:
        res = df_main[df_main['ключ']==key][serch_column].max()
        return res
    except Exception as ex_:
        print(f'ошибка функции {result_date_update.__name__} не удалось определеить входные параметры {key, serch_column}')
        None
        
        
def read_file_arhiv(name_objekt, name_list:str='oborotka', name_file_link:str = "copy_link_dir"):
    """функция получает имя объекта и считывает готовые данные из файла
    в определенной директории ища это имя объекта

    Args:
        name_objekt (_type_): имя объекта
        name_list - имя листа
        name_file_link - имя пути в файле 
    Returns:
        _type_: _description_
    """
    try:
        directory = links_main(fr"{script_dir}/file_links.txt", name_file_link)               # директория где ищем
        pattern = name_objekt                                                   # что ищем
        files = [i for i in os.listdir(directory) if pattern in i][0]           # результат
        df = pd.read_excel(fr'{directory}\{files}', sheet_name=name_list)      # считываем данные
        df = df[[i for i in df.columns if i not in 'Unnamed: 0']]
        return df
    except Exception as ex_:
        print(f'ошибка функции {read_file_arhiv.__name__} {ex_} не удалось найти объект {name_objekt} в дирректории{directory} или произошла ошибка чтения файла')
        
        
def return_link_directory(pattern, name_link):
    """_summary_

    Args:
        pattern (_type_): искомый объект в дирректории
        name_link (_type_): имя ключа на ссылку

    Returns:
        _type_: _description_
    """
    import os
    
    try:
        directory = links_main(fr"{script_dir}/file_links.txt", name_link)               # директория где ищем
        pattern = pattern                                                   # что ищем
        files = [i for i in os.listdir(directory) if pattern in i][0]           # результат
        link = fr'{directory}\{files}'    # считываем данные
        aktiv = os.path.isfile(link)
        if aktiv:
            return link
        else:
            print(f'ссфлка {link} не активна {aktiv}')
    except Exception as ex_:
        print(f'ошибка функции {return_link_directory.__name__} {ex_} входне параметры {pattern, name_link}')
        
        
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
        print(f'ошибка функции {return_link_directory.__name__}  {ex_}')
        
        
def sravnenie_arh_skl_k_tek(vin, df_arh, sf_tek):
    try:
        result = []
        for i in df_arh.columns:
            arh = list(df_arh[df_arh['vin']==vin][i])[0]
            tek = list(sf_tek[sf_tek['vin']==vin][i])[0]
            # print(arh, type(arh), tek, type(tek))
            if isinstance(tek, (int, float)):
                if float(arh)!=float(tek):
                    result.append(f'столбец {i} / было {arh} стало {tek}')
            elif isinstance(tek, (str)):
                if str(arh)!=str(tek):
                    result.append(f'столбец {i} / было {arh} стало {tek}')
            elif arh!=tek:
                result.append(f'столбец {i} / было {arh} стало {tek}')
                
        return result
    except Exception as ex_:
        if len(list(sf_tek[sf_tek['vin']==vin]))>=1:
            return f'vin {vin} есть но данных нет'
        else:
            return f'не удалось найти {vin}'
        
        
def update_file(link):
    """обновление сводной таблицы Excel
    # блок импортов для обновления сводных
    import pythoncom
    pythoncom.CoInitializeEx(0)
    import win32com.client
    Args:
        link (_type_): ссылка на файл - который нужно обновить
    """
    try:
        xlapp = win32com.client.DispatchEx("Excel.Application")
        wb = xlapp.Workbooks.Open(link)
        wb.Application.AskToUpdateLinks = False   # разрешает автоматическое  обновление связей (файл - парметры - дополнительно - общие - убирает галку запрашивать об обновлениях связей)
        wb.Application.DisplayAlerts = True  # отображает панель обновления иногда из-за перекрестного открытия предлагает ручной выбор обновления True - показать панель
        wb.RefreshAll()
        #xlapp.CalculateUntilAsyncQueriesDone() # удержит программу и дождется завершения обновления. было прописано time.sleep(30)
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
        print(f'ошибка функции {update_file.__name__} {ex_} не удалось обновить файл по ссылке {link}')
        
        
def return_email_except_df(obj, name_serch_col, name_col):
    """возыращает список данных из df

    Args:
        obj (_type_): искомый объект
        name_serch_col (_type_): имя столбца по которому идет фильтрация
        name_col (_type_): имя результирующего столбца по которому вернется ответ

    Returns:
        _type_: list
    """
    try:
        res = list(df_emal_exception[df_emal_exception[name_serch_col]==obj][name_col])[0]
        result = [i.strip() for i in res.split(';')] 
        return result
    except Exception as ex_:
        print(f'ошибка функции {return_email_except_df.__name__} значение для поиска {obj} поиск по колонке {name_serch_col} реузльтат по колонке {name_col} ошибка {ex_}')
        
        
def my_pass():
    """функция считывания пароля

    Returns:
        _type_: _description_
    """
    
    try:
        with open(links_main(fr"{script_dir}/file_links.txt", "pass_link"), 'r') as actual_pass:
            return actual_pass.read()
        
    except Exception as ex_:
        print(f'ошибка функции {my_pass.__name__} {ex_}')
        
        
def send_mail(send_to:list, file_link, file_name):
    """рассылка почты

    Args:
        send_to (list): список адресов для рассылки
        file_link(str): ссылка на файл
        file_name(str): имя файла в данном варианте нужно указывать с расширением 'BAIC_MSK.xlsx' 
    """
    from datetime import datetime, date, timedelta
    
    try:
        send_from = SEND_FROM                                                               
        subject = f"Проверка {file_name.split('_')[0]} NP и SCLAD на {(datetime.now()-timedelta(1)).strftime('%d-%m-%Y')}"                                                                 
        text = f"Здравствуйте\nВо вложении результат проверки NP и SCLAD на {(datetime.now()- timedelta(1)).strftime('%d-%m-%Y')}"                                                                   
        files = fr'{file_link.strip()}'
        server = SERVER
        port = PORT
        username=USER_NAME
        password = my_pass()
        isTls=True
        
        msg = MIMEMultipart()
        msg['From'] = send_from
        msg['To'] = ','.join(send_to)
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
        smtp.sendmail(send_from, send_to, msg.as_string())
        smtp.quit()
        
    except Exception as ex_:
        print(f'ошибка функции {send_mail.__name__} {ex_} входне параметры {send_to} {file_link} {file_name}')
        
        
def send_mail_2(send_to:list, file_link, file_name, them = '', body=''):
    """рассылка почты

    Args:
        send_to (list): список адресов для рассылки
        file_link(str): ссылка на файл
        file_name(str): имя файла в данном варианте нужно указывать с расширением 'BAIC_MSK.xlsx' 
        them(str) - тема письма
        
    """
    from datetime import datetime, date, timedelta
    
    try:
        send_from = SEND_FROM                                                               
        subject = f"{them} на {(datetime.now()-timedelta(1)).strftime('%d-%m-%Y')}"                                                                 
        text = f"Здравствуйте\n{body} на {(datetime.now()- timedelta(1)).strftime('%d-%m-%Y')}"                                                                   
        files = fr'{file_link.strip()}'
        server = SERVER
        port = PORT
        username=USER_NAME
        password = my_pass()
        isTls=True
        
        msg = MIMEMultipart()
        msg['From'] = send_from
        msg['To'] = ','.join(send_to)
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
        smtp.sendmail(send_from, send_to, msg.as_string())
        smtp.quit()
        
    except Exception as ex_:
        print(f'ошибка функции {send_mail.__name__} {ex_} входне параметры {send_to, file_link, file_name, them, body}')
        
        
def raznica_date_arhiv(df, serch_columns = 'календарь'):
    """находит разницу дней между датой на вчера и максимальной 
    датой в df в указанном столбце с датой (берет максимальноую)
    написана спец для работы с листом оборотки

    Args:
        df (_type_): _description_
        serch_columns - имя столбца с датой по которому идет поиск

    Returns:
        _type_: _description_
    """
    import datetime
    try:
        list_date = []
        res = df[serch_columns].max() # максимальная дата в архиве
        real_date = yesterday()                              # актуальная дата
        result_day = real_date - res                         # находим разницу между актуальной датой и максимальной в архиве
        result_day = result_day.days                         # кол-во дней
        for i in range(result_day):                          # собираем список пропущенных дат
            end_date = res + datetime.timedelta(days=1)
            res = end_date
            list_date.append(end_date)
        
        return list_date if len (list_date)>0 else None
    except Exception as ex_:
        print(f'ошибка функции {send_mail.__name__} ошибка {ex_}')
        
        
def protajka_stolbcov_v_arhivnoy_oborotke(df):
    """функция протяжки строк / столбцов в архивной оборотке
    добавляются даты сниз фреймв, протягиваются нужные (накопительные) столбцы / строки
    NAN заменяются на 0

    Args:
        df (_type_): df

    Returns:
        _type_: df
    """
    try:
        df = df
        x_test = pd.DataFrame(raznica_date_arhiv(df), columns=['календарь']) # создаем df с недостающими датами
        df = pd.concat([df,x_test])                                          # конкатинируем его к основному df
        df.reset_index(inplace=True, drop=True)                              # сбрасываем индексы
        # стлбцы которые нужно протянуть 
        columns_nakopitelno_ffill = ['всего_зкз_с_уч_откз_и_выд_кред','всего_зкз_с_уч_откз_и_выд_нал','всего_зкз_с_уч_откз_и_выд_всего',
                                     'выручка_накоп', 'себестоимость_накоп', 'наценка', 'ам_на_складе_своб', 'ам_на_складе_клиент', 
                                'склад_всего_ам', 'склад_конс_ам', 'ам_в_пути_выкуп', 'оборот_средства_без_демо', 'оборот_средства_без_демо_на_скл', 
                                'оборот_средства_демо', 'имя_объекта', 'марка', 'марка', 'регион']
        # протягиваем столбцы в df
        for i in columns_nakopitelno_ffill:
            df[i] = df[i].ffill()
            
        df = df.fillna(0)                                                    # заменяем NAN на 0
        return df
    except Exception as ex_:
        print(f'ошибка функции {protajka_stolbcov_v_arhivnoy_oborotke.__name__} ошибка {ex_}')
        
        
def exception_result_korrekt(df, column_serch):
    """функция для обработки таблицы с ошибками
    убирает ошибки на листах в работе, склад - так как они рабочие

    Args:
        df (_type_): фрейм который обрабатываем
        column_serch (_type_): столбец

    Returns:
        _type_: _description_
    """
    try:
        iskl = ['СКЛАД', 'Склад', 'В РАБОТЕ']
        df = df[~df[column_serch].str.contains('|'.join(iskl), na=False, case=False)]
        return df
    except Exception as ex_:
        print(f'Ошибка функции {exception_result_korrekt.__name__} {ex_}')
        
        
def arhivirovanie(link_directory_copy, link_directory_paste, pattern=''):
    """для копирования файлов из одной директории в другую
    import os
    import shutil
    Args:
        link_directory_copy (_type_): директория откуда копируем
        link_directory_paste (_type_): директория куда вставляем
        pattern (str, optional): _description_. Defaults to ''. - паттерн для фильтрования определенных файлов '.txt' '.xls'
    """
    try:
        directory_copy = link_directory_copy
        directory_paste = link_directory_paste
        # Получаем список файлов
        files = [i for i in os.listdir(directory_copy) if i.endswith(pattern)]
        print(f'Архивироание складов для сравнения')
        for i in files:
            print(f'{directory_copy}\{i} в {directory_paste}\{i}')
            shutil.copy2(fr'{directory_copy}\{i}', fr'{directory_paste}\{i}')
            LOG_inf(f'{directory_copy}\{i} в {directory_paste}\{i}', 'INFO')
    except Exception as ex_:
        print(f'ошибка функции {arhivirovanie.__name__} {ex_} входные данные {link_directory_copy, link_directory_paste, pattern}')
        LOG_inf(f'ошибка функции {arhivirovanie.__name__}', 'ERROR', ex_)
    

arhivirovanie(links_main(fr"{script_dir}/file_links.txt", "copy_link_dir"), links_main(fr"{script_dir}/file_links.txt", "paste_link_dir"))


# df с ключами поиска используются в классе ниже
df_keys_oplata = pd.read_excel(links_main(fr"{script_dir}/file_links.txt", "keys_columns"), sheet_name='ОПЛАТА')
df_keys_auto = pd.read_excel(links_main(fr"{script_dir}/file_links.txt", "keys_columns"), sheet_name='АВТО')
df_keys_sclad = pd.read_excel(links_main(fr"{script_dir}/file_links.txt", "keys_columns"), sheet_name='СКЛАД')
df_keys_arhiv = pd.read_excel(links_main(fr"{script_dir}/file_links.txt", "keys_columns"), sheet_name='АРХИВ')
df_keys_demo = pd.read_excel(links_main(fr"{script_dir}/file_links.txt", "keys_columns"), sheet_name='ДЕМО')
df_keys_neprofil = pd.read_excel(links_main(fr"{script_dir}/file_links.txt", "keys_columns"), sheet_name='НЕПРОФИЛЬ')


class Manufacturing_df_sborka:
    def __init__(self, uniq_name_autocentr, 
                 df):
        """класс сборки инофрмации из NP и СКЛАДОВ
        сбор без коррекции

        Args:
            uniq_name_autocentr (_type_): уникальное имя автоцентра которое считывается из поданного df
            df (_type_): df который соответствует ключу uniq_name_autocentr
            
        """
        self.uniq_name_autocentr:str = uniq_name_autocentr
        self.df = df
        self.date_update = None
        self.df_np_oplata = None                                                                                    # заполняется автоматически результирующий df NP Оплата
        self.df_np_auto = None                                                                                    # заполняется автоматически результирующий df NP Авто 
        self.df_sclad = None                                                                                 # заполняется автоматически результирующий df Склад (в нем может быть несколько листов СКЛАД, РЕАЛИЗАЦИЯ, ДЕМО с доп столбцом - имя_листа)
        
        # заводим 6 списков согласно нашему файлу keys_columns
        # все что в NP
        self.list_in_np_oplata = ['ОПЛАТА', 'ОПЛАТЫ', 'OPLATA']
        self.list_in_np_auto = ['АВТО', 'NP', 'СКЛАД', 'СКЛАД UKA-АУКЦИОН', 'АРХИВ']
        # все что в SCLAD
        self.list_in_sclad = ['СКЛАД', 'Склад', 'В РАБОТЕ', 'АВТО', 'СКЛАД UKA-АУКЦИОН']
        self.list_in_arhiv = ['РЕАЛИЗАЦИЯ', 'АРХИВ', 'PARK'] 
        self.list_in_demo = ['ДЕМО АВТО', 'ДЕМО АРХИВ', 'ТЕСТ-ДРАЙВ', 'DEMO', 'ДЕМО']
        self.list_in_neprofil = ['НЕПРОФИЛЬ', 'БУ', 'MAZDA NEXT']
        
        # столбцы значения в которых изменими в нижний регистр
        self.columns_lower_registr = ['склад_заказ', 'форма_оплаты', 'в_ар_хив', 'за_что', 'вид_поставки']
        
        self.fnc_auto()
    
    @staticmethod 
    def sclad_obrabotka(spisok_dopustimyh_listoy, listy_y_knige, obrabatymaevij_df, df_keys, ssulka_dlya_obrabotky, prinadlejnost, region, key_unique, kuda_sohranuem_list, poisk_vin = 'VIN'):
        """метод для поиска столбцов по ключу (ключом является ссылка) по поданной ссылке просиходит отбор с поданном df и согласно отобранным данным будет отфильтрован врем с оригинальными 
        названиями столцов, которые нужно привести к единому эталонному виду - ком являются наменования столбцов а содержание в строках - оригиналом для обработки

        Args:
            spisik_dopustimyh_listoy (_type_):               подается список с названиями которые могут встречаться в обрабатываемой книге
            listy_y_knige (_type_):                          здесь подается переменная в которую мы сохранили имена листов при считыывании книги 
            obrabatymaevij_df (_type_):                      подается df который будем обрабатывать
            df_keys (_type_):                                df с ключами, то есть лист книги со столбцами которые будеим искать и добавлять в сулчае отсутствия
            ssulka_dlya_obrabotky (_type_):                 ссылка - она вялется ключом, так как уникальна
            prinadlejnost (_type_):                         подается переменная в которой хранится значение принадлежности 
            region (_type_):                                подается переменная в которой хранится значение региона 
            key_unique (_type_):                            подается переменная в которой хранится значение ключ (не ссылка а например CHERY_vved_MSK)
            kuda_sohranuem_list (_type_):                   указывается имя переменной типа лист, куда сбразываются данные (накапливаются)
            poisk_vin (str, optional): _description_. Defaults to 'VIN'. - по какому параметру ищем шапку в листе, как правило это vin но может быть 'за что'
        """
        for list_name in  spisok_dopustimyh_listoy:                                                                                             # проходим по списку допустимых листов (те которые могут встречаться в книге)
                if list_name in listy_y_knige:                                                                                                  # если имя листа есть в книге

                    df_work = obrabatymaevij_df[list_name]                                                                                      # формируем df именно из этого листа
                    df_work = Shapka(df_work, poisk_vin)                                                                                        # находим шапку 

                    keys_import_file = df_keys[df_keys['ссылка']==fr'{ssulka_dlya_obrabotky}']                                                  # подтянули фреим с именами столбцов (фильтрация по поданной ссылке), которые нам наужны и будут переименованы
                    keys_import_file_columns = [i for i in keys_import_file.columns if i not in ['Unnamed: 0', 'ссылка']]
                    for i in keys_import_file_columns:                                                                                          # прходим по ключам и вытаскиваем названия столбцов какие нужно найти и переименовать
                        try:
                            df_work = df_work.rename(columns={list(keys_import_file[i])[0] : i}) 

                        except Exception as ex_:
                            print(f'{i} ошибка {ex_}')
    
                    for i in enumerate(df_work.columns):                                                                                        # ищем дубликаты имен столбцов и переименовываем добавляя индекс в конце
                        if i[1] in df_work.columns[:i[0]]:
                            print(f'найден дубликат столбца {i[1]} в листе {list_name} по ссылке {ssulka_dlya_obrabotky} переименован в {i[1]}_{i[0]}')
                            df_work = df_work.rename(columns={i[1]:f'{i[1]}_{i[0]}'})
                    
                    
                    # df_work = df_work.astype(str)
                    
                    for i in keys_import_file_columns:                                                                                          # проверяем наличие нужных столбцов во фрейме и если таковых нет - добавляем
                        if i not in df_work.columns:
                            df_work[i] = '0'
                     
                           
                    df_work = df_work[keys_import_file_columns]                                                                                 # оставляем только нужные нам столбцы
                    
                    
                    try:
                        df_work = df_work[~df_work['vin'].str.contains("|".join(['^VIN']), case=False, na=False)]                                 # удаляем строки если в книге было несколько строк столбцов/шапок
                    except:
                        df_work
                    
                    
                    df_work['принадлежность'] = prinadlejnost                                                                                   # добавляем столбец с принадлежностью
                    df_work['регион'] = region                                                                                                  # добавляем столбец с регионом
                    df_work['ключ'] = key_unique                                                                                                # добавляем столбец с ключом
                    df_work['с_листа'] = list_name                                                                                              # добавляем столбец с именем листа
                    df_work['ссылка'] = ssulka_dlya_obrabotky                                                                                   # ссылка на файл откуда взята информация
                    kuda_sohranuem_list.append(df_work)                                                                                         # забрасываем фрейм в базу хранения для дальнейшей конкатинации
                    
        
    def np_predobrabotka_oplata(self):
        """функция предобработки np и листов ОПЛАТА ориентирована только на эту задачу
        принимает фрейм с ключом выбирает из него все ссылки и проходит по каждой формируя отдельные фреймы 
        считывая имена листов и если находит лист ОПЛАТА формирует фреймы, копит их, конкатенирует, убирает NAN и лишние стоки 
        готовый сборный результат возвращает в свойство класса self.df_np_oplata
        """
        exception_ = ['BAIC_varsh_MSK', 'HYUNDAI_varsh_MSK', 'UKA_varsh_MSK'] # так как у них нет NP
        serch_prefix = 'SCLAD' if any(i in self.uniq_name_autocentr for i in exception_)==True else 'NP'
        
        links = list(self.df[self.df['принадлежность'].str.contains(serch_prefix)]['ссылка'])                                                   # отсортировали фрейм с ссылками для получений только файлов NP
        base_save_df = []                                                                                                                       # база хранения фреймов
        for link in links:                                                                                                                      # перебираем каждую полученную ссылку
            df_one_sroka = self.df[(self.df['принадлежность'].str.contains(serch_prefix)) & (self.df['ссылка']==link)]                          # подставляем ссылку формируя только одну строку с данными
            prinadlejnost = list(df_one_sroka['принадлежность'])[0]                                                                             # забираем значение принадлежности из фрейма
            region = list(df_one_sroka['регион'])[0]                                                                                            # забираем значение региона из фрейма
            key = list(df_one_sroka['ключ'])[0]                                                                                                 # забираем значение ключа из фрейма
            new_df, names_list = read_datafarme(fr'{link}')                                                                                     # запускай функцию предобработки получая все данные из листов и названия листов (в верх регистре)
            if self.uniq_name_autocentr not in exception_:
                self.sclad_obrabotka(self.list_in_np_oplata, names_list, new_df, df_keys_oplata, link, prinadlejnost, region, key, base_save_df, 'за что')         # обрабатываем листы СКЛАД в СКЛАДАХ
            else:
                self.sclad_obrabotka(self.list_in_np_oplata, names_list, new_df, df_keys_oplata, link, prinadlejnost, region, key, base_save_df, 'за_что') 
        
        res_df = pd.concat(base_save_df)                                                                                                        # объединили все фреймы
        res_df = res_df.dropna(subset=['дата_оплаты'])                                                                                          # удаляем nan по столбцу дата_оплаты
        
        res_df = res_df[res_df['дата_оплаты'].apply(lambda x: len(str(x))>5)]                                                                   # отсекаем все лишние значения по длине даты
        
        self.df_np_oplata = res_df                                                                                                              # присваиваем результат к переменной класса

            
    def np_predobrabotka_auto(self):
        """функция предобработки np и листов АВТО ориентирована только на эту задачу
        принимает фрейм с ключом выбирает из него все ссылки и проходит по каждой формируя отдельные фреймы 
        считывая имена листов и если находит лист АВТО формирует фреймы, копит их, конкатенирует, убирает NAN и лишние стоки 
        готовый сборный результат возвращает в свойство класса self.df_np_auto
        """
        exception_ = ['BAIC_varsh_MSK', 'HYUNDAI_varsh_MSK', 'UKA_varsh_MSK'] # так как у них нет NP
        serch_prefix = 'SCLAD' if any(i in self.uniq_name_autocentr for i in exception_)==True else 'NP'
        
        links = list(self.df[self.df['принадлежность'].str.contains(serch_prefix)]['ссылка'])                                                           # отсортировали фрейм с ссылками для получений только файлов NP
        base_save_df = []                                                                                                                       # база хранения фреймов
        for link in links:                                                                                                                      # перебираем каждую полученную ссылку
            df_one_sroka = self.df[(self.df['принадлежность'].str.contains(serch_prefix)) & (self.df['ссылка']==link)]                                  # поставляем ссылку формируя только одну строку с данными
            prinadlejnost = list(df_one_sroka['принадлежность'])[0]                                                                             # забираем значение принадлежности из фрейма
            region = list(df_one_sroka['регион'])[0]                                                                                            # забираем значение региона из фрейма
            key = list(df_one_sroka['ключ'])[0]                                                                                                 # забираем значение ключа из фрейма
            new_df, names_list = read_datafarme(fr'{link}')                                                                                     # запускай функцию предобработки получая все данные из листов и названия листов (в верх регистре)
            self.sclad_obrabotka(self.list_in_np_auto, names_list, new_df, df_keys_auto, link, prinadlejnost, region, key, base_save_df)        # обрабатываем листы СКЛАД в СКЛАДАХ
            
        
        res_df = pd.concat(base_save_df)                                                                                                        # объединили все фреймы
        res_df = res_df.dropna(subset=['модель', 'vin'], how='all')                                                                             # удаляем nan по столбцу ['модель', 'vin']
        res_df = res_df[res_df['дата_заказа'].apply(lambda x: len(str(x))>=5)]                                                                        # отсекаем все лишние значения по длине даты
        self.df_np_auto = res_df                                                                                                              # присваиваем результат к переменной класса
    
    
    def sclad_predobrabotka_all(self):
        """функция предобработки SCLAD и листов [СКЛАД, РЕАЛИЗАЦИЯ, В РАБОТЕ и тд.] ориентирована только на эту задачу
        принимает фрейм с ключом выбирает из него все ссылки и проходит по каждой формируя отдельные фреймы 
        считывая имена листов и если находит листы из списка self.list_in_sclad формирует фреймы, копит их, конкатенирует, 
        убирает NAN и лишние строки готовый сборный результат возвращает в свойство класса self.df_sclad
        """
        # так как у OVP имя файла NP а для формир SCLAD фильтруется этот префикс
        serch_prefix = 'NP' if 'OVP' in self.uniq_name_autocentr else 'SCLAD'                                                                   # если в ключе поиска есть 'OVP' то филтрует по NP во всех остальных случаях фильтрация складов по 'SCLAD'
            
        links = list(self.df[self.df['принадлежность'].str.contains(serch_prefix)]['ссылка'])                                                   # отсортировали фрейм с ссылками для получений только файлов SCLAD
        base_save_df = []                                                                                                                       # база хранения фреймов
        
        for link in links:                                                                                                                      # перебираем каждую полученную ссылку
            df_one_sroka = self.df[(self.df['принадлежность'].str.contains(serch_prefix)) & (self.df['ссылка']==link)]                          # поставляем ссылку формируя только одну строку с данными
            prinadlejnost = list(df_one_sroka['принадлежность'])[0]                                                                             # забираем значение принадлежности из фрейма
            region = list(df_one_sroka['регион'])[0]                                                                                            # забираем значение региона из фрейма
            key = list(df_one_sroka['ключ'])[0]                                                                                                 # забираем значение ключа из фрейма
            new_df, names_list = read_datafarme(fr'{link}')                                                                                     # запускай функцию предобработки получая все данные из листов и названия листов (в верх регистре)
            
            self.sclad_obrabotka(self.list_in_sclad, names_list, new_df, df_keys_sclad, link, prinadlejnost, region, key, base_save_df)         # обрабатываем листы СКЛАД в СКЛАДАХ
            self.sclad_obrabotka(self.list_in_arhiv, names_list, new_df, df_keys_arhiv, link, prinadlejnost, region,  key, base_save_df)        # обрабатываем листы АРХИВ в СКЛАДАХ
            self.sclad_obrabotka(self.list_in_demo, names_list, new_df, df_keys_demo, link, prinadlejnost, region,  key, base_save_df)          # обрабатываем листы ДЕМО в СКЛАДАХ
            self.sclad_obrabotka(self.list_in_neprofil, names_list, new_df, df_keys_neprofil, link, prinadlejnost, region,  key, base_save_df)  # обрабатываем листы НЕПРОФИЛЬ в СКЛАДАХ
        
        res_df = pd.concat(base_save_df)                                                                                                        # объединили все фреймы
        
        res_df = res_df.dropna(subset=['модель', 'vin'], how='all')                                                                             # удаляем nan по столбцу ['модель', 'vin']
        res_df = res_df[~res_df['модель'].str.contains("|".join(['Модель']), case=False, na=False)]                                             # удаляем строки если в книге было несколько строк столбцов
        self.df_sclad = res_df
        
    
  
    def lower_registr(self):
        """понижение регистра во фреймах в определенных столбцах
        """
        try:
            for df_registr in [self.df_np_oplata, self.df_np_auto, self.df_sclad]:                                                         # подаем три фрейма проходим по каждому перебирая столбцы
                for i in df_registr.columns:
                    if i in  self.columns_lower_registr:                                                                                            # если имя столбца есть в списке понижаем регистр
                        df_registr[i] = df_registr[i].apply(lambda x: str(x).strip().lower())
        except Exception as ex_:
                print(f'{self.uniq_name_autocentr} ошибка {ex_} функция - {self.lower_registr.__name__}')
            

    def result_date_update_cl(self):
        """получаем метаданные / дату обновления файла
        """
        try:
            self.date_update = result_date_update(self.uniq_name_autocentr, 'date_update')
        except Exception as ex_:
                print(f'{self.uniq_name_autocentr} ошибка {ex_} функция - {self.result_date_update_cl.__name__}')

    
  
    def save_object_class_excel(self):
        """функция сохранения промежуточного объекта класса с тремя собранными фреймами ОПЛАТА, АВТО, СКЛАД
        """
        try:
            # записываем в файл и если лист существует, то меняем лист
            with pd.ExcelWriter(rf'{links_main(fr"{script_dir}/file_links.txt", "save_file_sborka")}\{self.uniq_name_autocentr}.xlsx', 
                                engine='xlsxwriter', date_format = 'dd.mm.yyyy', datetime_format='dd.mm.yyyy') as writer:
            # Записать ваш DataFrame в файл на листы
                self.df_np_oplata.to_excel(writer, 'oplata')
                self.df_np_auto.to_excel(writer, 'auto')
                self.df_sclad.to_excel(writer, 'sclad')
        except Exception as ex_:
                print(f'{self.uniq_name_autocentr} ошибка {ex_} функция - {self.save_object_class_excel.__name__} не удалось записать данные в файл')
            

    
    def fnc_auto(self):
        """функция запуска функций
        при ручной проврке и отключении не забывать отключать сохраниение листов в функции save_object_class_excel
        """
        self.np_predobrabotka_oplata()
        self.np_predobrabotka_auto()
        self.sclad_predobrabotka_all()
        self.lower_registr()
        self.result_date_update_cl()
        self.save_object_class_excel()
        
        
# наполняем словарь базами данных создавая экземпляры класса
catalog_df = {} # словарь со всеми базами

catalog_exception_key = [] # ключи ошибок
count_bd = len(df_main.ключ.unique())
for i in df_main.ключ.unique():
    try:
        print(f'{i}-----------------------------------')
        LOG_inf(f'Создаем объект класса {Manufacturing_df_sborka.__name__}', 'INFO', i)
        catalog_df[i] = Manufacturing_df_sborka(i, df_main[df_main['ключ']==i])
        
        count_bd-=1
        print(f'Осталось создать {count_bd} объектов класса')
        LOG_inf(f'Осталось создать {count_bd} объектов класса', 'INFO')
    except Exception as ex_:
        catalog_exception_key.append(i)
        print(f'ошибка по ключу {i} будет попытка перезапуска')
        LOG_inf(f'ошибка по ключу {i} будет попытка перезапуска', 'ERROR')
        
if len(catalog_exception_key)>0:
    for i in catalog_exception_key:
        print(f'Повторная попытка создать объъект класса')
        print(f'{i}-----------------------------------')
        catalog_df[i] = Manufacturing_df_sborka(i, df_main[df_main['ключ']==i])
        
        

LOG_inf(f'Создано объектов класса {Manufacturing_df_sborka.__name__} в кол-ве {len(catalog_df)}', 'INFO')


# добавление демо авто из склада в NP и Оплата по Саратову (КОСТЫЛИ САРАТОВА)
LOG_inf(f'добавление демо авто из склада в NP и Оплата по Саратову (КОСТЫЛИ САРАТОВА)', 'INFO')
try:
    copy_object_class = copy.deepcopy(catalog_df['OMODA_SAR'])
    columns_intersection_np_sclad = set(copy_object_class.df_np_auto.columns).intersection(copy_object_class.df_sclad.columns) # столбцы пересекаются
    columns_difference_np_sclad = set(copy_object_class.df_np_auto.columns).difference(copy_object_class.df_sclad.columns)     # столбцы не пересекаются
    # выбираем только демо выданные
    new_upate_np = copy_object_class.df_sclad[copy_object_class.df_sclad['с_листа'].str.contains("|".join(['demo', 'демо']), case=False, na=False) 
                                            & copy_object_class.df_sclad['дата_продажи_факт'].notna()][list(columns_intersection_np_sclad)]
    # добавляем недостающие столбцы, которые должны быть в NP
    for i in columns_difference_np_sclad:
        new_upate_np[i]='0'
    # уполрядочиваем столбцы
    new_upate_np = new_upate_np[[i for i in copy_object_class.df_np_auto.columns]]
    # заполняем данные
    new_upate_np['форма_оплаты'] = 'нал'
    new_upate_np['дата_выдачи_факт'] = new_upate_np.apply(lambda x: one_unique_pokazatel(copy_object_class.df_sclad, x.vin, 'дата_продажи_факт'), axis=1)
    new_upate_np['дата_заказа'] = new_upate_np.apply(lambda x: one_unique_pokazatel(copy_object_class.df_sclad, x.vin, 'дата_продажи_факт'), axis=1)
    new_upate_np['дата_полной_оплаты_факт'] = new_upate_np.apply(lambda x: one_unique_pokazatel(copy_object_class.df_sclad, x.vin, 'дата_продажи_факт'), axis=1)
    new_upate_np['дата_справки_счет_факт'] = new_upate_np.apply(lambda x: one_unique_pokazatel(copy_object_class.df_sclad, x.vin, 'дата_продажи_факт'), axis=1)
    new_upate_np['получено_за_ам_руб'] = new_upate_np.apply(lambda x: one_unique_pokazatel(copy_object_class.df_sclad, x.vin, 'цена_продажи'), axis=1)
    new_upate_np['склад_заказ'] = 'на складе'
    # new_upate_np['в_ар_хив'] = 'да'
    new_upate_np['площадка'] = 'Саратов'
    new_upate_np['с_листа'] = 'АВТО'
    new_upate_np['сотрудник_продал'] = new_upate_np.apply(lambda x: one_unique_pokazatel(copy_object_class.df_sclad, x.vin, 'клиент'), axis=1)
    # объединяем основной NP и наш кусочек выбранный из СКЛАДА по ДЕМО
    res_new_np = pd.concat([copy_object_class.df_np_auto, new_upate_np])
    # присваиваем данный результат объекту класса, в переменную NP
    catalog_df['OMODA_SAR'].df_np_auto = copy.deepcopy(res_new_np)


    columns_intersection_oplata_sclad = set(copy_object_class.df_np_oplata.columns).intersection(copy_object_class.df_sclad.columns) # столбцы пересекаются
    columns_difference_oplata_sclad = set(copy_object_class.df_np_oplata.columns).difference(copy_object_class.df_sclad.columns)     # столбцы не пересекаются
    # выбираем только демо выданные
    new_upate_oplata = copy_object_class.df_sclad[copy_object_class.df_sclad['с_листа'].str.contains("|".join(['demo', 'демо']), case=False, na=False) 
                                            & copy_object_class.df_sclad['дата_продажи_факт'].notna()][list(columns_intersection_oplata_sclad)]
    # добавляем недостающие столбцы, которые должны быть в NP
    for i in columns_difference_oplata_sclad:
        new_upate_oplata[i]='0'
    # уполрядочиваем столбцы
    new_upate_oplata = new_upate_oplata[[i for i in copy_object_class.df_np_oplata.columns]]
    # заполняем данные
    new_upate_oplata['дата_оплаты'] = new_upate_oplata.apply(lambda x: one_unique_pokazatel(copy_object_class.df_sclad, x.vin, 'дата_продажи_факт'), axis=1)
    new_upate_oplata['внесено_в_рублях'] = new_upate_oplata.apply(lambda x: one_unique_pokazatel(copy_object_class.df_sclad, x.vin, 'цена_продажи'), axis=1)
    new_upate_oplata['за_что'] = 'а/м'
    new_upate_oplata['с_листа'] = 'ОПЛАТА'
    # объединяем основной фрейм оплатв и наш кусочек выбранный из СКЛАДА по ДЕМО
    res_new_oplata = pd.concat([copy_object_class.df_np_oplata, new_upate_oplata])
    # присваиваем данный результат объекту класса, в переменную оплата
    catalog_df['OMODA_SAR'].df_np_oplata = copy.deepcopy(res_new_oplata)

except Exception as ex_:
    print(f'ОШИБКА {ex_} не удалось добавить добавление демо авто из склада в NP и Оплата по Саратову (КОСТЫЛИ САРАТОВА) {ex_}')
    LOG_inf(f'добавление демо авто из склада в NP и Оплата по Саратову (КОСТЫЛИ САРАТОВА)', 'ERROR')
    
    
# запускать после отработки Manufacturing_df_sborka
def OMODA_JAECOO_SAR():
    """функция заберает из каталога catalog_df OMODA_SAR
    на основании его копии создает два объекта (JAECOO__SAR OMODA__SAR)
    фильтрует данные в  (df_np_auto, df_np_oplata, df_sclad) на 'JAECOO', 'OMODA'
    создает новые объекты по маркам 'JAECOO', 'OMODA'
    """
    LOG_inf(f'{OMODA_JAECOO_SAR.__name__}', 'INFO')
    try:
        for i in ['JAECOO', 'OMODA']:
            catalog_df[f'{i}__SAR'] = copy.deepcopy(catalog_df['OMODA_SAR'])
            catalog_df[f'{i}__SAR'].df_np_auto = catalog_df[f'{i}__SAR'].df_np_auto[catalog_df[f'{i}__SAR'].df_np_auto ['модель'].str.contains("|".join([i]), case=False)]
            catalog_df[f'{i}__SAR'].df_np_oplata =  catalog_df[f'{i}__SAR'].df_np_oplata[catalog_df[f'{i}__SAR'].df_np_oplata['модель'].str.contains("|".join([i]), case=False)]
            catalog_df[f'{i}__SAR'].df_sclad = catalog_df[f'{i}__SAR'].df_sclad[catalog_df[f'{i}__SAR'].df_sclad['модель'].str.contains("|".join([i]), case=False)]
            
        # удаляем объект из каталога так как дальше у нас все распределено
        del catalog_df['OMODA_SAR']
    except Exception as ex_:
        print(f'ошибка функции {OMODA_JAECOO_SAR.__name__} {ex_}')
        LOG_inf(f'ошибка функции {OMODA_JAECOO_SAR.__name__} {ex_}', 'ERROR')
        
        
OMODA_JAECOO_SAR()


# Разделяем ОВП ЯР на Яр и РЫБИНСК
LOG_inf(f'Разделяем ОВП ЯР на Яр и РЫБИНСК', 'INFO')
try:
    catalog_df['OVP__YAR'] = copy.deepcopy(catalog_df['OVP_YAR'])
    catalog_df['OVP__YAR'].df_np_oplata = catalog_df['OVP__YAR'].df_np_oplata[catalog_df['OVP__YAR'].df_np_oplata['локация'].str.contains("|".join(['Ярославль']), case=False, na=False)]
    catalog_df['OVP__YAR'].df_np_auto = catalog_df['OVP__YAR'].df_np_auto[catalog_df['OVP__YAR'].df_np_auto['площадка'].str.contains("|".join(['Ярославль']), case=False, na=False)]
    catalog_df['OVP__YAR'].df_sclad = catalog_df['OVP__YAR'].df_sclad[catalog_df['OVP__YAR'].df_sclad['площадка'].str.contains("|".join(['Ярославль']), case=False, na=False)]

    catalog_df['OVP__RYB'] = copy.deepcopy(catalog_df['OVP_YAR'])
    catalog_df['OVP__RYB'].df_np_oplata = catalog_df['OVP__RYB'].df_np_oplata[~catalog_df['OVP__RYB'].df_np_oplata['локация'].str.contains("|".join(['Ярославль']), case=False, na=False)]
    catalog_df['OVP__RYB'].df_np_auto = catalog_df['OVP__RYB'].df_np_auto[~catalog_df['OVP__RYB'].df_np_auto['площадка'].str.contains("|".join(['Ярославль']), case=False, na=False)]
    catalog_df['OVP__RYB'].df_sclad = catalog_df['OVP__RYB'].df_sclad[~catalog_df['OVP__RYB'].df_sclad['площадка'].str.contains("|".join(['Ярославль']), case=False, na=False)]
    del catalog_df['OVP_YAR'] # удаляем исходник, так как разделили на дву состоявляющих
except Exception as ex_:
    print(f'ошибка {ex_} не удалось разделить ОВП ЯР на Яр и РЫБИНСК')
    LOG_inf(f'ошибка {ex_} не удалось разделить ОВП ЯР на Яр и РЫБИНСК', 'ERROR')
    
    
# разделяем ДЖЕТУР МСК на ДЖЕТУР И НЕПРОФИЛЬ
LOG_inf(f'разделяем ДЖЕТУР МСК на ДЖЕТУР И НЕПРОФИЛЬ', 'INFO')
try:
    catalog_df['JETOUR_MSK'] = copy.deepcopy(catalog_df['JETOUR_vved_MSK'])
    catalog_df['JETOUR_MSK'].df_sclad = catalog_df['JETOUR_MSK'].df_sclad[~catalog_df['JETOUR_MSK'].df_sclad['с_листа'].str.contains("|".join(['НЕПРОФИЛЬ']), case=False, na=False)]
    catalog_df['JETOUR_MSK'].df_np_auto = catalog_df['JETOUR_MSK'].df_np_auto[~catalog_df['JETOUR_MSK'].df_np_auto['статус_оригинал'].str.contains("|".join(['б/у','next']), case=False, na=False)]
    id_zakaza_jetour = catalog_df['JETOUR_MSK'].df_np_auto['id'].unique()
    catalog_df['JETOUR_MSK'].df_np_oplata = catalog_df['JETOUR_MSK'].df_np_oplata[catalog_df['JETOUR_MSK'].df_np_oplata['id'].str.contains("|".join(id_zakaza_jetour), case=False, na=False)]

    catalog_df['JETOURneprof_MSK'] = copy.deepcopy(catalog_df['JETOUR_vved_MSK'])
    catalog_df['JETOURneprof_MSK'].df_sclad = catalog_df['JETOURneprof_MSK'].df_sclad[catalog_df['JETOURneprof_MSK'].df_sclad['с_листа'].str.contains("|".join(['НЕПРОФИЛЬ']), case=False, na=False)]
    catalog_df['JETOURneprof_MSK'].df_sclad['дата_оплаты_счета'] = catalog_df['JETOURneprof_MSK'].df_sclad.apply(lambda x: (x.дата_прихода_на_склад if len(str(x.дата_оплаты_счета))<5 else x.дата_оплаты_счета), axis=1)
    catalog_df['JETOURneprof_MSK'].df_np_auto = catalog_df['JETOURneprof_MSK'].df_np_auto[catalog_df['JETOURneprof_MSK'].df_np_auto['статус_оригинал'].str.contains("|".join(['б/у','next']), case=False, na=False)]
    id_zakaza_jetour_neprof = catalog_df['JETOURneprof_MSK'].df_np_auto['id'].unique()
    catalog_df['JETOURneprof_MSK'].df_np_oplata = catalog_df['JETOURneprof_MSK'].df_np_oplata[catalog_df['JETOURneprof_MSK'].df_np_oplata['id'].str.contains("|".join(id_zakaza_jetour_neprof), case=False, na=False)]
    del catalog_df['JETOUR_vved_MSK'] # удаляем исходник, так как разделили на дву состоявляющих
except Exception as ex_:
    print(f'ошибка {ex_} не удалось разделить ДЖЕТУР МСК на ДЖЕТУР И НЕПРОФИЛЬ')
    LOG_inf(f'ошибка {ex_} не удалось разделить ДЖЕТУР МСК на ДЖЕТУР И НЕПРОФИЛЬ', 'ERROR')
    
    
# разделяем МАЗДУ МСК на МАЗДУ И NEXT
LOG_inf(f'разделяем МАЗДУ МСК на МАЗДУ И NEXT', 'INFO')
try:
    catalog_df['MAZDA_MSK'] = copy.deepcopy(catalog_df['MAZDA_vved_MSK'])
    catalog_df['MAZDA_MSK'].df_sclad = catalog_df['MAZDA_MSK'].df_sclad[~catalog_df['MAZDA_MSK'].df_sclad['с_листа'].str.contains("|".join(['NEXT']), case=False, na=False)]
    catalog_df['MAZDA_MSK'].df_np_auto = catalog_df['MAZDA_MSK'].df_np_auto[~catalog_df['MAZDA_MSK'].df_np_auto['статус_оригинал'].str.contains("|".join(['б/у','next']), case=False, na=False)]
    id_zakaza_MAZDA = catalog_df['MAZDA_MSK'].df_np_auto['id'].unique()
    catalog_df['MAZDA_MSK'].df_np_oplata = catalog_df['MAZDA_MSK'].df_np_oplata[catalog_df['MAZDA_MSK'].df_np_oplata['id'].str.contains("|".join(id_zakaza_MAZDA), case=False, na=False)]


    catalog_df['MAZDAnext_MSK'] = copy.deepcopy(catalog_df['MAZDA_vved_MSK'])
    catalog_df['MAZDAnext_MSK'].df_sclad = catalog_df['MAZDAnext_MSK'].df_sclad[catalog_df['MAZDAnext_MSK'].df_sclad['с_листа'].str.contains("|".join(['NEXT']), case=False, na=False)]
    catalog_df['MAZDAnext_MSK'].df_sclad['дата_оплаты_счета'] = catalog_df['MAZDAnext_MSK'].df_sclad.apply(lambda x: (x.дата_прихода_на_склад if len(str(x.дата_оплаты_счета))<5 else x.дата_оплаты_счета), axis=1)
    catalog_df['MAZDAnext_MSK'].df_np_auto = catalog_df['MAZDAnext_MSK'].df_np_auto[catalog_df['MAZDAnext_MSK'].df_np_auto['статус_оригинал'].str.contains("|".join(['б/у','next']), case=False, na=False)]
    id_zakaza_MAZDA_next = catalog_df['MAZDAnext_MSK'].df_np_auto['id'].unique()
    catalog_df['MAZDAnext_MSK'].df_np_oplata = catalog_df['MAZDAnext_MSK'].df_np_oplata[catalog_df['MAZDAnext_MSK'].df_np_oplata['id'].str.contains("|".join(id_zakaza_MAZDA_next), case=False, na=False)]
    del catalog_df['MAZDA_vved_MSK'] # удаляем исходник, так как разделили на дву состоявляющих
except Exception as ex_:
    print(f'ошибка {ex_} не удалось разделить МАЗДУ МСК на МАЗДУ И NEXT')
    LOG_inf(f'ошибка {ex_} не удалось разделить МАЗДУ МСК на МАЗДУ И NEXT', 'ERROR')
    
    

class Manufacturing_df_predobrabotka:
    def __init__(self, manufacturing_df_sborka, name_Manufacturing_df_sborka, starter:True or False = None, save_excel: True or False = None):
        """промежуточный класс предобработки данных
        подготавливает данные для рассчета

        Args:
            Manufacturing_df_sborka (_type_): # Объект класса предобработки
            name_Manufacturing_df_sborka (_type_): # имя  Объекта класса предобработки
            starter - запускает или нет все функции в классе
            save_excel - сохраняем результаты в Excel
        """
        self.object_class_ = copy.deepcopy(manufacturing_df_sborka)             # Объект класса предобработки copy.deepcopy - копируем создавая новую ячейку хранения (чтоб не изменялся исходник)
        self.name_object_class = name_Manufacturing_df_sborka                   # имя  Объекта класса предобработки
        self.df_np_oplata = self.object_class_.df_np_oplata
        self.df_np_auto = self.object_class_.df_np_auto
        self.df_sclad = self.object_class_.df_sclad
        self.date_update = self.object_class_.date_update  # получаем три df из объекта класса Manufacturing_df и дату
        self.starter = starter
        self.save_excel = save_excel
        self.except_kum = None
        self.white_list_columns_except_logist = ['с_листа', 'регион', 'принадлежность', 'ключ', 'ссылка', 'vin', 'ошибка'] # наименование столбцов, которые оставляем при сборе ошибок для логистов
        
        # преобразование дат и чисел 
        # self.df_np_oplata, self.df_np_auto, self.df_sclad = datetime_columns_convertor(self.df_np_oplata), datetime_columns_convertor(self.df_np_auto), datetime_columns_convertor(self.df_sclad)
        # self.df_np_oplata, self.df_np_auto, self.df_sclad = numeric_columns_convertor(self.df_np_oplata), numeric_columns_convertor(self.df_np_auto), numeric_columns_convertor(self.df_sclad)
        
        self.fnc_auto()
       
    @staticmethod
    def new_df_except(df):
        
        """шаблонный df с фикс данными

        Args:
            df (_type_): _description_

        Returns:
            _type_: _description_
        """
        
        df_except = pd.DataFrame()
        df_except['с_листа'] = df['с_листа']
        df_except['регион'] = df['регион']
        df_except['принадлежность'] = df['принадлежность']
        df_except['ключ'] = df['ключ']
        df_except['ссылка'] = df['ссылка']
        return df_except
        
        
    def excepts_date(self):
        """сбор ошибок в столбцах дат для логистов
        """
        except_list = []
        for frame in [self.df_np_oplata, self.df_np_auto, self.df_sclad]:
            for i in frame.columns:
                if 'дата' in i:
                    try:
                        test_frame = copy.deepcopy(frame)
                        test_frame['ошибка'] = test_frame.apply(lambda x: f'{shablon_date_test_pravka(x[i])} в столбце {i}', axis=1) # shablon_date_test_2
                        test_frame = test_frame[~(test_frame[f'ошибка'].str.contains('|'.join(['ok', 'nan', 'None', 'NaT', '-', '00:00:00'])))]
                        test_frame = test_frame[self.white_list_columns_except_logist]
                        except_list.append(test_frame)
                        
                        # test_frame = copy.deepcopy(frame)
                        # test_frame['ошибка'] = test_frame.apply(lambda x: f'{x[i]} в столбце {i}', axis=1)
                        # # display(test_frame)
                        # test_frame = test_frame[~(test_frame[f'ошибка'].str.contains('|'.join(['ok', 'nan', 'None', 'NaT', '-', '00:00:00'])))]
                        # test_frame = test_frame[self.white_list_columns_except_logist]
                        # except_list.append(test_frame)
                    except:
                        None
        
        try:
            self.except_kum = pd.concat(except_list)
        except Exception as ex_:
            print(f' Ошибка конкатинции функции {self.excepts_date.__name__}')
        
    
        
        
    def korrektirovka(self):
        """корректируем форму оплаты
            предобробатываем даты 
            удаляем пробелы по VIN
        """
        for frame in [self.df_np_oplata, self.df_np_auto, self.df_sclad]:
            for i in frame.columns:
                if i == 'форма_оплаты':
                    frame[i] = frame[i].apply(lambda x: forma_pay(x))          # приводим вид оплат к формату кре/нал конверитируя все остальные формы  
                if 'дата' in i:
                    frame[i] = frame[i].apply(lambda x: del_letters_date(x))   # корректируем даты убираем буквы - оставляем даты
                    frame[i] = frame[i].apply(lambda x: shablon_date_test(x))  # проверяем даты по шаблону - очищаем и возвращаем только дату
                if 'vin' in i:
                    frame[i] = frame[i].apply(lambda x: str(x).strip())         # удаляем лишние пробелы в vin 
        
        
    def pravka_type_dataframe(self):
        """преобразует столлбцы дат и чисел
        """
        # преобразование дат и чисел 
        try: 
            self.df_np_oplata, self.df_np_auto, self.df_sclad = datetime_columns_convertor(self.df_np_oplata), datetime_columns_convertor(self.df_np_auto), datetime_columns_convertor(self.df_sclad)
            self.df_np_oplata, self.df_np_auto, self.df_sclad = numeric_columns_convertor(self.df_np_oplata), numeric_columns_convertor(self.df_np_auto), numeric_columns_convertor(self.df_sclad)
        except Exception as ex_:
            print(f'ошибка функции {self.pravka_type_dataframe.__name__}')
            
    
    def kostraciva_po_date(self):
        """обрезаем фреймы до определенного года
        данные меньше 2020 года - никто не смотрит 
        
        """
        try:
            # self.df_np_auto, self.df_sclad, self.df_np_oplata =  kostraciya(self.df_np_auto, self.df_sclad, self.df_np_oplata, KOSTRACIVA) # оригинал до kostraciya_2
            
            np, skl, opl = kostraciya(self.df_np_auto, self.df_sclad, self.df_np_oplata, KOSTRACIVA) # фреймы после кострации
            np1, skl1 = kostraciya_2(self.df_np_auto, self.df_sclad, self.df_np_oplata, KOSTRACIVA)  # фремы только НП и СКЛАД  только те авто которые пришли на склад до кострации
            self.df_np_auto = pd.concat([np, np1])                                                   # конкатинируем
            self.df_sclad = pd.concat([skl, skl1])
            self.df_np_oplata = opl
            
        except:
            print(f'ошибка функции {self.kostraciva_po_date.__name__}')

    
    def proverka_np_date(self): # добавить функцию исправления 
        """проверяет np на логику дат 
        """
        try:
            # если дата заказа больше даты фактической выдачи
            df_except = copy.deepcopy(self.df_np_auto)
            df_except['ошибка'] = df_except.apply(lambda x: (f'{x.vin} дата заказа {x.дата_заказа} больше даты выдачи {x.дата_выдачи_факт}' if x.дата_заказа>x.дата_выдачи_факт else None), axis=1)
            df_except = df_except[df_except['ошибка'].notna()]
            df_except = df_except[self.white_list_columns_except_logist]
            self.except_kum = pd.concat([self.except_kum ,df_except])
        
        
            # если дата прихода больше даты фактической выдачи 
            df_except = copy.deepcopy(self.df_np_auto)
            df_except['ошибка'] = df_except.apply(lambda x: (f'{x.vin} дата прихода {x.дата_прихода_на_склад} больше даты выдачи {x.дата_выдачи_факт}' if x.дата_прихода_на_склад>x.дата_выдачи_факт else None), axis=1)
            df_except = df_except[df_except['ошибка'].notna()]
            df_except = df_except[self.white_list_columns_except_logist]
            self.except_kum = pd.concat([self.except_kum ,df_except])
        except:
            None
        
            
    
    def pravka_formy_oplaty(self):  
        """функция проверяет форму оплаты по vin на складе, сверяя с NP по vin с заполненной датой выдачи
        форма оплаты по NP является приоритетной
        """
        try:
            self.df_sclad['форма_оплаты'] = self.df_sclad.apply(lambda x:  korrekt_forma_oplaty(self.df_np_auto, x.vin, x.форма_оплаты, 'дата_выдачи_факт', 'форма_оплаты'), axis=1)
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.pravka_formy_oplaty.__name__}')
        

    def proverka_otkaza(self):
        """проверяет логику дат выданных авто на отсутствие отказной даты
        если авто выдан - дата изм (дата отказа должна быть пустой)
        данные с ошибками собираются логистам для отправки
        """
        try:
            df_except = copy.deepcopy(self.df_np_auto)
            df_except['ошибка'] = self.df_np_auto.apply(lambda x: (f'{x.vin} авто выдан {x.дата_выдачи_факт} и не может быть отказа {x.дата_изм} в отказе дб - пусто' 
                                                                                 if len(str(x.дата_выдачи_факт))>5 and len(str(x.дата_изм))>5 else None), axis=1)
            df_except = df_except[df_except['ошибка'].notna()]
            df_except = df_except[self.white_list_columns_except_logist]
            self.except_kum = pd.concat([self.except_kum ,df_except])
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.proverka_otkaza.__name__}')
            None
        
    def pravka_otkaza(self): # добавить в ошибки для логистов proverka_pravka_otkaza - готово
        """корректирует отказ - столбец дата изменения NP / self.df_np_auto
        если авто выдан (заполнена дата фактической выдачи) удаляет значение отказа из дата_изм - для нормализации данных
        """
        try:
            self.df_np_auto['дата_изм'] = self.df_np_auto.apply(lambda x: (None if len(str(x.дата_выдачи_факт))>5 and len(str(x.дата_изм))>5 else x.дата_изм), axis=1)
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.pravka_otkaza.__name__}')
            self.df_np_auto = self.df_np_auto
            
    
    def proverka_otkaza_arhiva(self):
        """проверяет NP / self.df_np_auto если авто отказной статус ДА но нет даты отказа - 
        прописывает ошибку для исправления логистам 
        """
        
        try:
            df_except = copy.deepcopy(self.df_np_auto) 
            df_except['ошибка'] = self.df_np_auto.apply(lambda x: (f'{x.vin} авто отказной со статусом в архив - Да с датой заказа {x.дата_заказа} должна быть заполнена - дата отказа дата изм {x.дата_изм},\
            если есть дата выдачи {x.дата_выдачи_факт} убрать статус - Да / В архив' if str(x.в_ар_хив).lower()=='да' and len(str(x.дата_изм))<5  else None), axis=1)
            df_except = df_except[df_except['ошибка'].notna()]
            df_except = df_except[self.white_list_columns_except_logist]
            self.except_kum = pd.concat([self.except_kum ,df_except])
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.proverka_otkaza_arhiva.__name__}')
            None
        
        
    def pravka_otkaza_arhiva(self):
        """проверяет NP / self.df_np_auto если авто отказной статус ДА но нет даты отказа - 
        добавляет + 5 дней к дате заказа и проставляет значение в дату отказа - для нормализации данных
        """
        import datetime
        try:
            self.df_np_auto['дата_изм'] = self.df_np_auto.apply(lambda x: (x.дата_заказа+datetime.timedelta(days=5) if str(x.в_ар_хив).lower()=='да' and len(str(x.дата_изм))<5 else x.дата_изм), axis=1)
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.pravka_otkaza_arhiva.__name__}')
            self.df_np_auto = self.df_np_auto

        
    # def proverka_daty_prihoda_mejdy_sclad_i_np(self):
    #     """сравнивает даты прихода авто на складе с данными NP
    #     """
    #     try:
    #         df_except = copy.deepcopy(self.df_np_auto) 
    #         df_except['ошибка'] =self.df_np_auto.apply(lambda x: (sravn_date_prih(self.df_sclad, x.vin, x.дата_прихода_на_склад)), axis=1)
    #         df_except = df_except[df_except['ошибка'].notna()]
    #         df_except = df_except[self.white_list_columns_except_logist]
    #         self.except_kum = pd.concat([self.except_kum ,df_except])
    #     except Exception as ex_:
    #         print(f'{self.name_object_class} ошибка {ex_} функция - {self.proverka_daty_prihoda_mejdy_sclad_i_np.__name__}')
            
            
    # если заполнена дата оплаты счета то должна быть цена продажи
    def proverka_sklada_na_daty_oplaty_i_ceny_prodajy(self):
        """проверка склада на дату оплаты счета и цену продажи
        Если заполнена дата оплаты счета - должна стоять цена продажи
        """
        try:
            if self.name_object_class not in ('OVP_YAR', 'OVP__YAR', 'OVP__RYB', 'OVP_SAR', 'OVP_vved_MSK'):
                df_except = copy.deepcopy(self.df_sclad) 
                df_except['ошибка'] = self.df_sclad.apply(lambda x: (f'{x.vin} дата оплаты счета {x.дата_оплаты_счета} - цена продажи {x.цена_продажи}' 
                                                                    if len(str(x.дата_оплаты_счета))>5 and len(str(x.цена_продажи))<=1 else None), axis=1)
                df_except = df_except[df_except['ошибка'].notna()]
                df_except = df_except[self.white_list_columns_except_logist]
                self.except_kum = pd.concat([self.except_kum ,df_except])
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.proverka_sklada_na_daty_oplaty_i_ceny_prodajy.__name__}')
            None
        
    def proverka_daty_oplaty(self):
        """проверяет лист оплаты по дате оплаты, если она больше текущего дня 
        """
        from datetime import datetime
        timestamp = datetime.strptime(tek_day(), "%Y-%m-%d")
        try:
            
            df_except = copy.deepcopy(self.df_np_oplata)
            df_except['ошибка'] = self.df_np_oplata.apply(lambda x: (f'{x.vin} дата оплаты {x.дата_оплаты} больше сегодня {timestamp}' if x.дата_оплаты > timestamp else None), axis=1)
            df_except = df_except[df_except['ошибка'].notna()]
            df_except = df_except[self.white_list_columns_except_logist]
            self.except_kum = pd.concat([self.except_kum ,df_except])
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.proverka_daty_oplaty.__name__}')
            None
            
    ######################### Блок дат склада больше ли они текущего дня
    def proverka_daty_oplaty_scheta_na_sklade(self):
        """проверяет лист склад дата_оплаты_счета, если она больше текущего дня 
        """
        from datetime import datetime
        timestamp = datetime.strptime(tek_day(), "%Y-%m-%d")
        try:
            
            df_except = copy.deepcopy(self.df_sclad)
            df_except['ошибка'] = self.df_sclad.apply(lambda x: (f'{x.vin} дата оплаты {x.дата_оплаты_счета} больше сегодня {timestamp}' if x.дата_оплаты_счета > timestamp else None), axis=1)
            df_except = df_except[df_except['ошибка'].notna()]
            df_except = df_except[self.white_list_columns_except_logist]
            self.except_kum = pd.concat([self.except_kum ,df_except])
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.proverka_daty_oplaty_scheta_na_sklade.__name__}')
            None
    
    def proverka_daty_prihoda_na_sclad(self):
        """проверяет лист склад дата_прихода_на_склад, если она больше текущего дня 
        """
        from datetime import datetime
        timestamp = datetime.strptime(tek_day(), "%Y-%m-%d")
        try:
            
            df_except = copy.deepcopy(self.df_sclad)
            df_except['ошибка'] = self.df_sclad.apply(lambda x: (f'{x.vin} дата оплаты {x.дата_прихода_на_склад} больше сегодня {timestamp}' if x.дата_прихода_на_склад > timestamp else None), axis=1)
            df_except = df_except[df_except['ошибка'].notna()]
            df_except = df_except[self.white_list_columns_except_logist]
            self.except_kum = pd.concat([self.except_kum ,df_except])
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.proverka_daty_prihoda_na_sclad.__name__}')
            None
            
    def proverka_daty_contrakta_na_skale(self):
        """проверяет лист склад дата_контракта_заказа, если она больше текущего дня 
        """
        from datetime import datetime
        timestamp = datetime.strptime(tek_day(), "%Y-%m-%d")
        try:
            
            df_except = copy.deepcopy(self.df_sclad)
            df_except['ошибка'] = self.df_sclad.apply(lambda x: (f'{x.vin} дата оплаты {x.дата_контракта_заказа} больше сегодня {timestamp}' if x.дата_контракта_заказа > timestamp else None), axis=1)
            df_except = df_except[df_except['ошибка'].notna()]
            df_except = df_except[self.white_list_columns_except_logist]
            self.except_kum = pd.concat([self.except_kum ,df_except])
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.proverka_daty_contrakta_na_skale.__name__}')
            None
            
    def proverka_daty_prodajy_na_skale(self):
        """проверяет лист склад дата_продажи_факт, если она больше текущего дня 
        """
        from datetime import datetime
        timestamp = datetime.strptime(tek_day(), "%Y-%m-%d")
        try:
            
            df_except = copy.deepcopy(self.df_sclad)
            df_except['ошибка'] = self.df_sclad.apply(lambda x: (f'{x.vin} дата оплаты {x.дата_продажи_факт} больше сегодня {timestamp}' if x.дата_продажи_факт > timestamp else None), axis=1)
            df_except = df_except[df_except['ошибка'].notna()]
            df_except = df_except[self.white_list_columns_except_logist]
            self.except_kum = pd.concat([self.except_kum ,df_except])
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.proverka_daty_contrakta_na_skale.__name__}')
            None
    
    
    ####################
            
    
    def proverka_daty_prihoda(self):
        """проверяет NP дату прихода на склад - чтоб она была не больше текущей даты
        """
        from datetime import datetime
        timestamp = datetime.strptime(tek_day(), "%Y-%m-%d")
        try:
            
            df_except = copy.deepcopy(self.df_np_auto)
            df_except['ошибка'] = self.df_np_auto.apply(lambda x: (f'{x.vin} Дата прихода на склад {x.дата_прихода_на_склад} больше текущей даты {timestamp}' if x.дата_прихода_на_склад > timestamp else None), axis=1)
            df_except = df_except[df_except['ошибка'].notna()]
            df_except = df_except[self.white_list_columns_except_logist]
            self.except_kum = pd.concat([self.except_kum ,df_except])
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.proverka_daty_prihoda.__name__}')
            None
            


    def proverka_date_oplaty_na_min_date_prihoda(self):
        
        try:
            df_except = copy.deepcopy(self.df_np_oplata)
            df_except['ошибка'] = self.df_np_oplata.apply(lambda x: (f"{x.vin} {x.дата_оплаты} дата оплаты сильно меньше MIN даты прихода на склад {min_year_date_column(self.df_sclad, 'дата_прихода_на_склад')} год" 
                                                                    if (int(min_year_date_column(self.df_sclad, 'дата_прихода_на_склад')) - int(x.дата_оплаты.year))>=2 else None), axis=1)
            df_except = df_except[df_except['ошибка'].notna()]
            df_except = df_except[self.white_list_columns_except_logist]
            self.except_kum = pd.concat([self.except_kum ,df_except])
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.proverka_date_oplaty_na_min_date_prihoda.__name__}')
            None
            
            
    def proverka_date_prihoda_na_mean_date_prihoda(self):
        
        try:
            df_except = copy.deepcopy(self.df_sclad)
            df_except['ошибка'] = self.df_sclad.apply(lambda x: (f"{x.vin} {x.дата_прихода_на_склад} дата прихода сильно меньше СРЕДНЕЙ даты прихода на склад {mean_year_date_column(self.df_sclad, 'дата_прихода_на_склад')} год" 
                                                                    if abs((float(mean_year_date_column(self.df_sclad, 'дата_прихода_на_склад')) - float(x.дата_прихода_на_склад.year)))>5 else None), axis=1)
            df_except = df_except[df_except['ошибка'].notna()]
            df_except = df_except[self.white_list_columns_except_logist]
            self.except_kum = pd.concat([self.except_kum ,df_except])
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.proverka_date_prihoda_na_mean_date_prihoda.__name__}')
            None
            
    def proverka_spravki_schet_i_vidachy(self):
        """если есть правка счет но нет выдачи
        """
        try:
            df_except = copy.deepcopy(self.df_np_auto)
            df_except['ошибка'] = self.df_np_auto.apply(lambda x: (f'{x.vin} дата справки счет {x.дата_справки_счет_факт} выдачи нет {x.дата_выдачи_факт} дата спр.счет заполняется только по выданным авто' 
                                                                   if len(str(x.дата_справки_счет_факт))>5 and len(str(x.дата_выдачи_факт))<5  else None), axis=1)
            df_except = df_except[df_except['ошибка'].notna()]
            df_except = df_except[self.white_list_columns_except_logist]
            self.except_kum = pd.concat([self.except_kum ,df_except])
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.proverka_spravki_schet_i_vidachy.__name__}')
            None
            
    def pravka_spravki_schet_i_vidachy(self):
        """правим дату справки счет в NP 
        если есть дата спр счет но нет выдачи и есть статус ДА в рахив - удалям дату спр счет
        """
        try:
            self.df_np_auto['дата_справки_счет_факт'] = self.df_np_auto.apply(lambda x: (None if len(str(x.дата_справки_счет_факт))>5 
                                                                                         and len(str(x.дата_выдачи_факт))<5 
                                                                                         and str(x.в_ар_хив).lower().strip()=='да' else x.дата_справки_счет_факт), axis=1)
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.pravka_spravki_schet_i_vidachy.__name__}')
            self.df_np_auto = self.df_np_auto
            
    def proverka_vidachy_i_spravki_schet(self):
        """если есть выдача но нет даты справчки счет
        """
        try:
            if self.name_object_class not in ('OVP_YAR', 'OVP__YAR', 'OVP__RYB', 'OVP_SAR', 'OVP_vved_MSK'):
                df_except = copy.deepcopy(self.df_np_auto)
                df_except['ошибка'] = self.df_np_auto.apply(lambda x: (f'{x.vin} дата выдачи {x.дата_выдачи_факт} даты справки счет нет {x.дата_справки_счет_факт} авто выдан - оплачено {x.получено_за_ам_руб}' 
                                                                    if len(str(x.дата_выдачи_факт))>5 and len(str(x.дата_справки_счет_факт))<5  else None), axis=1)
                df_except = df_except[df_except['ошибка'].notna()]
                df_except = df_except[self.white_list_columns_except_logist]
                self.except_kum = pd.concat([self.except_kum ,df_except])
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.proverka_spravki_schet_i_vidachy.__name__}')
            None   
            
    def proverka_ceny_prodajy_i_daty_prodajy(self):
        """если есть цена_продажи но нет дата_продажи_факт
        """
        try:
            df_except = copy.deepcopy(self.df_sclad)
            df_except['ошибка'] = self.df_sclad.apply(lambda x: (f'{x.vin} цена продажи {x.цена_продажи} нет даты продажи {x.дата_продажи_факт}' 
                                                                if float(x.цена_продажи)>0 and (len(str(x.дата_продажи_факт))<5 if 'NaT' not in str(x.дата_продажи_факт) else True)  else None), axis=1)
            df_except = df_except[df_except['ошибка'].notna()]
            df_except = df_except[self.white_list_columns_except_logist]
            self.except_kum = pd.concat([self.except_kum ,df_except])
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.proverka_ceny_prodajy_i_daty_prodajy.__name__}')
            None   

    
    # ПОЗЖЕ УДАЛИТЬ причастные функции
    
    def individual_statvs_zakaza_VARSH_BAIK_UKA_HYUNDAI(self):
        """индивидуально только для _VARSH_BAIK_UKA_HYUNDAI так как у них нет NP
            и статусы склада свои, подгоняем под общий стандарт
        """
        try:
            if  self.name_object_class in ['BAIC_varsh_MSK', 'HYUNDAI_varsh_MSK', 'UKA_varsh_MSK']:
                self.df_np_auto['склад_заказ'] = self.df_np_auto['склад_заказ'].apply(lambda x: status_zakaza_VARSH_BAIK_UKA_HYUNDAI(x))
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.individual_statvs_zakaza_VARSH_BAIK_UKA_HYUNDAI.__name__}')
             
            
    def idividual_prauka_KIA_varsh_status_sklad(self):
        try:
            if self.name_object_class == 'KIA_vved_MSK':
                self.df_np_auto['склад_заказ'] = self.df_np_auto['склад_заказ'].apply(lambda x: pravka_statysa_KIA_(x))
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.idividual_prauka_KIA_varsh_status_sklad.__name__}')
             
            
    def individual_pravka_statvs_zakaza_OVP(self):
        """для всех объектов OVP применяем статус - 'на складе' и приравнивает 'дата_оплаты_счета' =  'дата_прихода_на_склад'
        """
        if 'OVP_' in self.name_object_class:
            self.df_np_auto['склад_заказ'] = 'на складе'
            self.df_sclad['дата_оплаты_счета'] = self.df_sclad['дата_прихода_на_склад']
            
    def individual_pravka_KIA_vved_MSK_drop_duplicates(self):
        if self.name_object_class == 'KIA_vved_MSK':
            self.df_sclad.drop_duplicates(subset=['vin', 'дата_продажи_факт', 'дата_прихода_на_склад'], inplace=True)
          
    
    def statys_zakaza_nan_(self):
        """прсоевряет стутус авто, соолбец склад_заказ на наличие nan - если находит,
        приводит в соответствие
        """
        try:
            self.df_np_auto['склад_заказ'] = self.df_np_auto.apply(lambda x: (status_zakaza_po_date(x.дата_заказа, x.дата_прихода_на_склад) if str(x.склад_заказ) == 'nan' else x.склад_заказ), axis=1)
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.statys_zakaza_nan_.__name__}')
            
    
    def except_column_korrekt(self):
        """функция для обработки таблицы с ошибками
            убирает ошибки на листах в работе, склад - так как они рабочие
        """
        try:
            self.except_kum = exception_result_korrekt(self.except_kum, 'с_листа')
        except Exception as ex_:
            print(f'{self.name_object_class} ошибка {ex_} функция - {self.except_column_korrekt.__name__}')    
    

    def save_object_class_excel(self):
        """функция сохранения промежуточного объекта класса с тремя собранными фреймами ОПЛАТА, АВТО, СКЛАД
        """
        # записываем в файл и если лист существует, то меняем лист
        with pd.ExcelWriter(rf'{links_main(fr"{script_dir}/file_links.txt", "save_file_predobrabotka")}\{self.name_object_class}.xlsx', 
                            engine='xlsxwriter', date_format = 'dd.mm.yyyy', datetime_format='dd.mm.yyyy') as writer:
            try:
                self.df_np_oplata.to_excel(writer, 'oplata')
                self.df_np_auto.to_excel(writer, 'auto')
                self.df_sclad.to_excel(writer, 'sclad')
                self.except_kum.to_excel(writer, 'except_kum')
            except Exception as ex_:
                print(f'Ошибка при сохранении в Excel {self.name_object_class} {ex_}')
            None
    
        
    def fnc_auto(self):
        """функция запуска функций
        при ручной проврке и отключении не забывать отключать сохраниение листов в функции save_object_class_excel
        """
        if self.starter:
            self.excepts_date()
            self.individual_pravka_statvs_zakaza_OVP()
            self.korrektirovka()
            self.pravka_type_dataframe()
            self.kostraciva_po_date()          
            self.proverka_np_date()
            self.pravka_formy_oplaty()
            self.proverka_otkaza()
            self.pravka_otkaza()
            self.proverka_otkaza_arhiva()
            self.pravka_otkaza_arhiva()
            # self.proverka_daty_prihoda_mejdy_sclad_i_np()
            self.proverka_sklada_na_daty_oplaty_i_ceny_prodajy()
            self.proverka_daty_oplaty()
            self.proverka_date_oplaty_na_min_date_prihoda()
            self.proverka_date_prihoda_na_mean_date_prihoda()
            self.individual_pravka_KIA_vved_MSK_drop_duplicates()
            self.proverka_spravki_schet_i_vidachy()
            self.proverka_vidachy_i_spravki_schet()
            self.proverka_ceny_prodajy_i_daty_prodajy()
            self.pravka_spravki_schet_i_vidachy()
            self.proverka_daty_prihoda()
            self.individual_statvs_zakaza_VARSH_BAIK_UKA_HYUNDAI()
            self.idividual_prauka_KIA_varsh_status_sklad()
            self.statys_zakaza_nan_()
            self.proverka_daty_oplaty_scheta_na_sklade() # тест
            self.proverka_daty_prihoda_na_sclad()
            self.proverka_daty_contrakta_na_skale()
            self.proverka_daty_prodajy_na_skale()
            self.except_column_korrekt()
            
        if self.save_excel:
        # запуск завершающих функций
            self.save_object_class_excel()
        


# наполняем словарь базами данных создавая экземпляры класса
catalog_df_predobrabotka = {} # словарь со всеми базами

count_bd = len(catalog_df.keys())
for i in catalog_df.keys():
    try:
        print(f'{i}-----------------------------------')
        LOG_inf(f'Создаем объект класса {Manufacturing_df_predobrabotka.__name__} {i}', 'INFO')
        catalog_df_predobrabotka[i] = Manufacturing_df_predobrabotka(catalog_df[i], i, True, True)
        count_bd-=1
        print(f'Осталось создать {count_bd} объектов класса')
    except Exception as ex_:
        print(f'{ex_}')
        LOG_inf(f'Не удалось создать объект класса {Manufacturing_df_predobrabotka.__name__} {i}', 'ERROR', [ex_])
        
        
        
class Manufacturing_df_oborotka:
    def __init__(self, name_object_class_Manufacturing_df_predobrabotka, object_class_Manufacturing_df_predobrabotka, starter=None, ignore_matadata=None):
        """класс заполнения листов оборотки на основании предобратанных объектов класса Manufacturing_df_sborka

        Args:
            name_object_class_Manufacturing_df (_type_): имя объекта класса
            df_oborotka_shablon (_type_): пустой шаблон df оборотки для заполнения
            object_class_Manufacturing_df (_type_): предобработанный объект класса Manufacturing_df_sborka
            ignore_matadada : True or False - при флаге True в классе принудительно запустится обновление всех объектов в обход матаданных
        """
        self.days_ago = 10                                                                                                                             # на сколько дней назад окатываемся для проверки обновления
        self.name_object_class = copy.deepcopy(name_object_class_Manufacturing_df_predobrabotka)                                                       # имя объекта класса Manufacturing_df_sborka 
        self.object_class = copy.deepcopy(object_class_Manufacturing_df_predobrabotka)                                                                 # инициализируем предобработанный объект класса Manufacturing_df_sborka 
        self.df_np_oplata = self.object_class.df_np_oplata
        self.df_np_auto = self.object_class.df_np_auto
        self.df_sclad = self.object_class.df_sclad
        self.except_kum = self.object_class.except_kum
        self.date_update =  self.object_class.date_update
        # проеверяем метаданные файлов объекта если дата больше вчера (сегодня минус self.days_ago)
        self.update_oborotka = self.object_class.date_update>=yesterday(self.days_ago)
        self.ignore_matadata = ignore_matadata # принудительный запуск обновлений в обход метаданных (не нужно смотреть старый ли файл) проставить признаку True в классе
        self.update_oborotka = True if self.ignore_matadata == True else self.update_oborotka           # проверка статуса игнорирования медатады
        print(f'{self.name_object_class} дата обновления файлов объекта {self.object_class.date_update} дата на вчера {yesterday(self.days_ago)} - обновление? {self.update_oborotka}')
        # если метаданные файлов объекта старше вчера то считываем данные оборотки из архива, в противном случае считаем заново
        # инициализируем шаблон заполнения оборотки, предварительно считав MIN даты для создания
        self.df_oborotka = df_oborotka_shablon(*min_date_test(copy.deepcopy(object_class_Manufacturing_df_predobrabotka))) if self.update_oborotka==True else read_file_arhiv(self.name_object_class)
        self.starter = starter
        self.fnc_auto()

        
    # # заказы
    def zakazy_st_1(self):
        """считает заказы
        """
        if self.update_oborotka==True:
            try:
                self.df_oborotka['зкз_кред'] = self.df_oborotka.apply(lambda x: zakazy(self.df_np_auto, x.календарь, 'дата_заказа', 'кре', 'форма_оплаты'), axis=1)
                self.df_oborotka['зкз_нал'] = self.df_oborotka.apply(lambda x: zakazy(self.df_np_auto, x.календарь, 'дата_заказа', 'нал', 'форма_оплаты'), axis=1)
                # self.df_oborotka['зкз_путь_кред'] = self.df_oborotka.apply(lambda x: zakazy(self.df_np_auto, x.календарь, 'дата_заказа', 'кре', 'форма_оплаты', 'в пути', 'склад_заказ'), axis=1)
                # self.df_oborotka['зкз_путь_нал'] = self.df_oborotka.apply(lambda x: zakazy(self.df_np_auto, x.календарь, 'дата_заказа', 'нал', 'форма_оплаты', 'в пути', 'склад_заказ'), axis=1)
                self.df_oborotka['зкз_путь_кред'] = self.df_oborotka.apply(lambda x: zakazy_vid_oplaty(self.df_np_auto, x.календарь, 'дата_заказа', 'в пути', 'склад_заказ', 'форма_оплаты', 'кре'), axis=1)
                self.df_oborotka['зкз_склад_кред'] = self.df_oborotka.apply(lambda x: zakazy_vid_oplaty(self.df_np_auto, x.календарь, 'дата_заказа', 'на складе', 'склад_заказ', 'форма_оплаты', 'кре'), axis=1)
                self.df_oborotka['зкз_путь_нал'] = self.df_oborotka.apply(lambda x: zakazy_vid_oplaty(self.df_np_auto, x.календарь, 'дата_заказа', 'в пути', 'склад_заказ', 'форма_оплаты', 'нал'), axis=1)
                self.df_oborotka['зкз_склад_нал'] = self.df_oborotka.apply(lambda x: zakazy_vid_oplaty(self.df_np_auto, x.календарь, 'дата_заказа', 'на складе', 'склад_заказ', 'форма_оплаты', 'нал'), axis=1)
            except Exception as ex_:
                print(f'{self.name_object_class} ошибка {ex_} функция - {self.zakazy_st_1.__name__}')
            
    # отказы
    def otkazy_st_2(self):    
        """считает отказы
        """
        if self.update_oborotka==True:
            try:
                self.df_oborotka['откз_кред'] = self.df_oborotka.apply(lambda x: otkazy(self.df_np_auto, x.календарь, 'дата_изм', 'кре', 'форма_оплаты', 'да', 'в_ар_хив'), axis=1)
                self.df_oborotka['откз_нал'] = self.df_oborotka.apply(lambda x: otkazy(self.df_np_auto, x.календарь, 'дата_изм', 'нал', 'форма_оплаты', 'да', 'в_ар_хив'), axis=1)
                # self.df_oborotka['откз_путь_кред'] = self.df_oborotka.apply(lambda x: otkazy(self.df_np_auto, x.календарь, 'дата_изм', 'кре', 'форма_оплаты', 'в пути', 'склад_заказ', 'да', 'в_ар_хив'), axis=1)
                # self.df_oborotka['откз_путь_нал'] = self.df_oborotka.apply(lambda x: otkazy(self.df_np_auto, x.календарь, 'дата_изм', 'нал', 'форма_оплаты', 'в пути', 'склад_заказ', 'да', 'в_ар_хив'), axis=1)
            except Exception as ex_:
                print(f'{self.name_object_class} ошибка {ex_} функция - {self.otkazy_st_2.__name__}')
            
            
    # выдачи
    def vidachy_st_3(self):
        """считает выдачи
        """
        if self.update_oborotka==True:
            try:
                self.df_oborotka['выдачи_кред'] = self.df_oborotka.apply(lambda x: kolichestyo_vidach(self.df_sclad, x.календарь, 'дата_продажи_факт', 'кре', 'форма_оплаты'), axis=1)
                self.df_oborotka['выдачи_нал'] = self.df_oborotka.apply(lambda x: kolichestyo_vidach(self.df_sclad, x.календарь, 'дата_продажи_факт', 'нал', 'форма_оплаты'), axis=1)
                self.df_oborotka['выдачи_всего'] = self.df_oborotka.apply(lambda x: (x.выдачи_кред + x.выдачи_нал), axis=1)
            except Exception as ex_:
                print(f'{self.name_object_class} ошибка {ex_} функция - {self.vidachy_st_3.__name__}')
        
    # всего заказов с учетом отказов и выдач
    # если что-то не идет накопительно значит в NP столбец Дата изм есть отказные авто с пустотами - нужно нагибать логистов (дописать проверку) !!!!!!!!!!
    def vsego_zakazov_s_vchetom_otkazov_st_4(self):
        """всего заказов с учетом отказов
        """
        if self.update_oborotka==True:
            try:
                self.df_oborotka['всего_зкз_с_уч_откз_и_выд_кред'] = self.df_oborotka.apply(lambda x: zakazy_s_vchetom_okazov_i_vidach(self.df_oborotka, x.календарь, 
                                                                                                                                    'календарь', 'зкз_кред', 
                                                                                                                                    'откз_кред', 'выдачи_кред'), axis=1)
                self.df_oborotka['всего_зкз_с_уч_откз_и_выд_нал'] = self.df_oborotka.apply(lambda x: zakazy_s_vchetom_okazov_i_vidach(self.df_oborotka, x.календарь, 
                                                                                                                                    'календарь', 'зкз_нал', 
                                                                                                                                    'откз_нал', 'выдачи_нал'), axis=1)
                self.df_oborotka['всего_зкз_с_уч_откз_и_выд_всего'] = self.df_oborotka.apply(lambda x: (x.всего_зкз_с_уч_откз_и_выд_кред + x.всего_зкз_с_уч_откз_и_выд_нал), axis=1)
            except Exception as ex_:
                print(f'{self.name_object_class} ошибка {ex_} функция - {self.vsego_zakazov_s_vchetom_otkazov_st_4.__name__}')
            
            
            
            
    # фин показатели
    def fin_pokazately_st_5(self):
        """фин показатели
        """
        if self.update_oborotka==True:
            try:
                self.df_oborotka['выдачи_выручка'] = self.df_oborotka.apply(lambda x: sum_finance_day(self.df_sclad, x.календарь, 'дата_продажи_факт', 'цена_продажи'), axis=1)
                self.df_oborotka['выдачи_себестоимость'] = self.df_oborotka.apply(lambda x: sum_finance_day(self.df_sclad, x.календарь, 'дата_продажи_факт', 'себестоимость_ам'), axis=1)
            except Exception as ex_:
                print(f'{self.name_object_class} ошибка {ex_} функция - {self.fin_pokazately_st_5.__name__}')
            
    # показатели накопительно
    def pokazately_nakopitelno_st_6(self):
        """показатели накопительно
        """
        if self.update_oborotka==True:
            try:
                self.df_oborotka['продано_ам_накоп'] = self.df_oborotka.apply(lambda x: sum_finance_day_nakopitelno(self.df_oborotka, x.календарь, 'календарь', 'выдачи_всего'), axis=1)
                self.df_oborotka['выручка_накоп'] = self.df_oborotka.apply(lambda x: sum_finance_day_nakopitelno(self.df_oborotka, x.календарь, 'календарь', 'выдачи_выручка'), axis=1)
                self.df_oborotka['себестоимость_накоп'] = self.df_oborotka.apply(lambda x: sum_finance_day_nakopitelno(self.df_oborotka, x.календарь, 'календарь', 'выдачи_себестоимость'), axis=1)
                self.df_oborotka['наценка'] = self.df_oborotka.apply(lambda x: nacenka(self.df_oborotka, x.календарь, 'календарь', 'выручка_накоп', 'себестоимость_накоп'), axis=1)
            except Exception as ex_:
                print(f'{self.name_object_class} ошибка {ex_} функция - {self.pokazately_nakopitelno_st_6.__name__}')
            
        
    # приход а-м сводные/клиентские
    def prihod_auto_st_7(self):
        """приход авто 
        """
        if self.update_oborotka==True:
            try:
                self.df_oborotka['приход_ам_своб'] = self.df_oborotka.apply(lambda x: prihod_auto(self.df_sclad, x.календарь, 'дата_прихода_на_склад', 'дата_контракта_заказа', 'svobod'), axis=1)
                self.df_oborotka['приход_ам_клиент'] = self.df_oborotka.apply(lambda x: prihod_auto(self.df_sclad, x.календарь, 'дата_прихода_на_склад', 'дата_контракта_заказа', 'klient'), axis=1)
            except Exception as ex_:
                print(f'{self.name_object_class} ошибка {ex_} функция - {self.prihod_auto_st_7.__name__}')
            
        
    # авто на складе
    def auto_na_sclade_st_8(self):
        if self.update_oborotka==True:
            try:
                self.df_oborotka['ам_на_складе_своб'] = self.df_oborotka.apply(lambda x: auto_na_sclade(self.df_sclad, x.календарь, 'дата_прихода_на_склад', 'дата_продажи_факт', 'дата_контракта_заказа', 'sclad'), axis=1)
                self.df_oborotka['ам_на_складе_клиент'] = self.df_oborotka.apply(lambda x: auto_na_sclade(self.df_sclad, x.календарь, 'дата_прихода_на_склад', 'дата_продажи_факт', 'дата_контракта_заказа', 'klient'), axis=1)
                self.df_oborotka['склад_всего_ам'] = self.df_oborotka.apply(lambda x: auto_na_sclade(self.df_sclad, x.календарь, 'дата_прихода_на_склад', 'дата_продажи_факт', 'дата_контракта_заказа', 'all'), axis=1)
                self.df_oborotka['склад_в_тч_демо_ам'] = self.df_oborotka.apply(lambda x: auto_na_sclade(self.df_sclad, x.календарь, 'дата_прихода_на_склад', 'дата_продажи_факт', 'дата_контракта_заказа', 'demo'), axis=1)
                self.df_oborotka['склад_конс_ам'] = self.df_oborotka.apply(lambda x: auto_na_sclade_consignacia(self.df_sclad, x.календарь), axis=1)
            except Exception as ex_:
                print(f'{self.name_object_class} ошибка {ex_} функция - {self.auto_na_sclade_st_8.__name__}')
            
        
    def auto_u_puti_st_9(self):
        if self.update_oborotka==True:
            try:
                self.df_oborotka['ам_в_пути_выкуп'] = self.df_oborotka.apply(lambda x: auto_u_puti_vikuplenie(self.df_sclad, x.календарь), axis=1)
            except Exception as ex_:
                print(f'{self.name_object_class} ошибка {ex_} функция - {self.auto_u_puti_st_9.__name__}')
        
        
    # оплаты
    def oplaty_st_10(self):
        if self.update_oborotka==True:
            try:
                self.df_oborotka['оплаты'] = self.df_oborotka.apply(lambda x: oplaty(self.df_np_oplata, x.календарь), axis=1)
            except Exception as ex_:
                print(f'{self.name_object_class} ошибка {ex_} функция - {self.oplaty_st_10.__name__}')
            
            
    def platejy_st_11(self):
        if self.update_oborotka==True:
            try:
                self.df_oborotka['платежи_ам_клиент_шт'] = self.df_oborotka.apply(lambda x: platejy(self.df_sclad, x.календарь, 'count_', 'klient'), axis=1)
                self.df_oborotka['платежи_ам_клиент_руб'] = self.df_oborotka.apply(lambda x: platejy(self.df_sclad, x.календарь, 'sum_', 'klient'), axis=1)
                self.df_oborotka['платежи_ам_всего_шт'] = self.df_oborotka.apply(lambda x: platejy(self.df_sclad, x.календарь, 'count_', 'sclad'), axis=1)
                self.df_oborotka['платежи_ам_всего_руб'] = self.df_oborotka.apply(lambda x: platejy(self.df_sclad, x.календарь, 'sum_', 'sclad'), axis=1)
                self.df_oborotka['платежи_ам_свободн_шт'] = self.df_oborotka.apply(lambda x: (x.платежи_ам_всего_шт - x.платежи_ам_клиент_шт), axis=1)
                self.df_oborotka['платежи_ам_свободн_руб'] = self.df_oborotka.apply(lambda x: (x.платежи_ам_всего_руб - x.платежи_ам_клиент_руб), axis=1)
            except Exception as ex_:
                print(f'{self.name_object_class} ошибка {ex_} функция - {self.platejy_st_11.__name__}')
    

    def oborotnie_sredstya_st_13(self):
        if self.update_oborotka==True:
            try:
                self.df_oborotka['оборот_средства_без_демо'] = self.df_oborotka.apply(lambda x: oborotnie_sredstya(self.df_sclad, x.календарь, 'not_demo'), axis=1)
                self.df_oborotka['оборот_средства_без_демо_на_скл'] = self.df_oborotka.apply(lambda x: oborotnie_sredstya(self.df_sclad, x.календарь, 'not_demo_na_sclade'), axis=1)
                self.df_oborotka['оборот_средства_демо'] = self.df_oborotka.apply(lambda x: oborotnie_sredstya(self.df_sclad, x.календарь, 'demo'), axis=1)
            except Exception as ex_:
                print(f'{self.name_object_class} ошибка {ex_} функция - {self.oborotnie_sredstya_st_13.__name__}')
            

    def proverka_oborotnih_sredsty_st_14(self):
        if self.update_oborotka==True:
            try:
                self.df_oborotka['проверка'] = self.df_oborotka.apply(lambda x: (proverka_oborotnih_sredsty(self.df_sclad, x.календарь)-x.оборот_средства_без_демо), axis=1)
            except Exception as ex_:
                print(f'{self.name_object_class} ошибка {ex_} функция - {self.proverka_oborotnih_sredsty_st_14.__name__}')
            
            
    def dop_informaciva_15(self):
        """добавляет доп столбцы с информацией названий/регионов
        """
        if self.update_oborotka==True:
            try:
                self.df_oborotka['имя_объекта'] = self.name_object_class
            except Exception as ex_:
                print(f'{self.name_object_class} ошибка {ex_} функция - {self.dop_informaciva.__name__}')
            
    def region_marka_16(self):
        """добавляем столбец регион в лист оборотки
        """
        if self.update_oborotka==True:
            try:
                self.df_oborotka['марка'] = self.name_object_class.split('_')[0]
                self.df_oborotka['регион'] = self.name_object_class.split('_')[-1]
                
            except Exception as ex_:
                print(f'{self.name_object_class} ошибка {ex_} функция - {self.region_marka_16.__name__}')
                
    def update_arhiv_oborotka_17(self):
        """в считанной архивной оборотке добавляем строки с датами 
        и протягиваем накопительные колонки
        """
        if self.update_oborotka==False:
            try:
                self.df_oborotka = protajka_stolbcov_v_arhivnoy_oborotke(self.df_oborotka)
            except Exception as ex_:
                print(f'{self.name_object_class} ошибка {ex_} функция - {self.update_arhiv_oborotka_17.__name__}')
            
            
    # def bonus_park_18(self):
    #     """цепляет бонусы из ПАРКА
    #     """
    #     try:
    #         self.df_oborotka['бонус'] = self.df_oborotka.apply(lambda x: (bonus_park(x.календарь, convertor_brands_in_PARK(x.марка, x.регион)[0], convertor_brands_in_PARK(x.марка, x.регион)[1], 'Доход, руб.') if x.календарь.day == 1 else None), axis=1)
            
    #     except Exception as ex_:
    #         print(f'{self.name_object_class} ошибка {ex_} функция - {self.bonus_park_18.__name__}')
        
        
    def save_object_class_excel(self):
        
        """функция сохранения промежуточного объекта класса с тремя собранными фреймами ОПЛАТА, АВТО, СКЛАД
        """
        try:
            # записываем в файл и если лист существует, то меняем лист
            with pd.ExcelWriter(rf'{links_main(fr"{script_dir}/file_links.txt", "save_file_oborotka")}\{self.name_object_class}.xlsx', 
                                engine='xlsxwriter', date_format = 'dd.mm.yyyy', datetime_format='dd.mm.yyyy') as writer:
            # Записать ваш DataFrame в файл на листы
                self.df_np_oplata.to_excel(writer, 'oplata')    # отключить сохранение после завершения проекта оставить только для 'oborotka'
                self.df_np_auto.to_excel(writer, 'auto')        # отключить сохранение после завершения проекта оставить только для 'oborotka'
                self.df_sclad.to_excel(writer, 'sclad')         # отключить сохранение после завершения проекта оставить только для 'oborotka'
                self.except_kum.to_excel(writer, 'except_kum')
                self.df_oborotka.to_excel(writer, 'oborotka')
        except Exception as ex_:
                print(f'{self.name_object_class} ошибка {ex_} функция - {self.save_object_class_excel.__name__} видимо файл открыт')
            
    def save_object_class_excel_exception(self):
        
        """функция сохранения  объекта класса только с ошибками
        """
        try:
            # записываем в файл и если лист существует, то меняем лист
            with pd.ExcelWriter(rf'{links_main(fr"{script_dir}/file_links.txt", "save_file_exception")}\{self.name_object_class}.xlsx', 
                                engine='xlsxwriter', date_format = 'dd.mm.yyyy', datetime_format='dd.mm.yyyy') as writer:
            # Записать ваш DataFrame в файл на листы
                self.except_kum.to_excel(writer, 'except')
        except Exception as ex_:
                print(f'{self.name_object_class} ошибка {ex_} функция - {self.save_object_class_excel_exception.__name__} видимо файл открыт')

    
    
    def fnc_auto(self):
        """функция запуска функций
        при ручной проврке и отключении не забывать отключать сохраниение листов в функции save_object_class_excel
        """
        if self.starter:
            self.zakazy_st_1()
            self.otkazy_st_2()
            self.vidachy_st_3()
            self.vsego_zakazov_s_vchetom_otkazov_st_4()
            self.fin_pokazately_st_5()
            self.pokazately_nakopitelno_st_6()
            self.prihod_auto_st_7()
            self.auto_na_sclade_st_8()
            self.auto_u_puti_st_9()
            self.oplaty_st_10()
            self.platejy_st_11()
            self.oborotnie_sredstya_st_13()
            self.proverka_oborotnih_sredsty_st_14()
            self.dop_informaciva_15()
            self.region_marka_16()
            self.update_arhiv_oborotka_17()

            self.save_object_class_excel()
            self.save_object_class_excel_exception()
            
            
            
catalog_df_oborotka = {} # словарь со всеми базами

counter_update = []
counter_not_update = []
count_bd = len(catalog_df_predobrabotka.keys())
for i in catalog_df_predobrabotka.keys():
    print(f'{i}-----------------------------------')
    LOG_inf(f'Создание объекта класса {Manufacturing_df_oborotka.__name__} {i}', 'INFO')
    catalog_df_oborotka[i] = Manufacturing_df_oborotka(i, catalog_df_predobrabotka[i], True, False)
    count_bd-=1
    if catalog_df_oborotka[i].update_oborotka:
        counter_update.append(i)
    else:
        counter_not_update.append(i)
    print(f'Осталось создать {count_bd} объектов класса')
    
print(f'Обновлено объектов {len(counter_update)} {counter_update}')
print(f'Не обнавлено объектов {len(counter_not_update)} {counter_not_update}')
LOG_inf(f'Обновлено объектов {len(counter_update)} {counter_update}', 'INFO')
LOG_inf(f'Не обнавлено объектов {len(counter_not_update)} {counter_not_update}', 'INFO')


#конкатинируем в одну БД
LOG_inf(f'Конкатинируем оборотку в одну БД', 'INFO')
try: 
    result_svod = pd.concat([catalog_df_oborotka[i].df_oborotka for i in catalog_df_oborotka.keys()])
except Exception as ex_:
    print(ex_)
    LOG_inf(f'Конкатинируем оборотку в одну БД', 'ERROR', ex_)
    
    
# блок бонусов из ПАРКА цепляем к оборотке
LOG_inf(f'Цепляем бонусы из ПАРКА', 'INFO')
try:
    df_step_1 = PARK.copy()
    df_step_1 = df_step_1[df_step_1['ТИП'] == 'Бонус'][['мес' , 'Подразделение/площадка', 'Марка', 'Бонус']]
    df_step_1 = df_step_1[abs(df_step_1['Бонус']) > 0]
    df_step_1 = df_step_1.merge(CONNECTION_BRAND_PARK, left_on=['Подразделение/площадка', 'Марка'], right_on=['подразделение', 'марка'], how='left')[['мес' , 'марка_фильтр',  'регион_фильтр' , 'Бонус',]]
    df_step_1 = df_step_1.rename(columns={'мес':'календарь', 'марка_фильтр':'марка', 'регион_фильтр':'регион'})
    result_svod = pd.concat([result_svod, df_step_1])
except Exception as ex_:
    print(ex_)
    LOG_inf(f'Цепляем бонусы из ПАРКА', 'ERROR', ex_)
    
    
# блок планов цепляем к оборотке
LOG_inf(f'Цепляем ПЛАНЫ', 'INFO')
try:
    df_step_2 = PLAN_AUTO.copy()
    df_step_2['календарь'] = df_step_2.apply(lambda x: (individ_date_plan(x.year, x.mnth)), axis=1)
    df_step_2['календарь'] = pd.to_datetime(df_step_2['календарь'])
    df_step_2 = df_step_2[df_step_2['type_ind'] == 'Авто'][['календарь' , 'reg', 'item_ind', 'zone','ПЛН']]
    df_step_2 = df_step_2[abs(df_step_2['ПЛН']) > 0]
    df_step_2 = df_step_2.merge(CONNECTION_BRAND_PLAN_AUTO, how='left')[['календарь','марка_фильтр',  'регион_фильтр', 'ПЛН']]
    df_step_2 = df_step_2.rename(columns={'марка_фильтр':'марка', 'регион_фильтр':'регион'})
    result_svod = pd.concat([result_svod, df_step_2])
except Exception as ex_:
    print(ex_)
    LOG_inf(f'Цепляем ПЛАНЫ', 'ERROR', ex_)
    
    
for i in catalog_df_predobrabotka.keys():
    try:
        res = [i for i in  catalog_df_predobrabotka[i].df_np_auto['склад_заказ'].unique() if i not in ['на складе', 'в пути']]
        if len(res)>0:
            print(f'Проверка значений в столбце склад_заказ отличных от [на складе / в пути]', i, res)
    except Exception as ex_:
        print(f'{i} нет столбца')

# обрезаем лишнее
result_svod = result_svod.dropna(subset=['марка', 'регион']) 

# сохраняем оборотку
LOG_inf(f'сохраняем оборотку', 'INFO')
try:
    result_svod.to_excel(links_main(fr"{script_dir}/file_links.txt", "save_sborka_oborotka"))
except Exception as ex_:
    print(ex_)
    LOG_inf(f'сохраняем оборотку', 'ERROR', ex_)
 
 
# собираем и сохраняем склады и np c новыми ключами и сохраняем, нужно Ж.Р.А.

LOG_inf(f'собираем и сохраняем склады и np c новыми ключами и сохраняем, нужно Ж.Р.А.', 'INFO')
try:
    all_skl = []
    all_np = []
    for i in catalog_df_oborotka.keys():
        skl = copy.deepcopy(catalog_df_oborotka[i].df_sclad)
        np = copy.deepcopy(catalog_df_oborotka[i].df_np_auto)
        skl['key_new'] = i
        np['key_new'] = i
        all_skl.append(skl)
        all_np.append(np)

    result_sclad = pd.concat(all_skl)
    result_np_auto = pd.concat(all_np)
    result_sclad.to_excel(links_main(fr"{script_dir}/file_links.txt", "save_sborka_sclad"))
    result_np_auto.to_excel(links_main(fr"{script_dir}/file_links.txt", "save_sborka_np_auto"))
except Exception as ex_:
    print(ex_)
    LOG_inf(f'сохраняем оборотку', 'ERROR', ex_)
    
    
# обновление сводной таблицы
LOG_inf(f'обновление сводной таблицы', 'INFO')
try:
    update_file(links_main(fr"{script_dir}/file_links.txt", "update_file"))
except Exception as ex_:
    print(ex_)
    LOG_inf(f'обновление сводной таблицы', 'ERROR', ex_)
    
    
# БЛОК РАССЫЛКИ СООБЩЕНИЙ С ОШИБКАМИ ПО ДАННЫМ

# считываем фрейм с адресами почты и связью файлов для рассылки
LOG_inf(f'считывание фрейма для рассылки сообщений с ошибками', 'INFO')
try:
    df_emal_exception = pd.read_excel(links_main(fr"{script_dir}/file_links.txt", "email_exception"))
    df_emal_exception
except Exception as ex_:
    print(ex_)
    LOG_inf(f'считывание фрейма для рассылки сообщений с ошибками', 'ERROR', ex_)
    
# получаем список уникальных объектов рассылки
unique_name_object_email = df_emal_exception['object'].unique() # список уникальных имен объектов


# перебираем уникальные элементы объектов с собранными ошибками и получаем адреса ответственных лиц
LOG_inf(f'рассылаем сообщения с ошибками', 'INFO')

try:
    status_email_flag = True
    pause_sleep = 15
    if status_email_flag:
        for i in unique_name_object_email:
            name = i                                                            # получаем имя объекта
            email_adress_go = return_email_except_df(name, 'object', 'email')   # получаем email
            link = return_link_directory(name ,'save_file_exception')           # получаем ссылку на файл
            df = pd.read_excel(link)                                            # считываем фрейм
            count_srok = df.shape[0]                                            # считываем параметры табл строки/столб
            if count_srok>=1:                                                   # если строк больше 1
                print(email_adress_go, link, name)
                LOG_inf(f'{email_adress_go, link, name}', 'INFO')
                send_mail(email_adress_go, link, name)                          # отправляем почту
                time.sleep(pause_sleep)                                         # без задержки вылетает ошибка отправки почты менее 15 сек

except Exception as ex_:
    print(ex_)
    LOG_inf(f'рассылаем сообщения с ошибками', 'ERROR', ex_)   
    
    
# СРАВНЕНИЕ СКЛАДА ТЕКУЩЕГО С АРХИВНЫМ

LOG_inf(f'сборка результатов сравнения склада архивного и текущего', 'INFO')

try:
    sborka_df = []
    for i in catalog_df_oborotka.keys():

        try:
            print(i)
            df_skl_arh = copy.deepcopy(read_file_arhiv(i, 'sclad', "paste_link_dir"))  # архивный склад
            df_skl = copy.deepcopy(catalog_df_oborotka[i].df_sclad)                    # текущий склад
            df_skl_arh = df_skl_arh.fillna(0)
            df_skl = df_skl.fillna(0)
            for col in df_skl.columns:
                df_skl[col] = df_skl[col].apply(lambda x: 0 if str(x) in ['nan'] else x)
            df_skl_arh['сравнение'] = df_skl_arh.apply(lambda x: (sravnenie_arh_skl_k_tek(x.vin, df_skl_arh, df_skl)), axis=1)
            df_skl_arh = df_skl_arh[['vin','сравнение']]
            df_skl_arh['сравнение'] = df_skl_arh['сравнение'].apply(lambda x: 'удалить' if len(str(x))<=3 else x)
            df_skl_arh = df_skl_arh[~df_skl_arh['сравнение'].str.contains("|".join(['удалить', 'не удалось найти 0', ' было 0 стало nan', 'vin 0 есть но данных нет']), case=False, na=False)] #'не удалось найти 0'
            df_skl_arh['объект'] = i
            df_skl_arh['дата'] = tek_day()
            df_skl_arh['пропал_vin'] = df_skl_arh.apply(lambda x: (x.vin if 'не удалось найти' in str(x.сравнение) else None), axis=1)
            sborka_df.append(df_skl_arh) # добавляем фрейм
            
        except Exception as ex_:
            print(f'ошибка - Не удалось сранвить {i}')
            LOG_inf(f'ошибка - Не удалось сранвить {i}', 'ERROR', ex_)
            
            
    df_sravn_arh = pd.read_excel(links_main(fr"{script_dir}/file_links.txt", "sravnrnie_sclada"))      # считываем передыдущие результаты
    sborka_df.append(df_sravn_arh)
    result_sravnenie_sclada = pd.concat(sborka_df)                                      # объединяем все данные
    result_sravnenie_sclada = result_sravnenie_sclada[[i for i in result_sravnenie_sclada.columns if 'Unnamed: 0' not in i]]
    result_sravnenie_sclada.to_excel(links_main(fr"{script_dir}/file_links.txt", "sravnrnie_sclada"))  # сохраняем результат
    
except Exception as ex_:
    print(ex_)
    LOG_inf(f'сборка результатов сравнения склада архивного и текущего', 'ERROR', ex_)  
    
    
LOG_inf(f'Пророверка обновления сводной таблицы', 'INFO')
try:
    test_udate_file_svod_tab = file_update(links_main(fr"{script_dir}/file_links.txt", 'update_file')) # получааем метаданные сводной таблицы
    result_updatefile_true_or_false = test_udate_file_svod_tab>yesterday(1)
    LOG_inf(f'Метаданные сводной таблицы {test_udate_file_svod_tab} обновлена {result_updatefile_true_or_false}', 'INFO')
except Exception as ex_:
    print(ex_)
    LOG_inf(f'Пророверка обновления сводной таблицы', 'ERROR', ex_)
    
    
# отпрака результатов логирования
send_mail_2(['skrutko@sim-auto.ru', 'zhurin@sim-auto.ru'], 
            links_main(fr"{script_dir}/file_links.txt", 'log'), 
            'log.log', 
            them = 'ОБОРОТКА',
            body='Результат отработки скрипта по оборотке ')
