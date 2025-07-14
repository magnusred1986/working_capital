import pandas as pd
import copy
# блок логирования
import logging
from functools import wraps
import time

import datetime as DT
from datetime import timedelta
import xlrd

import os
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
    
    
logging.basicConfig(level=logging.INFO, filename=fr"{DIR}\py_log_temp_new.log",filemode="w", format="%(asctime)s %(levelname)s %(message)s")
def LOG_inf(name, type_='INFO' or 'ERROR', *args):
    try:
        if type_ == 'INFO': logging.info(f"{name} {args}")
        elif type_ == 'ERROR': logging.error(f"{name} {args}")
    except Exception as ex_:
        print(f'Х_НЯ с логированием {ex_}')


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
        file = pd.read_csv(name_file, sep=';')
        result = list(file[file['ключ']==key]['значение'])[0]
        return result
    except Exception as ex_:
        print(f'ошибка функции {links_main.__name__} не удалось считать файл {name_file} или данные в нем {key} ошибка {ex_}')
        


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
        print(f'{OVP_YAR.__name__} {ex_}')
        
        
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
        print(f'{saratov_marka.__name__} {ex_}')
        

def kiapi_rename(word):
    try:
        if 'KIAPI'== word:
            return 'KIAimp'
        else:
            return word
    except Exception as ex_:
        print(f'{kiapi_rename.__name__} {ex_}')
        
        
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
        print(f'{jetoor_msk.__name__} {ex_}')
        
        
def mazda_msk_next(marka, region, komment ):
    """ делит JETOUR МСК на JETOUR OVP_JETOU

    Args:
        marka (_type_): _description_
        region (_type_): _description_
        komment (_type_): _description_

    Returns:
        _type_: _description_
    """
    try:
        if marka == 'MAZDA' and region=='MSK' and str(komment).strip() == 'б/у':
            return 'OVP_MAZDA'
        else:
            return marka
    except Exception as ex_:
        print(f'{jetoor_msk.__name__} {ex_}')
        
        
def individ_date_plan(year, month):
    try:
        year = str(year)
        month = str(month)
        month = month if len(month)==2 else '0'+month
        day = '01'
        return f'{year}-{month}-{day}'
    except Exception as ex_:
        print(f'Ошибка функции {individ_date_plan.__name__} {ex_} не удалось преобразовать {year}{month}')
        
        
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
        print(f'ошибка функции {yesterday_new.__name__}  {ex_}')
        
        
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
        print(f'ошибка функции {korp_rozn.__name__}  {ex_}')
        
        
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
        print(f'ошибка функции {kre_nal_ovp.__name__}  {ex_}')


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
        time.sleep(60) # задержка 60 секунд, чтоб уж точно обновились сводные wb.RefreshAll() - иначе будет ошибка 
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


def my_pass():
    """функция считывания пароля

    Returns:
        _type_: _description_
    """
    
    try:
        with open(links_main(fr'{DIR}\file_links.txt', 'pass'), 'r') as actual_pass:
            return actual_pass.read()
        
    except Exception as ex_:
        print(f'ошибка функции {my_pass.__name__} {ex_}')


def read_email(link):
    try:
        df_email = pd.read_excel(link)
        res = list(df_email['email'])
        return res
    except Exception as ex_:
        print(f'ошибка функции {read_email.__name__} {ex_} входне параметры {link}')


def send_mail(send_to:list, file_link, file_name, them = '', body=''):
    """рассылка почты

    Args:
        send_to (list): список адресов для рассылки
        file_link(str): ссылка на файл
        file_name(str): имя файла в данном варианте нужно указывать с расширением 'BAIC_MSK.xlsx' 
        them(str) - тема письма
        body(str) - тело письма
        
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
        password = PASSWORD
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
        

LOG_inf(f'считываем данные для обработки', 'INFO')   
print(f'считываем данные для обработки')
try:
    CONNECTION_BRAND_PLAN_AUTO = pd.read_excel(links_main(fr'{DIR}\file_links.txt', 'connection_brand'), sheet_name='PLAN_AUTO')
    PLAN_AUTO = pd.read_excel(links_main(fr'{DIR}\file_links.txt', 'plan_auto'))
    KOSTRACIA = '2023-01-01'
    df = pd.read_excel(links_main(fr'{DIR}\file_links.txt', 'read_file_main'))
    df = df.drop(columns='Unnamed: 0', axis=1)
except Exception as ex_:
        print(f'не удалось считывать данные для обработки {ex_}')
        LOG_inf(f'не удалось считывать данные для обработки', 'ERROR', ex_)


black_list = ['статус_оригинал', 'id', 'vin_novogo', 'model_novogo', 'дата_прихода_на_склад', 
              'дата_полной_оплаты_факт', 'дата_справки_счет_факт','с_листа', 'ссылка', 
              'сотрудник_продал', 'склад_заказ',	'в_ар_хив', 'получено_за_ам_руб']
# df = df[df['принадлежность']!='SCLAD_OMODA_SAR.xlsx'] # какого-то лешего затесался склад SCLAD_OMODA_SAR.xlsx
df = df[[i for i in df.columns if i not in black_list]]
df = df.rename(columns={'дата_изм':'дата_отказа'})
df['дата_отказа'] = pd.to_datetime(df['дата_отказа'], errors='ignore')


df_vidacha = copy.deepcopy(df)
df_zakaz = copy.deepcopy(df)
df_otkaz = copy.deepcopy(df)

LOG_inf(f'создаем дату выдачи', 'INFO')
print(f'создаем дату выдачи')
try:
    column_name_vidacha = 'дата_выдачи_факт'
    df_vidacha = df_vidacha[[sort_columns('дата', column_name_vidacha, i) for i in df_vidacha.columns if sort_columns('дата', column_name_vidacha, i)!=None]]
    df_vidacha['выдача'] = df_vidacha[column_name_vidacha].apply(lambda x: 1 if len(str(x))>5 else 0)
    df_vidacha = df_vidacha.rename(columns={column_name_vidacha:'дата'})
except Exception as ex_:
        print(f'не удалось создать дату выдачи {ex_}')
        LOG_inf(f'не удалось создать дату выдачи', 'ERROR', ex_)


LOG_inf(f'создаем дату заказа', 'INFO')
print(f'создаем дату заказа')
try:
    column_name_zakaz = 'дата_заказа'
    df_zakaz = df_zakaz[[sort_columns('дата', column_name_zakaz, i) for i in df_zakaz.columns if sort_columns('дата', column_name_zakaz, i)!=None]]
    df_zakaz['заказ'] = df_zakaz[column_name_zakaz].apply(lambda x: 1 if len(str(x))>5 else 0)
    df_zakaz = df_zakaz.rename(columns={column_name_zakaz:'дата'})
except Exception as ex_:
        print(f'не удалось создать дату заказа {ex_}')
        LOG_inf(f'не удалось создать дату заказа', 'ERROR', ex_)

LOG_inf(f'создаем дату отказа', 'INFO')
print(f'создаем дату отказа')
try:
    column_name_otkaz = 'дата_отказа'
    df_otkaz = df_otkaz[[sort_columns('дата', column_name_otkaz, i) for i in df_otkaz.columns if sort_columns('дата', column_name_otkaz, i)!=None]]
    df_otkaz['отказ'] = df_otkaz[column_name_otkaz].apply(lambda x: 1 if len(str(x))>5 else 0)
    df_otkaz = df_otkaz.rename(columns={column_name_otkaz:'дата'})
except Exception as ex_:
        print(f'не удалось создать дату отказа {ex_}')
        LOG_inf(f'не удалось создать дату отказа', 'ERROR', ex_)


LOG_inf(f'конкатинируем данные приводим столбцы в порядок (марка регион)', 'INFO')
print(f'конкатинируем данные приводим столбцы в порядок (марка регион)')
try:
    result = pd.concat([df_vidacha, df_zakaz])
    result['день'] = result['дата'].dt.day
    result = result.dropna(subset='дата')
    result['марка'] = result['принадлежность'].apply(lambda x: str(x).split('_')[1] )
    result['регион'] = result['принадлежность'].apply(lambda x: str(x).split('_')[-1].split('.')[0] )
    result['регион'] = result.apply(lambda x: (OVP_YAR(x.марка, x.регион, x.площадка)), axis=1)
    result['марка'] = result.apply(lambda x: (saratov_marka(x.регион, x.марка, x.модель)), axis=1)
    result['марка'] = result.apply(lambda x: (kiapi_rename(x.марка)), axis=1)
    result['марка'] = result.apply(lambda x: (jetoor_msk(x.марка, x.регион, x.комментарий)), axis=1)
    result['марка'] = result.apply(lambda x: (mazda_msk_next(x.марка, x.регион, x.комментарий)), axis=1)
except Exception as ex_:
        print(f'не удалось сконкатинировать данные привести столбцы в порядок (марка регион) {ex_}')
        LOG_inf(f'не удалось сконкатинировать данные привести столбцы в порядок (марка регион)', 'ERROR', ex_)


LOG_inf(f'получаем планы и мерджим их с марками и регионами', 'INFO')
print(f'получаем планы и мерджим их с марками и регионами')
try:
    PLAN_AUTO_2 = copy.deepcopy(PLAN_AUTO)
    PLAN_AUTO_2['календарь'] = PLAN_AUTO_2.apply(lambda x: (individ_date_plan(x.year, x.mnth)), axis=1)
    PLAN_AUTO_2['календарь'] = pd.to_datetime(PLAN_AUTO_2['календарь'])
    PLAN_AUTO_2 = PLAN_AUTO_2[PLAN_AUTO_2['type_ind'] == 'Авто'][['календарь' , 'reg', 'item_ind', 'zone','ПЛН']]
    PLAN_AUTO_2 = PLAN_AUTO_2[abs(PLAN_AUTO_2['ПЛН']) > 0]
    PLAN_AUTO_2 = PLAN_AUTO_2.merge(CONNECTION_BRAND_PLAN_AUTO, how='left')[['календарь','марка_фильтр',  'регион_фильтр', 'ПЛН']]
    PLAN_AUTO_2 = PLAN_AUTO_2.rename(columns={'марка_фильтр':'марка', 'регион_фильтр':'регион', 'календарь':'дата'})
except Exception as ex_:
        print(f'не удалось получить планы или смерджить их с марками и регионами {ex_}')
        LOG_inf(f'не удалось получить планы или смерджить их с марками и регионами', 'ERROR', ex_)


LOG_inf(f'объединяем планы с df', 'INFO')
print(f'объединяем планы с df')
try:
    result_svod = pd.concat([result, PLAN_AUTO_2])
    result_svod['кредит'] = result_svod.apply(lambda x: kre_nal(x.выдача, x.форма_оплаты), axis=1)
    result_svod['кредит'] = result_svod.apply(lambda x: kre_nal_ovp(x.марка, x.кредит, x.комментарий, x.выдача), axis=1)
    result_svod['корпоратив'] = result_svod.apply(lambda x: korp_rozn(x.клиент), axis=1)
except Exception as ex_:
        print(f'не удалось объединbnm планы с df {ex_}')
        LOG_inf(f'не удалось объединbnm планы с df', 'ERROR', ex_)

LOG_inf(f'кастрируем df и сохраняем', 'INFO')
print(f'кастрируем df и сохраняем')
try:
    # обрезаем больше даты кострации и меньше либо равно текущая дата +31 день
    result_svod = result_svod[result_svod['дата']>=KOSTRACIA]
    result_svod = result_svod[result_svod['дата']<=yesterday_new(31, '+')]
    result_svod.to_excel(links_main(fr'{DIR}\file_links.txt', 'save_result'))
except Exception as ex_:
        print(f'не удалось кастрировать df или сохранить {ex_}')
        LOG_inf(f'не удалось кастрировать df или сохранить', 'ERROR', ex_)


LOG_inf(f'обновляем дашборд ТЕМП', 'INFO')
print(f'обновляем дашборд ТЕМП')
try:
    update_file(links_main(fr'{DIR}\file_links.txt', 'uptate_dashboard'))
except Exception as ex_:
        print(f'не удалось обновить дашборд ТЕМП {ex_}')
        LOG_inf(f'не удалось обновить дашборд ТЕМП', 'ERROR', ex_)


LOG_inf(f'получаем данные для отправки почты', 'INFO')
print(f'получаем данные для отправки почты')
try:
    SEND_FROM = links_main(fr'{DIR}\file_links.txt', 'SEND_FROM')
    SERVER = links_main(fr'{DIR}\file_links.txt', 'SERVER')
    PORT = int(links_main(fr'{DIR}\file_links.txt', 'PORT'))
    USER_NAME = links_main(fr'{DIR}\file_links.txt', 'USER_NAME')
    PASSWORD = my_pass()
except Exception as ex_:
        print(f'не удалось получить данные для отправки почты {ex_}')
        LOG_inf(f'не удалось получить данные для отправки почты', 'ERROR', ex_)   


LOG_inf(f'отправляем почту с ТЕМП-ом', 'INFO')
print(f'отправляем почту с ТЕМП-ом')
try:
    send_mail(read_email(links_main(fr'{DIR}\file_links.txt', 'email_adress')), 
            links_main(fr'{DIR}\file_links.txt', 'uptate_dashboard'), 
            'temp.xlsx', 
            them = 'Темпы', 
            body='Во вложении темпы')
except Exception as ex_:
        print(f'не удалось отправить почту с ТЕМП-ом {ex_}')
        LOG_inf(f'не удалось отправить почту с ТЕМП-ом', 'ERROR', ex_) 