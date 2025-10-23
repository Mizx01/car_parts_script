import pandas as pd
import requests
import sys
from bs4 import BeautifulSoup as BS
import re  # Модуль для работы с регулярными выражениями
from fake_useragent import UserAgent
import time
from urllib.parse import unquote
from pathlib import Path
import xlwings as xw   #нужно чтобы достать активный лист


current_dir = Path(__file__).parent 
wb_path = sys.argv[1]    #второй аргумент (путь к файлу) при запуске скрипта через cmd 
first_cell = sys.argv[2]
first_row = int(sys.argv[3])
user_row_number = int(sys.argv[4])
last_row = first_row + user_row_number - 1
user_column = first_cell.split("$")[1]
first_cell = f'{user_column}{first_row}'
last_cell = f'{user_column}{last_row}'

#wb_path = current_dir / 'part.xlsx'
#wb_path = current_dir / 'part_PL.xlsx'

#print(f'first cell:{first_cell} first row:{first_row} user row number:{user_row_number} last_row:{last_row}')


t0 = time.time()

ua = UserAgent()

# Функция для парсинга данных со страницы dexup
def parse_page_dexup(url):
    try:
        agent = {"User-Agent": ua.random}
        page = requests.get(url, headers=agent)
        if page.status_code == 200:
            doc_dexup = BS(page.text, "html.parser")
            
            # Парсинг основного описания
            tags_dexup = doc_dexup.find_all(class_="goodsInfoDescr")
            product_name_dexup = "No data found"
            for tag in tags_dexup:
                product_name_dexup = tag.get_text(strip=True)
            
            # Парсинг массы
            mass_dexup = None
            characteristics_dexup = doc_dexup.find_all(class_="characteristicsListRow")
            for row in characteristics_dexup:
                property_tag = row.find(class_="property")
                if property_tag and 'Масса, кг:' in property_tag.get_text(strip=True):
                    mass_tag = row.find_all('span')[-1]
                    if mass_tag:
                        try:
                            mass_dexup = round(float(mass_tag.get_text(strip=True)), 2)
                        except ValueError:
                            mass_dexup = None
                    break

            # Парсинг материала
            material_dexup = None
            for row in characteristics_dexup:
                property_tag = row.find(class_="property")
                if property_tag and 'Материал:' in property_tag.get_text(strip=True):
                    material_tag = row.find_all('div')[-1]
                    if material_tag:
                        material_dexup = material_tag.get_text(strip=True)
                    break

            return product_name_dexup, mass_dexup, material_dexup
        else:
            return "Failed to retrieve page", None, None
    except Exception as e:
        return f"Error: {e}", None, None

# Функция для парсинга данных со страницы port3


# Функция для очистки артикула (удаление пробелов и символов)
def clean_artikul(artikul):
    return re.sub(r'\W+', '', artikul)  # Убираем все неалфавитно-цифровые символы

# Функция для приведения первой буквы к заглавной
def capitalize_first_letter(text):
    if text:
        return text[0].upper() + text[1:].lower()
    return text

# Словарь для замены марок
brand_replacement = {
    "PEUGEOT / CITROEN" : "PEUGEOT%20CITROEN",
    "MAGNETI MARELLI" : "MAGNETI MARELLI",
    "AUTOMEGA DELLO" : "AUTOMEGA",
    "MAHLE / KNECHT" : "MAHLE",
    "PEUGEOTCITROEN" : "PEUGEOT%20CITROEN",
    "AUTOMEGADELLO" : "AUTOMEGA%20DELLO",
    "KALE RADYATOR" : "KALE",
    "KOLBENSCHMIDT" : "KOLBENSCHMIDT",
    "KALERADYATOR" : "KALE",
    "MAHLE/KNECHT" : "MAHLE",
    "MERCEDESBENZ" : "MERCEDES%20BENZ",
    "VICTOR REINZ" : "VICTOR%20REINZ",
    "VİCTOR REİNZ" : "VICTOR%20REINZ",
    "MANN-FILTER" : "MANN-FILTER",
    "CONTINENTAL" : "CONTINENTAL",
    "EREN BALATA" : "EREN%20BALATA",
    "HBJAKOPARTS" : "JAKOPARTS",
    "MAHLEKNECHT" : "MAHLE",
    "VICTORREINZ" : "VICTOR%20REINZ",
    "VİCTORREİNZ" : "VICTOR%20REINZ",
    "BLUE PRINT" : "Blue%20Print",
    "ERENBALATA" : "EREN%20BALATA",
    "HERTH+BUSS" : "H%2BB%20JAKOPARTS",
    "MANNFILTER" : "MANN%20FILTER",
    "SCHMITZORG" : "SCHMITZ",
    "SSANG YONG" : "SSANG%20YONG",
    "VICTOR REI" : "VICTOR%20REINZ",
    "VİCTOR REİ" : "VICTOR%20REINZ",
    "SSANGYONG" : "SSANG%20YONG",
    "BLUEPRINT" : "BLUEPRINT",
    "EUROREPAR" : "EUROREPAR",
    "GKNLOEBRO" : "GKN",
    "HERTHBUSS" : "H%2BB",
    "KACMAZLAR" : "KAÇMAZLAR",
    "LEMFORDER" : "LEMFORDER",
    "NTN / SNR" : "NTN",
    "VICTORREI" : "VICTOR%20REINZ",
    "VİCTORREİ" : "VICTOR%20REINZ",
    "KRAFTVOLL" : "KRAFTVOLL",
    "EURORAPE" : "EUROREPAR",
    "LMFORDER" : "LEMFORDER",
    "AUTOMEGA" : "AUTOMEGA",
    "BILSTEIN" : "BILSTEIN",
    "EUROREPA" : "EUROREPAR",
    "GOODYEAR" : "GOODYEAR",
    "HYD HOME" : "HYD%20HOME",
    "MBTRUCKS" : "MB%20TRUCKS",
    "PIERBURG" : "PIERBURG",
    "TEKNOROT" : "TEKNOROT",
    "VOLVOORG" : "VOLVO",
    "BILSTEN" : "BILSTEIN",
    "AUGERCE" : "AUGER",
    "CORTECO" : "CORTECO",
    "E.REPAR" : "EUROREPAR",
    "FEDERAL" : "FEDERAL%20MOGUL",
    "FILTRON" : "FILTRON",
    "HYDHOME" : "HYD%20HOME",
    "MARELLI" : "MAGNETI%20MARELLI",
    "MAGNETI" : "MAGNETI%20MARELLI",
    "METELLI" : "METELLI",
    "PSA-PEUG": "PEUGEOT-CITROEN",
    "NISSENS" : "NISSENS",
    "OPTIMAL" : "OPTIMAL",
    "PIEBURG" : "PIERBURG",
    "PURFLUX" : "PURFLUX",
    "PLEKSAN" : "PLEKSAN",
    "SNR-NTN" : "SNR",
    "RENAULT" : "RENAULT",
    "GARRETT" : "GARRETT",
    "V.REINZ" : "VICTOR%20REINZ",
    "KACMAZ" : "KACMAZLAR",
    "KAÇMAZ" : "KACMAZLAR",
    "KOLBEN" : "KOLBENSCHMIDT",
    "KONEKS" : "KONEKS",
    "HENGST" : "HENGST",
    "WAHLER" : "WAHLER",
    "AIRTEX" : "AIRTEX",
    "BREMBO" : "BREMBO",
    "DELPHI" : "DELPHI",
    "ELRING" : "ELRING",
    "EREPAR" : "EUROREPAR",
    "GOETZE" : "GOETZE",
    "EYQUEM" : "EYQUEM",
    "FERODO" : "FERODO",
    "HOLSET" : "HOLSET",
    "MONROE" : "MONROE",
    "NTNSNR" : "NTN%20SNR",
    "OTOSAN" : "OTOSAN",
    "PACCAR" : "PACCAR",
    "PROVIA" : "PROVIA",
    "REPAIR" : "EUROREPAR",
    "TEXTAR" : "TEXTAR",
    "TIRSAN" : "TIRSAN",
    "TITANX" : "TITANX",
    "TOPRAN" : "TOPRAN",
    "VERNET" : "VERNET",
    "VREINZ" : "VICTOR%20REINZ",
    "YENMAK" : "YENMAK",
    "YILMAZ" : "YILMAZ",
    "TURTEL" : "TURTEL",
    "AISIN" : "AISIN",
    "BANDO" : "BANDO",
    "BESER" : "BESER",
    "BOSCH" : "BOSCH",
    "CIFAM" : "CIFAM",
    "CONTI" : "CONTINENTAL",
    "DAYCO" : "DAYCO",
    "DENSO" : "DENSO",
    "DEKAR" : "DEKAR",
    "FACET" : "FACET",
    "GATES" : "GATES",
    "GLYCO" : "GLYCO",
    "IBRAS" : "IBRAS",
    "HELLA" : "HELLA",
    "IVECO" : "IVECO",
    "LUCAS" : "LUCAS",
    "ONPER" : "ONPER",
    "MAHLE" : "MAHLE",
    "MANDO" : "MANDO",
    "MEYLE" : "MEYLE",
    "NURAL" : "NURAL",
    "NÜRAL" : "NÜRAL",
    "OSRAM" : "OSRAM",
    "SAHIN" : "SAHIN",
    "RAPRO" : "RAPRO",
    "REINZ" : "VICTOR%20REINZ",
    "SACHS" : "SACHS",
    "VALEO" : "VALEO",
    "VOLVO" : "VOLVO",
    "WABCO" : "WABCO",
    "BLUE" : "BLUE%20PRINT",
    "FILT" : "FILTRON",
    "BEHR" : "BEHR",
    "BERU" : "BERU",
    "BSCH" : "BOSCH",
    "CAVO" : "CAVO",
    "DOLZ" : "DOLZ",
    "FEBI" : "FEBI",
    "FEBİ" : "FEBI",
    "FORD" : "FORD",
    "HUCO" : "HUCO",
    "KALE" : "KALE",
    "UCEL" : "UC-EL",
    "VALS" : "VALEO",
    "VALE" : "VALEO",
    "VALA" : "VALEO",
    "KAYA" : "KAYA",
    "KING" : "KING",
    "MANN" : "MANN%20FILTER",
    "MARS" : "MARS",
    "MEHA" : "MEHA",
    "MİBA" : "MIBA",
    "TRSN" : "TIRSAN",
    "ONKA" : "ONKA",
    "OPEL" : "OPEL",
    "SWAG" : "SWAG",
    "AIS" : "AISIN",
    "BAN" : "BANDO",
    "BCH" : "BOSCH",
    "BER" : "BERU",
    "BIL" : "BILSTEIN",
    "BLU" : "BLUE%20PRINT",
    "BOS" : "BOSCH",
    "BSH" : "BOSCH",
    "BMW" : "BMW",
    "BRB" : "BREMBO",
    "BRE" : "BREMBO",
    "BRU" : "BERU",
    "BSC" : "BOSCH",
    "CIF" : "CIFAM",
    "CNT" : "CONTINENTAL",
    "CON" : "CONTINENTAL",
    "COR" : "CORTECO",
    "CRT" : "CORTECO",
    "DAY" : "DAYCO",
    "DEL" : "DELPHI",
    "DEN" : "DENSO",
    "DNS" : "DENSO",
    "DOL" : "DOLZ",
    "ELR" : "ELRING",
    "FAC" : "FACET",
    "FBI" : "FEBI",
    "FEB" : "FEBI",
    "FLT" : "FILTRON",
    "FOR" : "FORD",
    "FRD" : "FORD",
    "GAT" : "GATES",
    "GKN" : "GKN",
    "GLY" : "GLYCO",
    "GTS" : "GATES",
    "HEL" : "HELLA",
    "HLL" : "HELLA",
    "HNG" : "HENGST",
    "INA" : "INA",
    "KOL" : "KOLBENSCHMIDT",
    "LEM" : "LEMFORDER",
    "LMF" : "LEMFORDER",
    "MAH" : "MAHLE",
    "MAI" : "RENAULT",
    "MAN" : "MANDO",
    "MHL" : "MAHLE",
    "MON" : "MONROE",
    "MTL" : "METELLI",
    "NGK" : "NGK",
    "NIS" : "NISSENS",
    "NıS" : "NISSENS",
    "NRV" : "NARVA",
    "OPL" : "OPEL",
    "OPT" : "OPTIMAL",
    "OSM" : "OSRAM",
    "OSR" : "OSRAM",
    "PIE" : "PIERBURG",
    "PUR" : "PURFLUX",
    "RAP" : "RAPRO",
    "RNZ" : "VICTOR%20REINZ",
    "SAC" : "SACHS",
    "SCH" : "SACHS",
    "SCS" : "SACHS",
    "SCX" : "SACHS",
    "SHS" : "SACHS",
    "SKF" : "SKF",
    "SNR" : "SNR",
    "SWG" : "SWAG",
    "TEK" : "TEKNOROT",
    "TRW" : "TRW",
    "VAL" : "VALEO",
    "VCT" : "VICTOR%20REINZ",
    "VER" : "VERNET",
    "YEN" : "YENMAK",
    "DYC" : "DAYCO",
    "FAE" : "FAE",
    "FAG" : "FAG",
    "FRJ" : "FIAT",
    "KRF" : "KRAFTVOLL",
    "FMN" : "FMN",
    "FTE" : "FTE",
    "GMB" : "GMB",
    "KLR" : "KALE",
    "KAL" : "KALE",
    "KYB" : "KYB",
    "LUK" : "LUK",
    "MGA" : "AUTOMEGA",
    "MMA" : "MAGNETI%20MARELLI",
    "MND" : "MANDO",
    "MNN" : "MANN-FILTER",
    "MGA" : "MGA",
    "NRF" : "NRF",
    "NTN" : "NTN",
    "OES" : "OES",
    "GNS" : "GUNES",
    "IBR" : "IBRAS",
    "OTO" : "OTO",
    "POJ" : "PEUGEOT%20CITROEN",
    "PRB" : "PIERBURG",
    "PRG" : "PIERBURG",
    "PSA" : "PSA",
    "RYL" : "ROYAL",
    "SKT" : "SKT",
    "SWF" : "SWF",
    "TXT" : "TEXTAR",
    "TUR" : "TURTEL",
    "TIR" : "TIRSAN",
    "UFI" : "UFI",
    "WIN" : "WIN",
    "WOD" : "WOD",
    "YTT" : "YTT",
    "YNM" : "YENMAK",
    "ORJ" : "VAG",
    #"ORJ" : "ORIGINAL",
    #"ORJ" : "PEUGEOT-CITROEN",
    "TX" : "TEXTAR",
    "VR" : "VICTOR%20REINZ"

}


def excel_value_to_string(value):
    """
    Преобразует значение из Excel в строку, корректно обрабатывая числа
    """
    if value is None or value == '':
        return ''
    
    # Если это число (int или float)
    if isinstance(value, (int, float)):
        # Проверяем, является ли это целым числом
        if float(value).is_integer():
            return str(int(value))  # Возвращаем как целое число без .0
        else:
            return str(value)  # Возвращаем как есть для дробных чисел
    
    # Для всех остальных типов просто преобразуем в строку
    return str(value).strip()








try:
    wb = xw.Book(wb_path)
    active_sheet = wb.sheets.active
    sht = wb.sheets[active_sheet.name]
    data = active_sheet.range(f'{first_cell}:{last_cell}').value   #берет все данные с активного листа
    #addr = active_sheet.api.Application.ActiveCell.Address
    #print("Адрес:", addr)
except FileNotFoundError:
    print(f"НЕТ ФАЙЛА в папке {wb_path}")
    sys.exit()  #выход ибо нет файла




total_rows = user_row_number

print(f"Артикулы берем из файла '{Path(wb_path).name}'")
print(f"Из листа '{active_sheet.name}' в {first_cell}-{last_cell}" )
print(f"Всего {total_rows} позиций.")
print("")
print(f'{"Позиция".ljust(10):7}{"Артикул".ljust(20):15}{"Марка".ljust(20):15}Наименование')

#if total_rows == 1:


#print(data)

# Проходим по строкам файла и парсим данные
for row_index, row in enumerate(data):

    #if all(cell is None or cell == '' for cell in row):
     #   continue #пропускаем пустые строки

    #print(f"row_index:{row_index}, row:{row}")
    
    #raw_art = row.strip()
    raw_art = excel_value_to_string(row).upper() #в будущем надо будет использовать индексы если данные неодномерные
    #raw_art = str(row[0]).strip()  # все данные в row 
    #print(f"raw_art: {raw_art}" )
    new_art = raw_art
    marka = None
    #print(marka)
    #marka = str(row[1]).strip()  # марка из второго столбца

    original_proiz = ((sht.range(row_index + first_row, 12).value) or "").upper()   # or "" - чтобы не выдавал ошибку из-за None
    original_marka = ((sht.range(row_index + first_row, 13).value) or "").upper()
    #print(f"proiz: {proiz}, marka: {marka}")

    

    for brand in brand_replacement:
        if brand in raw_art:                        # если марки в артикуле
            new_art = raw_art.replace(brand, "")
            marka = brand_replacement[brand]
            break
    
    if original_marka != "":
        marka = original_marka
    elif original_proiz != "":
        marka = original_proiz
    
    if marka in brand_replacement:
        marka = brand_replacement[marka]

    if "BSG" in new_art:                            # это сделано потому что артикулы BSG содержат BSG на сайте dexup
        marka = "BSG"                               # BSG список исключений не добавлю потому что нельзя убирать BSG из артикула




        
 

    # Очистка артикула
    art = clean_artikul(new_art)
    
    # Проверка наличия марки в словаре и замена, если необходимо
    #if marka in brand_replacement:
    #    marka = brand_replacement[marka]
    
    #Формирование URL и парсинг данных с dexup
    url_dexup = f"https://dexup.ru/parts/{marka}/{art}"
    
    # Вызов функции парсинга страницы dexup
    product_name_dexup, mass_dexup, material_dexup = parse_page_dexup(url_dexup)

    
    
    # Приведение данных к корректному регистру
    product_name_dexup = capitalize_first_letter(product_name_dexup)
    material_dexup = capitalize_first_letter(material_dexup)

    # decode marka
    marka = unquote(marka or "")

    pos_index = (str(row_index + 1) + '/' + str(total_rows)).ljust(10)    #номер позиции
    print(f'{pos_index:7}{art.ljust(20):15}{marka.ljust(20):15}{product_name_dexup}')
    




    sht.range(row_index + first_row, 4).value = [
        product_name_dexup
    ]

    sht.range(row_index + first_row, 12).value = [              # марки в столбцы 12 и 13 (L и M)
        marka,
        marka
    ]

    sht.range(row_index + first_row, 16).value = [              # марки в столбцы 12 и 13 (L и M)
        mass_dexup
    ]


    #sht.range(row_index + first_row, 4).value = [
    #    product_name_dexup,
    #    mass_dexup,
    #    decoded_marka,
    #    material_dexup,
    #    url_dexup,
    #]


    time.sleep(1)
    
# Сохраняем обновленный Excel файл
#active_sheet.range("A1").value = data
print("")
print("Данные успешно сохранены в файл:", wb_path)
t1 = time.time()
print("Процесс занял", round((t1 - t0)/60), "минут")
print("")
