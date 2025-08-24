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
part_file = sys.argv[1]    #второй аргумент (путь к файлу) при запуске скрипта через cmd 
part_PL_file = sys.argv[1]
first_cell = sys.argv[2]
first_row = int(sys.argv[3])
user_row_number = int(sys.argv[4])
last_row = first_row + user_row_number - 1
user_column = first_cell.split("$")[1]

#part_file = current_dir / 'part.xlsx'
#part_PL_file = current_dir / 'part_PL.xlsx'

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
    "AUGERCE" : "AUGER",
    "AUTOMEGA DELLO" : "AUTOMEGA",
    "BAN": "BANDO",
    "BER": "BERU",
    "BIL": "BILSTEIN",
    "BILSTEN": "BILSTEIN",
    "BLU": "Blue%20Print",
    "BLUE PRINT": "Blue%20Print",
    "BLUE": "Blue%20Print",
    "BRU": "BERU",	
    "BOS": "BOSCH",
    "BCH": "BOSCH",
    "BSC": "BOSCH",
    "BRE": "BREMBO",
    "BRB": "BREMBO",
    "CIF": "CIFAM",
    "CON": "Continental",
    "CNT": "Continental",
    "COR": "CORTECO",
    "COR": "CORTECO",
    "CRT": "CORTECO",
    "DAY": "DAYCO",
    "DEL": "DELPHI",
    "DEN": "DENSO",
    "DNS": "DENSO",
    "DOL": "DOLZ",
    "E.REPAR": "EUROREPAR",
    "EURORAPE": "EUROREPAR",
    "ELR": "ELRING",
    "EREN BALATA": "EREN%20BALATA",
    "FAC": "FACET",
    "FBI": "FEBI",
    "FEB": "FEBI",
    "FILT": "FILTRON",
    "FLT": "FILTRON",
    "FOR": "FORD",
    "FRD": "FORD",
    "GTS": "GATES",
    "GAT": "GATES",
    "GKN": "GKN%20%28Loebro%29",
    "GLY": "GLYCO",
    "HEL": "HELLA",
    "HERTH+BUSS" : "H%2BB%20JAKOPARTS",
    "HLL": "HELLA",
    "HLL": "HELLA",
    "HYD HOME": "HYD%20HOME",
    "INA": "INA",
    "KALE RADYATÖR": "KALE",
    "KACMAZ" : "KACMAZLAR",
    "KAÇMAZ" : "KACMAZLAR",
    "KOL": "KOLBENSCHMIDT",
    "KOLBEN": "KOLBENSCHMIDT",
    "LEM": "LEMFORDER",
    "LMF": "LEMFORDER",
    "LMFORDER": "LEMFORDER",
    "MAI": "RENAULT",
    "MAH": "MAHLE%2FKNECHT",
    "MHL": "MAHLE%2FKNECHT",
    "MAHLE / KNECHT" : "MAHLE%2FKNECHT",
    "MAHLE/KNECHT" : "MAHLE%2FKNECHT",
    "MAN": "MANN-FILTER",
    "MND": "MANDO",
    "MB" : "MERCEDES-BENZ",
    "MON": "MONROE",
    "MTL": "METELLI",
    "NGK": "NGK",
    "NTN / SNR": "NTN",
    "OPL": "OPEL",
    "OSR": "OSRAM",
    "ORJ" : "PEUGEOT-CITROEN",
    "PEUGEOT / CITROEN" : "PEUGEOT-CITROEN",
    "PIE": "PIERBURG",
    "PUR": "PURFLUX",
    "RAP": "RAPRO",
    "SAC": "SACHS",
    "SCS": "SACHS",
    "SHS": "SACHS",
    "SCH": "SACHS",
    "SCX": "SACHS",
    "SKF": "SKF",
    "SNR": "SNR",
    "SSANG YONG": "SSANG%20YONG",
    "SSANGYONG" : "SSANG%20YONG",
    "SWG": "SWAG",
    "TEK": "TEKNOROT",
    "TRW": "TRW",
    "VR": "VICTOR%20REINZ",
    "RNZ": "VICTOR%20REINZ",
    "V.REINZ": "VICTOR%20REINZ",
    "VAL" : "VALEO",
    "VCT": "VICTOR%20REINZ",
    "VER": "VERNET",
    "VICTOR REI" : "VICTOR%20REINZ",
    "VİCTOR REİ" : "VICTOR%20REINZ",
    "VICTOR REINZ": "VICTOR%20REINZ",
    "VİCTOR REİNZ": "VICTOR%20REINZ",
    "YEN": "YENMAK"
}









try:
    wb = xw.Book(part_file)
    active_sheet = wb.sheets.active
    sht = wb.sheets[active_sheet.name]
    data = active_sheet.range(first_cell).current_region.value   #берет все данные с активного листа
    #addr = active_sheet.api.Application.ActiveCell.Address
    #print("Адрес:", addr)
except FileNotFoundError:
    print(f"НЕТ ФАЙЛА part в папке {current_dir}\n")
    sys.exit()  #выход ибо нет файла

print


total_rows = user_row_number

print(f"Артикулы берем из файла '{Path(part_file).name}'")
print(f"Из листа '{active_sheet.name}' в {user_column}{first_row}-{user_column}{last_row}" )
print(f"Всего {total_rows} позиций.")
print("")
print(f'{"Позиция".ljust(10):7}{"Артикул".ljust(20):15}{"Марка".ljust(20):15}Наименование')



# Проходим по строкам файла и парсим данные
for row_index, row in enumerate(data):
    if all(cell is None or cell == '' for cell in row):
        continue #пропускаем пустые строки
    #print(f"row_index:{row_index}, row:{row}")
    
    raw_art = str(row[0]).strip()  # все данные в row 
    marka = str(row[1]).strip()  
    
    #print(f"raw_art: {raw_art}, marka: {marka}")
 

    # Очистка артикула
    art = clean_artikul(raw_art)
    
    # Проверка наличия марки в словаре и замена, если необходимо
    if marka in brand_replacement:
        marka = brand_replacement[marka]
    
    #Формирование URL и парсинг данных с dexup
    url_dexup = f"https://dexup.ru/parts/{marka}/{art}"
    
    # Вызов функции парсинга страницы dexup
    product_name_dexup, mass_dexup, material_dexup = parse_page_dexup(url_dexup)

    
    # Приведение данных к корректному регистру
    product_name_dexup = capitalize_first_letter(product_name_dexup)
    material_dexup = capitalize_first_letter(material_dexup)

    'DECODING URL'
    decoded_marka = unquote(marka)

    pos_index = (str(row_index + 1) + '/' + str(total_rows)).ljust(10)    #номер позиции
    print(f'{pos_index:7}{art.ljust(20):15}{decoded_marka.ljust(20):15}{product_name_dexup}')

    sht.range(row_index + first_row, 3).value = [
        product_name_dexup,
        mass_dexup,
        decoded_marka,
        material_dexup,
        url_dexup,
    ]


    time.sleep(1)
    
# Сохраняем обновленный Excel файл
#active_sheet.range("A1").value = data
print("")
print("Данные успешно сохранены в файл:", part_PL_file)
t1 = time.time()
print("Процесс занял", round((t1 - t0)/60), "минут(ы)")
print("")
