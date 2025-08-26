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
first_cell = f'{user_column}{first_row}'
last_cell = f'{user_column}{last_row}'

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
    "PEUGEOT / CITROEN" : "PEUGEOT%20CITROEN",
    "PEUGEOTCITROEN" : "PEUGEOT%20CITROEN",
    "AUTOMEGA DELLO" : "AUTOMEGA",
    "MAHLE / KNECHT" : "MAHLE",
    "AUTOMEGADELLO" : "AUTOMEGA%20DELLO",
    "KOLBENSCHMIDT" : "KOLBENSCHMIDT",
    "KALE RADYATOR" : "KALE",
    "KALERADYATOR" : "KALE",
    "MERCEDESBENZ" : "MERCEDES%20BENZ",
    "MAHLE/KNECHT" : "MAHLE",
    "VICTOR REINZ" : "VICTOR%20REINZ",
    "VİCTOR REİNZ" : "VICTOR%20REINZ",
    "CONTINENTAL" : "CONTINENTAL",
    "HBJAKOPARTS" : "JAKOPARTS",
    "MAHLEKNECHT" : "MAHLE",
    "VICTORREINZ" : "VICTOR%20REINZ",
    "VİCTORREİNZ" : "VICTOR%20REINZ",
    "EREN BALATA" : "EREN%20BALATA",
    "ERENBALATA" : "EREN%20BALATA",
    "SCHMITZORG" : "SCHMITZ",
    "MANNFILTER" : "MANN%20FILTER",
    "BLUE PRINT" : "Blue%20Print",
    "HERTH+BUSS" : "H%2BB%20JAKOPARTS",
    "SSANG YONG" : "SSANG%20YONG",
    "VICTOR REI" : "VICTOR%20REINZ",
    "VİCTOR REİ" : "VICTOR%20REINZ",
    "BLUEPRINT" : "BLUEPRINT",
    "EUROREPAR" : "EUROREPAR",
    "GKNLOEBRO" : "GKN",
    "HERTHBUSS" : "H%2BB",
    "KACMAZLAR" : "KAÇMAZLAR",
    "LEMFORDER" : "LEMFORDER",
    "SSANGYONG" : "SSANGYONG",
    "VICTORREI" : "VICTOR%20REINZ",
    "VİCTORREİ" : "VICTOR%20REINZ",
    "NTN / SNR" : "NTN",
    "SSANGYONG" : "SSANG%20YONG",
    "AUTOMEGA" : "AUTOMEGA",
    "GOODYEAR" : "GOODYEAR",
    "BILSTEIN" : "BILSTEIN",
    "EURORAPE" : "EUROREPAR",
    "MBTRUCKS" : "MB%20TRUCKS",
    "EUROREPA" : "EUROREPAR",
    "LMFORDER" : "LEMFORDER",
    "PIERBURG" : "PIERBURG",
    "TEKNOROT" : "TEKNOROT",
    "VOLVOORG" : "VOLVO",
    "EURORAPE" : "EUROREPAR",
    "HYD HOME" : "HYD%20HOME",
    "LMFORDER" : "LEMFORDER",
    "BILSTEN" : "BILSTEIN",
    "CORTECO" : "CORTECO",
    "FEDERAL" : "FEDERAL%20MOGUL",
    "FILTRON" : "FILTRON",
    "MARELLI" : "MAGNETI%20MARELLI",
    "OPTIMAL" : "OPTIMAL",
    "HYDHOME" : "HYD%20HOME",
    "METELLI" : "METELLI",
    "NISSENS" : "NISSENS",
    "PIEBURG" : "PIERBURG",
    "PURFLUX" : "PURFLUX",
    "RENAULT" : "RENAULT",
    "AUGERCE" : "AUGER",
    "BILSTEN" : "BILSTEIN",
    "E.REPAR" : "EUROREPAR",
    "V.REINZ" : "VICTOR%20REINZ",
    "AIRTEX" : "AIRTEX",
    "BREMBO" : "BREMBO",
    "DELPHI" : "DELPHI",
    "ELRING" : "ELRING",
    "HOLSET" : "HOLSET",
    "EREPAR" : "EUROREPAR",
    "FERODO" : "FERODO",
    "KACMAZ" : "KACMAZ",
    "PACCAR" : "PACCAR",
    "PROVIA" : "PROVIA",
    "KAÇMAZ" : "KAÇMAZ",
    "KOLBEN" : "KOLBENSCHMIDT",
    "MONROE" : "MONROE",
    "NTNSNR" : "NTN%20SNR",
    "OTOSAN" : "OTOSAN",
    "REPAIR" : "EUROREPAR",
    "TEXTAR" : "TEXTAR",
    "TITANX" : "TITANX",
    "TIRSAN" : "TIRSAN",
    "TOPRAN" : "TOPRAN",
    "VERNET" : "VERNET",
    "VREINZ" : "VICTOR%20REINZ",
    "YENMAK" : "YENMAK",
    "KACMAZ" : "KACMAZLAR",
    "KAÇMAZ" : "KACMAZLAR",
    "KOLBEN" : "KOLBENSCHMIDT",
    "AISIN" : "AISIN",
    "BANDO" : "BANDO",
    "BESER" : "BESER",
    "BOSCH" : "BOSCH",
    "CIFAM" : "CIFAM",
    "CONTI" : "CONTINENTAL",
    "DAYCO" : "DAYCO",
    "DENSO" : "DENSO",
    "FACET" : "FACET",
    "GATES" : "GATES",
    "GLYCO" : "GLYCO",
    "HELLA" : "HELLA",
    "LUCAS" : "LUCAS",
    "IVECO" : "IVECO",
    "MAHLE" : "MAHLE",
    "MANDO" : "MANDO",
    "NÜRAL" : "NÜRAL",
    "MEYLE" : "MEYLE",
    "OSRAM" : "OSRAM",
    "RAPRO" : "RAPRO",
    "NURAL" : "NURAL",
    "REINZ" : "VICTOR%20REINZ",
    "SACHS" : "SACHS",
    "VOLVO" : "VOLVO",
    "VALEO" : "VALEO",
    "WABCO" : "WABCO",
    "BERU" : "BERU",
    "BEHR" : "BEHR",
    "BLUE" : "BLUE%20PRINT",
    "CAVO" : "CAVO",
    "DOLZ" : "DOLZ",
    "FEBI" : "FEBI",
    "FEBİ" : "FEBI",
    "FILT" : "FILTRON",
    "MİBA" : "MIBA",
    "FORD" : "FORD",
    "HUCO" : "HUCO",
    "KALE" : "KALE",
    "MANN" : "MANN%20FILTER",
    "ONKA" : "ONKA",
    "OPEL" : "OPEL",
    "SWAG" : "SWAG",
    "BLUE" : "Blue%20Print",
    "BSCH" : "BOSCH",
    "FILT" : "FILTRON",
    "BAN" : "BANDO",
    "BCH" : "BOSCH",
    "BER" : "BERU",
    "BIL" : "BILSTEIN",
    "BLU" : "BLUE%20PRINT",
    "BOS" : "BOSCH",
    "BRB" : "BREMBO",
    "BRE" : "BREMBO",
    "BRU" : "BLUEPRINT",
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
    "DYC" : "DAYCO",
    "ELR" : "ELRING",
    "FAC" : "FACET",
    "FAE" : "FAE",
    "FAG" : "FAG",
    "FBI" : "FEBI",
    "FEB" : "FEBI",
    "FLT" : "FILTRON",
    "FOR" : "FORD",
    "FRD" : "FORD",
    "FRJ" : "FIAT",
    "FTE" : "FTE",
    "GAT" : "GATES",
    "GKN" : "GKN",
    "GLY" : "GLYCO",
    "GTS" : "GATES",
    "HEL" : "HELLA",
    "HLL" : "HELLA",
    "INA" : "INA",
    "KAL" : "KALE",
    "KOL" : "KOLBENSCHMIDT",
    "LEM" : "LEMFORDER",
    "LMF" : "LEMFORDER",
    "LUK" : "LUK",
    "MAH" : "MAHLE",
    "MAI" : "MAHLE",
    "MAN" : "MANN%20FILTER",
    "MGA" : "AUTOMEGA",
    "MHL" : "MAHLE",
    "MND" : "MANDO",
    "MON" : "MONROE",
    "MTL" : "METELLI",
    "NGK" : "NGK",
    "NRF" : "NRF",
    "NTN" : "NTN",
    "OES" : "OES",
    "OPL" : "OPEL",
    "ORJ" : "ORIGINAL",
    "OSR" : "OSRAM",
    "OTO" : "OTO",
    "PIE" : "PIERBURG",
    "PRB" : "PIERBURG",
    "PRG" : "PIERBURG",
    "TRW" : "TRW",
    "HNG" : "HENGST",
    "PSA" : "PSA",
    "PUR" : "PURFLUX",
    "RAP" : "RAPRO",
    "RNZ" : "VICTOR%20REINZ",
    "NRV" : "NRV",
    "SAC" : "SACHS",
    "SCH" : "SACHS",
    "SCS" : "SACHS",
    "SCX" : "SACHS",
    "SHS" : "SACHS",
    "SKF" : "SKF",
    "SNR" : "SNR",
    "SWF" : "SWF",
    "SWG" : "SWAG",
    "TEK" : "TEKNOROT",
    "TRW" : "TRW",
    "UFI" : "UFI",
    "VAL" : "VALEO",
    "VCT" : "VICTOR%20REINZ",
    "VER" : "VERNET",
    "YEN" : "YENMAK",
    "BAN" : "BANDO",
    "BER" : "BERU",
    "BIL" : "BILSTEIN",
    "BLU" : "Blue%20Print",
    "BRU" : "BERU",
    "BOS" : "BOSCH",
    "BCH" : "BOSCH",
    "BSC" : "BOSCH",
    "BRE" : "BREMBO",
    "BRB" : "BREMBO",
    "CIF" : "CIFAM",
    "CON" : "Continental",
    "CNT" : "Continental",
    "COR" : "CORTECO",
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
    "GTS" : "GATES",
    "GAT" : "GATES",
    "GKN" : "GKN",
    "GLY" : "GLYCO",
    "HEL" : "HELLA",
    "HLL" : "HELLA",
    "HLL" : "HELLA",
    "INA" : "INA",
    "KOL" : "KOLBENSCHMIDT",
    "LEM" : "LEMFORDER",
    "LMF" : "LEMFORDER",
    "MAI" : "RENAULT",
    "MAH" : "MAHLE",
    "MHL" : "MAHLE",
    "MAN" : "MANN-FILTER",
    "MMA" : "MAGNETI%20MARELLI",
    "MND" : "MANDO",
    "MON" : "MONROE",
    "MTL" : "METELLI",
    "NGK" : "NGK",
    "OPL" : "OPEL",
    "OSR" : "OSRAM",
    "POJ" : "PEUGEOT%20CITROEN",
    "ORJ" : "PEUGEOT-CITROEN",
    "PIE" : "PIERBURG",
    "PUR" : "PURFLUX",
    "RAP" : "RAPRO",
    "RYL" : "ROYAL",
    "SAC" : "SACHS",
    "SCS" : "SACHS",
    "SKT" : "SKT",
    "SHS" : "SACHS",
    "SCH" : "SACHS",
    "SCX" : "SACHS",
    "HNG" : "HENGST",
    "NRV" : "NARVA",
    "SKF" : "SKF",
    "SNR" : "SNR",
    "SWG" : "SWAG",
    "TEK" : "TEKNOROT",
    "TRW" : "TRW",
    "RNZ" : "VICTOR%20REINZ",
    "VAL" : "VALEO",
    "VCT" : "VICTOR%20REINZ",
    "VER" : "VERNET",
    "WIN" : "WIN",
    "WOD" : "WOD",
    "YEN" : "YENMAK",
    "YNM" : "YENMAK",

}









try:
    wb = xw.Book(part_file)
    active_sheet = wb.sheets.active
    sht = wb.sheets[active_sheet.name]
    data = active_sheet.range(f'{first_cell}:{last_cell}').value   #берет все данные с активного листа
    #addr = active_sheet.api.Application.ActiveCell.Address
    #print("Адрес:", addr)
except FileNotFoundError:
    print(f"НЕТ ФАЙЛА part в папке {current_dir}\n")
    sys.exit()  #выход ибо нет файла

print


total_rows = user_row_number

print(f"Артикулы берем из файла '{Path(part_file).name}'")
print(f"Из листа '{active_sheet.name}' в {first_cell}-{last_cell}" )
print(f"Всего {total_rows} позиций.")
print("")
print(f'{"Позиция".ljust(10):7}{"Артикул".ljust(18):15}{"Марка".ljust(20):15}Наименование')



# Проходим по строкам файла и парсим данные
for row_index, row in enumerate(data):
    
    #if all(cell is None or cell == '' for cell in row):
     #   continue #пропускаем пустые строки

    #print(f"row_index:{row_index}, row:{row}")
    row = str(row)
    raw_art = row.strip()
    
    #raw_art = str(row[0]).strip()  # все данные в row 
    #print(f"raw_art: {raw_art}" )
    new_art = raw_art
    marka = None
    #print(marka)
    #marka = str(row[1]).strip()  # марка из второго столбца

    proiz = ((sht.range(row_index + first_row, 12).value) or "").upper()   # or "" - чтобы не выдавал ошибку из-за None
    marka = ((sht.range(row_index + first_row, 13).value) or "").upper()
    #print(f"proiz: {proiz}, marka: {marka}")

    for brand in brand_replacement:
        if brand in raw_art:                        # если марки в артикуле
            new_art = raw_art.replace(brand, "")
            marka = brand_replacement[brand]
        elif proiz in brand_replacement:            # если в столбце произв сокращенные марки
            marka = brand_replacement[proiz]
        elif marka in brand_replacement:            # если в столбце марки сокращенные марки
            marka = brand_replacement[marka]
        
 

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
    print(f'{pos_index:7}{art.ljust(18):15}{marka.ljust(20):15}{product_name_dexup}')
    
    sht.range(row_index + first_row, 4).value = [
        product_name_dexup
    ]


    sht.range(row_index + first_row, 12).value = [              # марки в столбцы 12 и 13 (L и M)
        marka,
        marka
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
print("Данные успешно сохранены в файл:", part_PL_file)
t1 = time.time()
print("Процесс занял", round((t1 - t0)/60), "минут(ы)")
print("")
