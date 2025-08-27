import pandas as pd
import requests
import sys
from bs4 import BeautifulSoup as BS
import re  # Модуль для работы с регулярными выражениями
from fake_useragent import UserAgent
import time
from urllib.parse import unquote
from pathlib import Path

current_dir = Path(__file__).parent 
part_file = current_dir / 'part.xlsx'
part_PL_file = current_dir / 'part_PL.xlsx'

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
    "METELLI" : "METELLI",
    "NISSENS" : "NISSENS",
    "OPTIMAL" : "OPTIMAL",
    "PIEBURG" : "PIERBURG",
    "PURFLUX" : "PURFLUX",
    "RENAULT" : "RENAULT",
    "V.REINZ" : "VICTOR%20REINZ",
    "KACMAZ" : "KACMAZLAR",
    "KAÇMAZ" : "KACMAZLAR",
    "KOLBEN" : "KOLBENSCHMIDT",
    "AIRTEX" : "AIRTEX",
    "BREMBO" : "BREMBO",
    "DELPHI" : "DELPHI",
    "ELRING" : "ELRING",
    "EREPAR" : "EUROREPAR",
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
    "IVECO" : "IVECO",
    "LUCAS" : "LUCAS",
    "MAHLE" : "MAHLE",
    "MANDO" : "MANDO",
    "MEYLE" : "MEYLE",
    "NURAL" : "NURAL",
    "NÜRAL" : "NÜRAL",
    "OSRAM" : "OSRAM",
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
    "MANN" : "MANN%20FILTER",
    "MİBA" : "MIBA",
    "ONKA" : "ONKA",
    "OPEL" : "OPEL",
    "SWAG" : "SWAG",
    "BAN" : "BANDO",
    "BCH" : "BOSCH",
    "BER" : "BERU",
    "BIL" : "BILSTEIN",
    "BLU" : "BLUE%20PRINT",
    "BOS" : "BOSCH",
    "BRB" : "BREMBO",
    "BRE" : "BREMI",
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
    "NRV" : "NARVA",
    "OPL" : "OPEL",
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
    "FTE" : "FTE",
    "GMB" : "GMB",
    "KAL" : "KALE",
    "KYB" : "KYB",
    "LUK" : "LUK",
    "MGA" : "AUTOMEGA",
    "MMA" : "MAGNETI%20MARELLI",
    "MND" : "MANDO",
    "NRF" : "NRF",
    "NTN" : "NTN",
    "OES" : "OES",
    "OTO" : "OTO",
    "POJ" : "PEUGEOT%20CITROEN",
    "PRB" : "PIERBURG",
    "PRG" : "PIERBURG",
    "PSA" : "PSA",
    "RYL" : "ROYAL",
    "SKT" : "SKT",
    "SWF" : "SWF",
    "TXT" : "TEXTAR",
    "UFI" : "UFI",
    "WIN" : "WIN",
    "WOD" : "WOD",
    "YTT" : "YTT",
    "YNM" : "YENMAK",
    "ORJ" : "ORIGINAL",
    "ORJ" : "PEUGEOT-CITROEN",

    

}

try:
    df = pd.read_excel(part_file, header=None, dtype=object)  # Читаем файл без заголовков. dtype=object решает проблему с dtype в df.at[index, 2] = product_name_dexup
except FileNotFoundError:
    print(f"НЕТ ФАЙЛА part в папке {current_dir}\n")
    sys.exit()  #выход ибо нет файла

total_rows = df.shape[0] - 1

print(f"Артикулы берем из {part_file}.")
print(f"Всего {total_rows} позиций.")
print("")
print(f'{"Позиция".ljust(10):7}{"Артикул".ljust(20):15}{"Марка".ljust(20):15}Наименование')


# Проходим по строкам файла и парсим данные
for index, row in df.iterrows():
    if index == 0:
        continue  # Пропускаем первую строку
    
    raw_art = str(row[0])  # Исходный артикул в первом столбце
    marka = str(row[1]).strip()  # Марка во втором столбце, очищаем от пробелов в начале и в конце
    
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

    row_index = (str(index) + '/' + str(total_rows)).ljust(10)    #номер позиции
    print(f'{row_index:7}{art.ljust(20):15}{decoded_marka.ljust(20):15}{product_name_dexup}')

    # Сохранение данных в соответствующие столбцы
    df.at[index, 2] = product_name_dexup # Основное описание с dexup
    df.at[index, 3] = mass_dexup if mass_dexup is not None else float('nan')  # Масса с dexup
    df.at[index, 4] = decoded_marka	
    df.at[index, 5] = material_dexup if material_dexup is not None else ""  # Материал с dexup
    df.at[index, 6] = url_dexup  # Ссылка на страницу dexup

    time.sleep(1)

# Сохраняем обновленный Excel файл
df.to_excel(part_PL_file, index=False, header=False)
print("")
print("Данные успешно сохранены в файл:", part_PL_file)
t1 = time.time()
print("Процесс занял", round((t1 - t0)/60), "минут(ы)")
print("")