import pandas as pd
import requests
from bs4 import BeautifulSoup as BS
import re  # Модуль для работы с регулярными выражениями
from fake_useragent import UserAgent
import time
from urllib.parse import unquote


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
    "PIE": "PIEBURG",
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

# Загрузка Excel файла
file_path = "C:/GitHub/car_parts_script/part.xlsx"  # Путь к вашему файлу
print(f"Данные берем из {file_path}")
df = pd.read_excel(file_path, header=None)  # Читаем файл без заголовков

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
    print(f"Парсим данные по ссылке: {url_dexup.ljust(50)}", end='')
    
    # Вызов функции парсинга страницы dexup
    product_name_dexup, mass_dexup, material_dexup = parse_page_dexup(url_dexup)

    
    # Приведение данных к корректному регистру
    product_name_dexup = capitalize_first_letter(product_name_dexup)
    material_dexup = capitalize_first_letter(material_dexup)

    print(product_name_dexup)
    
    # Сохранение данных в соответствующие столбцы
    df.at[index, 2] = product_name_dexup  # Основное описание с dexup
    df.at[index, 3] = mass_dexup if mass_dexup is not None else float('nan')  # Масса с dexup
    df.at[index, 4] = material_dexup if material_dexup is not None else ""  # Материал с dexup
    df.at[index, 5] = url_dexup  # Ссылка на страницу dexup

    'DECODING URL'
    decoded_marka = unquote(marka)
    df.at[index, 6] = decoded_marka	


    time.sleep(1)

# Сохраняем обновленный Excel файл
output_file = "C:/GitHub/car_parts_script/part_PL.xlsx"  # Путь для сохранения файла
df.to_excel(output_file, index=False, header=False)

print("Данные успешно сохранены в файл:", output_file)
