import sys
import re
import sqlite3
from pathlib import Path
from urllib.parse import unquote

import xlwings as xw


# -----------------------------------------------------------------------------
# Настройки
# -----------------------------------------------------------------------------
DATABASE_PATH = Path(r"C:\meylis\car-parts-script\car-parts-db.db")

ARTICLE_COLUMN = 2       # B
PRODUCT_NAME_COLUMN = 4  # D
BRAND_L_COLUMN = 12      # L
BRAND_M_COLUMN = 13      # M
WEIGHT_COLUMN = 16       # P


# -----------------------------------------------------------------------------
# Марки и сокращения из рабочего Dexup-парсера
# -----------------------------------------------------------------------------
# ВАЖНО:
# - это словарь именно для РАСПОЗНАВАНИЯ марки;
# - марка внутри артикула считается маркой только как ПРЕФИКС;
# - найденный префикс никогда сам по себе не является достаточным основанием:
#   ниже обязательно проверяется точная пара (артикул + марка) в SQLite.
#
# Специальные случаи приведены к тому виду, в котором марка реально
# сохраняется вашим Dexup-парсером в Excel/SQLite:
#   MAIS -> MAIS
#   FIA  -> FIAT
#   BSG  -> BSG
#
# В исходном словаре "MGA" указан дважды. В Python dict последнее значение
# заменяет предыдущее, поэтому фактическое значение рабочего парсера: MGA -> MGA.
BRAND_REPLACEMENT = {'PEUGEOT / CITROEN': 'PEUGEOT%20CITROEN',
 'MAGNETI MARELLI': 'MAGNETI MARELLI',
 'AUTOMEGA DELLO': 'AUTOMEGA',
 'MAHLE / KNECHT': 'MAHLE',
 'PEUGEOTCITROEN': 'PEUGEOT%20CITROEN',
 'AUTOMEGADELLO': 'AUTOMEGA%20DELLO',
 'KALE RADYATOR': 'KALE',
 'KOLBENSCHMIDT': 'KOLBENSCHMIDT',
 'KALERADYATOR': 'KALE',
 'MAHLE/KNECHT': 'MAHLE',
 'MERCEDESBENZ': 'MERCEDES%20BENZ',
 'VICTOR REINZ': 'VICTOR%20REINZ',
 'VİCTOR REİNZ': 'VICTOR%20REINZ',
 'MANN-FILTER': 'MANN-FILTER',
 'CONTINENTAL': 'CONTINENTAL',
 'EREN BALATA': 'EREN%20BALATA',
 'HBJAKOPARTS': 'JAKOPARTS',
 'MAHLEKNECHT': 'MAHLE',
 'VICTORREINZ': 'VICTOR%20REINZ',
 'VİCTORREİNZ': 'VICTOR%20REINZ',
 'SACHS YEREL': 'SACHS',
 'BLUE PRINT': 'Blue%20Print',
 'ERENBALATA': 'EREN%20BALATA',
 'HERTH+BUSS': 'H%2BB%20JAKOPARTS',
 'MANNFILTER': 'MANN-FILTER',
 'SCHMITZORG': 'SCHMITZ',
 'SSANG YONG': 'SSANG%20YONG',
 'VICTOR REI': 'VICTOR%20REINZ',
 'VİCTOR REİ': 'VICTOR%20REINZ',
 'HUTCHINSON': 'HUTCHINSON',
 'SSANGYONG': 'SSANG%20YONG',
 'BLUEPRINT': 'BLUEPRINT',
 'EUROREPAR': 'EUROREPAR',
 'GKNLOEBRO': 'GKN',
 'BORSEHUNG': 'BORSEHUNG',
 'HERTHBUSS': 'H%2BB',
 'KACMAZLAR': 'KAÇMAZLAR',
 'LEMFORDER': 'LEMFORDER',
 'NTN / SNR': 'NTN',
 'VICTORREI': 'VICTOR%20REINZ',
 'VİCTORREİ': 'VICTOR%20REINZ',
 'KRAFTVOLL': 'KRAFTVOLL',
 'CONTITECH': 'CONTITECH',
 'EURORAPE': 'EUROREPAR',
 'LMFORDER': 'LEMFORDER',
 'KOLBENSC': 'KOLBENSCHMIDT',
 'AUTOMEGA': 'AUTOMEGA',
 'BILSTEIN': 'BILSTEIN',
 'EUROREPA': 'EUROREPAR',
 'GOODYEAR': 'GOODYEAR',
 'HYD HOME': 'HYD%20HOME',
 'MBTRUCKS': 'MB%20TRUCKS',
 'PIERBURG': 'PIERBURG',
 'TEKNOROT': 'TEKNOROT',
 'VOLVOORG': 'VOLVO',
 'EUROBUMP': 'EUROBUMP',
 'CONTITEC': 'CONTITECH',
 'BILSTEN': 'BILSTEIN',
 'AUGERCE': 'AUGER',
 'CORTECO': 'CORTECO',
 'E.REPAR': 'EUROREPAR',
 'FEDERAL': 'FEDERAL%20MOGUL',
 'FILTRON': 'FILTRON',
 'HYDHOME': 'HYD%20HOME',
 'MARELLI': 'MAGNETI%20MARELLI',
 'MAGNETI': 'MAGNETI%20MARELLI',
 'METELLI': 'METELLI',
 'PSA-PEUG': 'PEUGEOT-CITROEN',
 'NISSENS': 'NISSENS',
 'OPTIMAL': 'OPTIMAL',
 'PIEBURG': 'PIERBURG',
 'PURFLUX': 'PURFLUX',
 'PLEKSAN': 'PLEKSAN',
 'SNR-NTN': 'SNR',
 'RENAULT': 'RENAULT',
 'GARRETT': 'GARRETT',
 'V.REINZ': 'VICTOR%20REINZ',
 'KOLBEN': 'KOLBENSCHMIDT',
 'KACMAZ': 'KACMAZLAR',
 'KAÇMAZ': 'KACMAZLAR',
 'KONEKS': 'KONEKS',
 'HENGST': 'HENGST',
 'WAHLER': 'WAHLER',
 'AIRTEX': 'AIRTEX',
 'BREMBO': 'BREMBO',
 'DELPHI': 'DELPHI',
 'ELRING': 'ELRING',
 'EREPAR': 'EUROREPAR',
 'GOETZE': 'GOETZE',
 'EYQUEM': 'EYQUEM',
 'FERODO': 'FERODO',
 'HOLSET': 'HOLSET',
 'MONROE': 'MONROE',
 'NTNSNR': 'NTN%20SNR',
 'OTOSAN': 'OTOSAN',
 'PACCAR': 'PACCAR',
 'PROVIA': 'PROVIA',
 'REPAIR': 'EUROREPAR',
 'TEXTAR': 'TEXTAR',
 'TIRSAN': 'TIRSAN',
 'TITANX': 'TITANX',
 'TOPRAN': 'TOPRAN',
 'VERNET': 'VERNET',
 'VREINZ': 'VICTOR%20REINZ',
 'YENMAK': 'YENMAK',
 'YILMAZ': 'YILMAZ',
 'TURTEL': 'TURTEL',
 'AISIN': 'AISIN',
 'AJUSA': 'AJUSA',
 'BANDO': 'BANDO',
 'BESER': 'BESER',
 'BOSCH': 'BOSCH',
 'CIFAM': 'CIFAM',
 'CONTI': 'CONTITECH',
 'DAYCO': 'DAYCO',
 'DENSO': 'DENSO',
 'DEKAR': 'DEKAR',
 'FACET': 'FACET',
 'GATES': 'GATES',
 'GLYCO': 'GLYCO',
 'IBRAS': 'IBRAS',
 'HELLA': 'HELLA',
 'IVECO': 'IVECO',
 'LUCAS': 'LUCAS',
 'ONPER': 'ONPER',
 'MAHLE': 'MAHLE',
 'MANDO': 'MANDO',
 'MEYLE': 'MEYLE',
 'NURAL': 'NURAL',
 'NÜRAL': 'NÜRAL',
 'OSRAM': 'OSRAM',
 'SAHIN': 'SAHIN',
 'RAPRO': 'RAPRO',
 'REINZ': 'VICTOR%20REINZ',
 'SACHS': 'SACHS',
 'VADEN': 'VADEN',
 'VALEO': 'VALEO',
 'VOLVO': 'VOLVO',
 'WABCO': 'WABCO',
 'BLUE': 'BLUE%20PRINT',
 'FILT': 'FILTRON',
 'BEHR': 'BEHR',
 'BERU': 'BERU',
 'BSCH': 'BOSCH',
 'CAVO': 'CAVO',
 'DOLZ': 'DOLZ',
 'DEPO': 'DEPO',
 'FEBI': 'FEBI',
 'FEBİ': 'FEBI',
 'FORD': 'FORD',
 'HUCO': 'HUCO',
 'KALE': 'KALE',
 'UCEL': 'UC-EL',
 'VALS': 'VALEO',
 'VALE': 'VALEO',
 'VALA': 'VALEO',
 'VIKA': 'VIKA',
 'KAYA': 'KAYA',
 'KING': 'KING',
 'MANN': 'MANN-FILTER',
 'MARS': 'MARS',
 'MAIS': 'MAIS',
 'MEHA': 'MEHA',
 'MİBA': 'MIBA',
 'TRSN': 'TIRSAN',
 'GRAF': 'GRAF',
 'ONKA': 'ONKA',
 'OPEL': 'OPEL',
 'SWAG': 'SWAG',
 'AIS': 'AISIN',
 'ASP': 'ASPOCK',
 'FOR': 'FORD',
 'AYF': 'AYFAR',
 'GMB': 'GMB',
 'BAN': 'BANDO',
 'BCH': 'BOSCH',
 'BER': 'BERU',
 'BIL': 'BILSTEIN',
 'BLU': 'BLUE%20PRINT',
 'BLP': 'BLUE%20PRINT',
 'BOS': 'BOSCH',
 'BSH': 'BOSCH',
 'BMW': 'BMW',
 'BRB': 'BREMBO',
 'BRE': 'BREMBO',
 'BRU': 'BERU',
 'BRS': 'BORSEHUNG',
 'BSC': 'BOSCH',
 'CIF': 'CIFAM',
 'CNT': 'CONTITECH',
 'CNR': 'CNR',
 'CTR': 'CTR',
 'CON': 'CONTITECH',
 'COR': 'CORTECO',
 'CRT': 'CORTECO',
 'DAY': 'DAYCO',
 'DEL': 'DELPHI',
 'DEN': 'DENSO',
 'DEG': 'DE-GA',
 'DEGA': 'DE-GA',
 'DNS': 'DENSO',
 'DPO': 'DEPO',
 'DOL': 'DOLZ',
 'ECE': 'ECEM',
 'ELR': 'ELRING',
 'ERA': 'ERA',
 'EUR': 'EUROBUMP',
 'FAC': 'FACET',
 'FBI': 'FEBI',
 'FEB': 'FEBI',
 'FLT': 'FILTRON',
 'FIA': 'FIAT',
 'FRD': 'FORD',
 'GAT': 'GATES',
 'GKN': 'GKN',
 'GLY': 'GLYCO',
 'GTS': 'GATES',
 'GSP': 'GSP',
 'GVA': 'GVA',
 'KSC': 'KOLBENSCHMIDT',
 'HEL': 'HELLA',
 'REMSA': 'REMSA',
 'HOL': 'HOLSET',
 'HLL': 'HELLA',
 'HNG': 'HENGST',
 'INA': 'INA',
 'KOL': 'KOLBENSCHMIDT',
 'KNG': 'KONGSBERG',
 'LEM': 'LEMFORDER',
 'LMF': 'LEMFORDER',
 'MAH': 'MAHLE',
 'MAI': 'RENAULT',
 'MAN': 'MANN-FILTER',
 'MAY': 'MAYSAN%20MANDO',
 'MER': 'MERCEDES-BENZ',
 'MHL': 'MAHLE',
 'MON': 'MONROE',
 'MTA': 'MTA',
 'MTL': 'METELLI',
 'NGK': 'NGK',
 'NIS': 'NISSENS',
 'NıS': 'NISSENS',
 'NRV': 'NARVA',
 'KRA': 'KRAFTVOLL',
 'OPL': 'OPEL',
 'OPT': 'OPTIMAL',
 'OSM': 'OSRAM',
 'OSR': 'OSRAM',
 'ULO': 'ULO',
 'PIE': 'PIERBURG',
 'PUR': 'PURFLUX',
 'RAP': 'RAPRO',
 'RNZ': 'VICTOR%20REINZ',
 'SAC': 'SACHS',
 'SCH': 'SACHS',
 'SCS': 'SACHS',
 'SCX': 'SACHS',
 'SHS': 'SACHS',
 'SKF': 'SKF',
 'SNR': 'SNR',
 'SWG': 'SWAG',
 'TEK': 'TEKNOROT',
 'TPR': 'TOPRAN',
 'TRW': 'TRW',
 'VAL': 'VALEO',
 'VCT': 'VICTOR%20REINZ',
 'VER': 'VERNET',
 'VIK': 'VIKA',
 'YEN': 'YENMAK',
 'DAF': 'DAF',
 'DYC': 'DAYCO',
 'FAE': 'FAE',
 'FAG': 'FAG',
 'FRJ': 'FIAT',
 'KRF': 'KRAFTVOLL',
 'LPR': 'LPR',
 'FMN': 'FMN',
 'FSE': 'FASE',
 'FTE': 'FTE',
 'KLR': 'KALE',
 'KAL': 'KALE',
 'KYB': 'KYB',
 'LUK': 'LUK',
 'MGA': 'MGA',
 'MMA': 'MAGNETI%20MARELLI',
 'MND': 'MANDO',
 'MNN': 'MANN-FILTER',
 'MAP': 'MAPA',
 'MKS': 'MKS',
 'NRF': 'NRF',
 'NTN': 'NTN',
 'PAY': 'PAYEN',
 'OES': 'OES',
 'GNS': 'GUNES',
 'IBR': 'IBRAS',
 'IVE': 'IVECO',
 'OTO': 'OTO',
 'POJ': 'PEUGEOT%20CITROEN',
 'PEU': 'PEUGEOT%20CITROEN',
 'PRB': 'PIERBURG',
 'PRG': 'PIERBURG',
 'PSA': 'PEUGEOT-CITROEN',
 'CAV': 'CAVO',
 'REN': 'RENAULT',
 'RYL': 'ROYAL',
 'MAR': 'MARS',
 'SKT': 'SKT',
 'SMP': 'SAMPART',
 'SWF': 'SWF',
 'TXT': 'TEXTAR',
 'TUR': 'TURTEL',
 'TIR': 'TIRSAN',
 'UFI': 'UFI',
 'WIN': 'WIN',
 'WOD': 'WOD',
 'WHL': 'WAHLER',
 'YTT': 'YTT',
 'YNM': 'YENMAK',
 'VAG': 'VAG',
 'ORJ': 'VAG',
 'K&B': 'K&B',
 'TX': 'TEXTAR',
 'BR': 'BERU',
 'GM': 'GENERAL MOTORS',
 'VR': 'VICTOR%20REINZ',
 'LF': 'LEMFORDER',
 'MB': 'MERCEDES%20BENZ',
 'ZF_': 'LEMFORDER',
 'ZFT': 'ZF',
 'BSG': 'BSG'}


# -----------------------------------------------------------------------------
# Вспомогательные функции
# -----------------------------------------------------------------------------
def excel_value_to_string(value):
    """Корректно переводит значение Excel в строку, не добавляя .0 целым числам."""
    if value is None or value == "":
        return ""

    if isinstance(value, bool):
        return str(value)

    if isinstance(value, (int, float)):
        try:
            if float(value).is_integer():
                return str(int(value))
        except (TypeError, ValueError, OverflowError):
            pass
        return str(value)

    return str(value).strip()



def product_name_cell_needs_fill(value):
    """
    True = ячейку D можно заполнить/исправить.

    Разрешено:
    - пустая ячейка;
    - значение, начинающееся с Error;
    - значение, содержащее not found.

    Любое другое непустое значение считается нормальным наименованием
    и не должно изменяться.
    """
    text = excel_value_to_string(value).strip()
    text_lower = text.lower()

    if not text:
        return True

    if text_lower.startswith("error"):
        return True

    if "not found" in text_lower:
        return True

    return False


def normalize_article(value):
    """
    Нормализация артикула для поиска.
    По смыслу соответствует clean_artikul() из парсера.
    """
    text = excel_value_to_string(value).upper().strip()
    if not text:
        return ""

    return re.sub(r"\W+", "", text, flags=re.UNICODE)


def normalize_brand(value):
    """
    Нормализует марку для сравнения с БД.

    URL-кодирование из словаря (%20, %2B и т.п.) сначала декодируется.
    Пунктуация/пробелы не влияют на сравнение.
    """
    text = unquote(excel_value_to_string(value)).upper().strip()
    if not text:
        return ""

    # Unicode-буквы сохраняются; пробелы, дефисы, +, /, _ и т.п. убираются.
    return re.sub(r"[\W_]+", "", text, flags=re.UNICODE)


def _add_unique(target, value):
    if value and value not in target:
        target.append(value)


def build_brand_alias_index():
    """
    Строит:
      1) индекс сокращений для значений из M/L;
      2) список префиксов для распознавания марки внутри артикула.

    Длинные префиксы проверяются раньше коротких.
    """
    hint_index = {}
    prefix_entries = []

    for alias, db_brand in BRAND_REPLACEMENT.items():
        alias_brand_key = normalize_brand(alias)
        target_brand_key = normalize_brand(db_brand)
        alias_article_prefix = normalize_article(alias)

        if not alias_brand_key or not target_brand_key or not alias_article_prefix:
            continue

        hint_index.setdefault(alias_brand_key, [])
        _add_unique(hint_index[alias_brand_key], target_brand_key)

        prefix_entries.append(
            (
                alias_article_prefix,
                target_brand_key,
                alias,
            )
        )

    # Самый длинный вариант вперед:
    # VICTORREINZ должен проверяться раньше VCT/VR,
    # MONROE раньше MON и т.д.
    prefix_entries.sort(
        key=lambda item: len(item[0]),
        reverse=True,
    )

    return hint_index, prefix_entries


BRAND_HINT_INDEX, BRAND_PREFIX_ENTRIES = build_brand_alias_index()


def brand_keys_from_excel_hint(value):
    """
    Возвращает допустимые ключи марки из M или L.

    Например:
      MON -> MONROE
      BOS -> BOSCH
      LEM -> LEMFORDER

    Полное название марки тоже остается допустимым.
    """
    text = excel_value_to_string(value).strip()
    if not text:
        return []

    result = []

    # Сначала пробуем значение как нормальную полную марку.
    direct_key = normalize_brand(text)
    _add_unique(result, direct_key)

    # Потом добавляем перевод сокращения из словаря.
    for mapped_key in BRAND_HINT_INDEX.get(direct_key, []):
        _add_unique(result, mapped_key)

    return result


def build_plain_article_candidates(raw_article):
    """
    Консервативные варианты артикула БЕЗ удаления произвольных букв.

    В отличие от старого варианта здесь НЕ вырезается случайный текст
    из середины артикула и НЕ берутся отдельные числовые куски.
    """
    result = []
    normalized = normalize_article(raw_article)

    if normalized:
        result.append(normalized)

    return result


def find_embedded_brand_records(
    raw_article,
    exact_map,
    allowed_brand_keys=None,
):
    """
    Ищет марку, вшитую В НАЧАЛО артикула.

    Пример:
      MONG7455       -> MON -> MONROE, артикул G7455
      BOS0986475987  -> BOS -> BOSCH,  артикул 0986475987
      LEM12345       -> LEM -> LEMFORDER, артикул 12345

    Защита от ложных совпадений:
    - ищем только префикс, НЕ произвольное вхождение внутри строки;
    - длинные алиасы проверяются первыми;
    - после префикса должен остаться артикул;
    - в остатке должна быть хотя бы одна цифра;
    - обязательно должна существовать точная пара
      (очищенный артикул, марка) в SQLite;
    - если найдено несколько РАЗНЫХ товаров, функция сообщает неоднозначность.
    """
    normalized_full = normalize_article(raw_article)
    if not normalized_full:
        return [], []

    allowed = None
    if allowed_brand_keys:
        allowed = set(allowed_brand_keys)

    matches = {}
    matched_aliases = []

    for alias_prefix, brand_key, alias_text in BRAND_PREFIX_ENTRIES:
        if allowed is not None and brand_key not in allowed:
            continue

        if not normalized_full.startswith(alias_prefix):
            continue

        remainder = normalized_full[len(alias_prefix):]

        if not remainder:
            continue

        # Для коротких кодов это критическая защита от случайного совпадения.
        # Такой же принцип используется системами, где "склеенный" brand-code
        # распознается только когда оставшаяся часть похожа на MPN и содержит цифры.
        if not any(ch.isdigit() for ch in remainder):
            continue

        record = exact_map.get((remainder, brand_key))
        if record is None:
            continue

        record_key = (
            normalize_article(record["article"]),
            normalize_brand(record["brand"]),
        )

        matches[record_key] = record
        matched_aliases.append(
            (alias_text, remainder, record["brand"])
        )

    return list(matches.values()), matched_aliases


def find_part(
    raw_article,
    brand_m,
    brand_l,
    exact_map,
    article_map,
    ambiguous_articles,
):
    """
    Безопасный порядок поиска:

    1. Если M заполнен — M является основным brand-hint.
       Если M пуст — используется L.
    2. Сначала точная пара article + brand.
    3. Если марка вшита в НАЧАЛО артикула — отделяем только известный
       префикс/сокращение и снова требуем точную пару article + brand в БД.
       Если M/L заполнены, embedded-brand разрешен только для той же марки.
    4. Только если марки нет ни в M, ни в L и в артикуле она не распознана,
       разрешается старый fallback по одному артикулу — но лишь когда
       этот артикул в текущей БД однозначно принадлежит одной марке.
    """
    brand_hint = excel_value_to_string(brand_m).strip()
    if not brand_hint:
        brand_hint = excel_value_to_string(brand_l).strip()

    plain_articles = build_plain_article_candidates(raw_article)
    explicit_brand_keys = brand_keys_from_excel_hint(brand_hint)

    # ------------------------------------------------------------------
    # 1) Есть M/L: НИКОГДА не игнорируем указанную там марку.
    # ------------------------------------------------------------------
    if explicit_brand_keys:
        for brand_key in explicit_brand_keys:
            for article_key in plain_articles:
                record = exact_map.get((article_key, brand_key))
                if record is not None:
                    return record, "M/L", None

        # Если в B стоит, например, MONG7455, а M = MONROE:
        # разрешаем MON только потому, что он ведет к той же марке MONROE.
        embedded_records, aliases = find_embedded_brand_records(
            raw_article,
            exact_map,
            allowed_brand_keys=explicit_brand_keys,
        )

        if len(embedded_records) == 1:
            return embedded_records[0], "ARTICLE_PREFIX", aliases

        # Если M/L заполнены, не делаем article-only fallback:
        # иначе можно молча проигнорировать заданную пользователем марку.
        return None, "BRAND_HINT_NO_MATCH", aliases

    # ------------------------------------------------------------------
    # 2) M/L пусты: пробуем известный префикс/сокращение в артикуле.
    # ------------------------------------------------------------------
    embedded_records, aliases = find_embedded_brand_records(
        raw_article,
        exact_map,
        allowed_brand_keys=None,
    )

    if len(embedded_records) == 1:
        return embedded_records[0], "ARTICLE_PREFIX", aliases

    if len(embedded_records) > 1:
        # Не угадываем между двумя товарами.
        return None, "AMBIGUOUS_ARTICLE_PREFIX", aliases

    # ------------------------------------------------------------------
    # 3) Последний fallback: только сам артикул, только если он однозначен
    #    по марке внутри текущей базы.
    # ------------------------------------------------------------------
    for article_key in plain_articles:
        if article_key in ambiguous_articles:
            continue

        record = article_map.get(article_key)
        if record is not None:
            return record, "UNIQUE_ARTICLE", None

    return None, "NOT_FOUND", None

def load_latest_parts(connection):
    """
    Один раз загружает последние записи из базы.

    Возвращает:
      exact_map[(article, brand)] -> record
      article_map[article]        -> record, только если артикул однозначен по марке
      ambiguous_articles          -> set артикулов, встречающихся у разных марок
    """
    exact_map = {}
    article_map = {}
    ambiguous_articles = set()

    view_exists = connection.execute(
        """
        SELECT 1
        FROM sqlite_master
        WHERE type='view' AND name='latest_parts'
        LIMIT 1
        """
    ).fetchone()

    if view_exists:
        query = """
            SELECT article, brand, weight_kg, product_name, source_file, inserted_at
            FROM latest_parts
        """
    else:
        query = """
            SELECT
                p.article,
                p.brand,
                p.weight_kg,
                p.product_name,
                p.source_file,
                p.inserted_at
            FROM parts AS p
            INNER JOIN (
                SELECT article, brand, MAX(id) AS max_id
                FROM parts
                GROUP BY article, brand
            ) AS latest
                ON latest.max_id = p.id
        """

    rows = connection.execute(query).fetchall()

    for row in rows:
        article = normalize_article(row[0])
        brand = normalize_brand(row[1])

        if not article:
            continue

        record = {
            "article": excel_value_to_string(row[0]),
            "brand": excel_value_to_string(row[1]),
            "weight_kg": row[2],
            "product_name": excel_value_to_string(row[3]),
            "source_file": excel_value_to_string(row[4]),
            "inserted_at": excel_value_to_string(row[5]),
        }

        exact_map[(article, brand)] = record

        if article not in article_map:
            article_map[article] = record
        else:
            previous_brand = normalize_brand(article_map[article]["brand"])
            if previous_brand != brand:
                ambiguous_articles.add(article)

    return exact_map, article_map, ambiguous_articles, len(rows)



def main():
    if len(sys.argv) != 5:
        print(
            "ОШИБКА: ожидались аргументы:\n"
            "1) путь к Excel\n"
            "2) имя листа\n"
            "3) первая строка\n"
            "4) последняя строка"
        )
        sys.exit(1)

    wb_path = Path(sys.argv[1])
    sheet_name = sys.argv[2]
    first_row = int(sys.argv[3])
    last_row = int(sys.argv[4])

    if first_row < 1:
        raise ValueError("Первая строка должна быть >= 1.")

    if last_row < first_row:
        print("Нет строк для обработки.")
        return

    if not DATABASE_PATH.is_file():
        print("")
        print("ОШИБКА: база данных не найдена.")
        print(f"Ожидался файл: {DATABASE_PATH}")
        sys.exit(1)

    if not wb_path.is_file():
        print("")
        print("ОШИБКА: Excel-файл не найден.")
        print(f"Файл: {wb_path}")
        sys.exit(1)

    print(f"Excel: {wb_path.name}")
    print(f"Лист: {sheet_name}")
    print(f"Строки: {first_row}-{last_row}")
    print(f"База: {DATABASE_PATH}")
    print("")

    connection = None

    try:
        connection = sqlite3.connect(str(DATABASE_PATH), timeout=30)

        exact_map, article_map, ambiguous_articles, loaded_count = load_latest_parts(
            connection
        )

        print(f"Последних записей загружено из базы: {loaded_count}")

        # xw.Book(full_path) сначала ищет уже открытую книгу среди
        # запущенных экземпляров Excel. Если книга уже открыта, xlwings
        # подключается именно к ней и работает с ее текущими (в том числе
        # еще не сохраненными) значениями.
        try:
            wb = xw.Book(str(wb_path))
        except Exception as e:
            raise RuntimeError(
                f"Не удалось подключиться к открытой книге Excel "
                f"'{wb_path.name}': {e}"
            )

        try:
            ws = wb.sheets[sheet_name]
        except Exception:
            raise RuntimeError(
                f"Лист '{sheet_name}' не найден в книге '{wb_path.name}'."
            )

        found_count = 0
        not_found_count = 0
        empty_count = 0
        skipped_existing_name_count = 0
        found_by_article_brand_count = 0

        total = last_row - first_row + 1

        print("")
        print(
            f'{"Позиция".ljust(12)}'
            f'{"Артикул".ljust(24)}'
            f'{"Марка".ljust(22)}'
            f'{"Вес, кг".ljust(12)}'
            f'Наименование'
        )

        for index, row_number in enumerate(range(first_row, last_row + 1), start=1):
            current_product_name = ws.cells(
                row_number,
                PRODUCT_NAME_COLUMN
            ).value

            if not product_name_cell_needs_fill(current_product_name):
                skipped_existing_name_count += 1

                raw_article_for_print = excel_value_to_string(
                    ws.cells(row_number, ARTICLE_COLUMN).value
                )

                print(
                    f'{(str(index) + "/" + str(total)).ljust(12)}'
                    f'{raw_article_for_print.ljust(24)}'
                    f'{"".ljust(22)}'
                    f'{"".ljust(12)}'
                    f'ПРОПУЩЕНО: в D уже есть нормальное наименование'
                )
                continue

            raw_article = ws.cells(row_number, ARTICLE_COLUMN).value
            article_for_print = excel_value_to_string(raw_article)

            if not article_for_print:
                empty_count += 1
                print(
                    f'{(str(index) + "/" + str(total)).ljust(12)}'
                    f'{"".ljust(24)}'
                    f'{"".ljust(22)}'
                    f'{"".ljust(12)}'
                    f'Пустой артикул'
                )
                continue

            brand_m = excel_value_to_string(
                ws.cells(row_number, BRAND_M_COLUMN).value
            )
            brand_l = excel_value_to_string(
                ws.cells(row_number, BRAND_L_COLUMN).value
            )

            brand_for_print = brand_m or brand_l

            record, match_source, match_details = find_part(
                raw_article,
                brand_m,
                brand_l,
                exact_map,
                article_map,
                ambiguous_articles,
            )

            if record is None:
                not_found_count += 1
                print(
                    f'{(str(index) + "/" + str(total)).ljust(12)}'
                    f'{article_for_print.ljust(24)}'
                    f'{brand_for_print.ljust(22)}'
                    f'{"".ljust(12)}'
                    f'НЕ НАЙДЕНО В БАЗЕ'
                )
                continue

            if match_source == "ARTICLE_PREFIX":
                found_by_article_brand_count += 1

            # D: заполняем только если пусто / Error... / ...not found...
            current_product_name = ws.cells(
                row_number,
                PRODUCT_NAME_COLUMN
            ).value

            if product_name_cell_needs_fill(current_product_name):
                ws.cells(
                    row_number,
                    PRODUCT_NAME_COLUMN
                ).value = record["product_name"]

            # L/M/P: существующие непустые значения не перезаписываем.
            current_brand_l = excel_value_to_string(
                ws.cells(row_number, BRAND_L_COLUMN).value
            )
            current_brand_m = excel_value_to_string(
                ws.cells(row_number, BRAND_M_COLUMN).value
            )
            current_weight = excel_value_to_string(
                ws.cells(row_number, WEIGHT_COLUMN).value
            )

            if not current_brand_l:
                ws.cells(row_number, BRAND_L_COLUMN).value = record["brand"]

            if not current_brand_m:
                ws.cells(row_number, BRAND_M_COLUMN).value = record["brand"]

            if not current_weight:
                if record["weight_kg"] is None:
                    weight_for_print = ""
                else:
                    ws.cells(
                        row_number,
                        WEIGHT_COLUMN
                    ).value = record["weight_kg"]
                    weight_for_print = excel_value_to_string(
                        record["weight_kg"]
                    )
            else:
                weight_for_print = current_weight

            found_count += 1

            print(
                f'{(str(index) + "/" + str(total)).ljust(12)}'
                f'{article_for_print.ljust(24)}'
                f'{record["brand"].ljust(22)}'
                f'{weight_for_print.ljust(12)}'
                f'{record["product_name"]}'
            )

        print("")
        print("Готово.")
        print(f"Найдено в базе и заполнено: {found_count}")
        print(f"Не найдено: {not_found_count}")
        print(f"Пустых артикулов: {empty_count}")
        print(
            f"Найдено по марке/сокращению внутри артикула: "
            f"{found_by_article_brand_count}"
        )
        print(
            f"Пропущено строк, где в D уже было нормальное наименование: "
            f"{skipped_existing_name_count}"
        )
        print("")
        print("Изменения внесены в открытую книгу Excel.")
        print("Сохранение книги остается под контролем пользователя.")

    finally:
        if connection is not None:
            connection.close()


if __name__ == "__main__":
    main()
