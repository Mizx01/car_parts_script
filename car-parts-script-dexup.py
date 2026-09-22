import sys
import re
import time
import random
import atexit
import sqlite3
from datetime import datetime
import subprocess
import importlib
import tempfile
import shutil
from pathlib import Path
from urllib.parse import unquote


# -----------------------------------------------------------------------------
# Автоматическая проверка и установка необходимых Python-пакетов
# -----------------------------------------------------------------------------
def ensure_package(import_name, pip_name=None):
    """Устанавливает пакет только если соответствующий модуль не импортируется."""
    pip_name = pip_name or import_name

    try:
        importlib.import_module(import_name)
        return
    except ImportError:
        pass

    print(f"Не найден пакет '{pip_name}'. Устанавливаем автоматически...")
    try:
        subprocess.check_call([
            sys.executable,
            "-m",
            "pip",
            "install",
            "-U",
            pip_name,
        ])
        importlib.invalidate_caches()
        importlib.import_module(import_name)
        print(f"Пакет '{pip_name}' установлен.")
    except Exception as e:
        print(f"ОШИБКА: не удалось установить пакет '{pip_name}'.")
        print(f"Причина: {e}")
        sys.exit(1)


REQUIRED_PACKAGES = [
    ("pandas", "pandas"),
    ("bs4", "beautifulsoup4"),
    ("xlwings", "xlwings"),
    ("selenium", "selenium"),
    ("psutil", "psutil"),
]

if sys.platform == "win32":
    REQUIRED_PACKAGES.append(("win32api", "pywin32"))

for _import_name, _pip_name in REQUIRED_PACKAGES:
    ensure_package(_import_name, _pip_name)


import pandas as pd
from bs4 import BeautifulSoup as BS
import xlwings as xw
import psutil

if sys.platform == "win32":
    import win32con
    import win32gui
    import win32process

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    WebDriverException,
    SessionNotCreatedException,
)


current_dir = Path(__file__).parent
wb_path = sys.argv[1]    # второй аргумент (путь к файлу) при запуске скрипта через cmd
first_cell = sys.argv[2]
first_row = int(sys.argv[3])
user_row_number = int(sys.argv[4])
last_row = first_row + user_row_number - 1
user_column = first_cell.split("$")[1]
first_cell = f'{user_column}{first_row}'
last_cell = f'{user_column}{last_row}'

# wb_path = current_dir / 'part.xlsx'
# wb_path = current_dir / 'part_PL.xlsx'

t0 = time.time()


# -----------------------------------------------------------------------------
# Накопительная база данных SQLite
# -----------------------------------------------------------------------------
# ОДНА база для всех обрабатываемых Excel-файлов.
# Она создается только если ее еще нет и затем всегда ДОПОЛНЯЕТСЯ.
DATABASE_DIR = Path(r"C:\meylis\car-parts-script")
DATABASE_PATH = DATABASE_DIR / "car-parts-db.db"

db_connection = None
db_inserted_count = 0


def init_parts_database():
    """
    Открывает единственную накопительную SQLite-базу и создает структуру
    только при первом запуске. Существующие данные НЕ удаляются.
    """
    global db_connection

    DATABASE_DIR.mkdir(parents=True, exist_ok=True)

    connection = sqlite3.connect(
        str(DATABASE_PATH),
        timeout=30,
    )

    # Обычный rollback-journal: основной файл базы остается car-parts-db.db.
    # busy_timeout дает Excel/другому читателю время освободить файл.
    connection.execute("PRAGMA journal_mode=DELETE;")
    connection.execute("PRAGMA busy_timeout=30000;")

    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS parts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            article TEXT NOT NULL,
            brand TEXT NOT NULL,
            weight_kg REAL,
            product_name TEXT NOT NULL,
            source_file TEXT NOT NULL,
            inserted_at TEXT NOT NULL
        )
        """
    )

    connection.execute(
        """
        CREATE INDEX IF NOT EXISTS idx_parts_article_brand
        ON parts(article, brand)
        """
    )

    connection.execute(
        """
        CREATE INDEX IF NOT EXISTS idx_parts_source_file
        ON parts(source_file)
        """
    )

    connection.execute(
        """
        CREATE INDEX IF NOT EXISTS idx_parts_inserted_at
        ON parts(inserted_at)
        """
    )

    # Представление для будущего Excel/VBA:
    # по article + brand возвращает самую свежую добавленную запись.
    connection.execute(
        """
        CREATE VIEW IF NOT EXISTS latest_parts AS
        SELECT
            p.id,
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
    )

    connection.commit()
    db_connection = connection
    return connection


def close_parts_database():
    """Закрывает SQLite-соединение при завершении скрипта."""
    global db_connection

    connection = db_connection
    db_connection = None

    if connection is not None:
        try:
            connection.commit()
        except Exception:
            pass

        try:
            connection.close()
        except Exception:
            pass


def get_parts_database():
    global db_connection

    if db_connection is None:
        return init_parts_database()

    return db_connection


def is_successful_parse_result(article, brand, product_name):
    """
    Разрешает запись в БД только для полноценного успешного результата.

    Не записываются:
    - пустой артикул;
    - пустая марка;
    - пустое наименование;
    - Error...;
    - Failed...;
    - No data found;
    - Not found и любые результаты, содержащие "not found".
    """
    article_text = str(article or "").strip()
    brand_text = str(brand or "").strip()
    text = str(product_name or "").strip()
    text_lower = text.lower()

    if not article_text:
        return False

    if not brand_text:
        return False

    if not text:
        return False

    if text_lower in (
        "failed to retrieve page",
        "no data found",
        "not found",
    ):
        return False

    if text_lower.startswith("error"):
        return False

    if text_lower.startswith("failed"):
        return False

    if "not found" in text_lower:
        return False

    return True


def add_part_to_database(article, brand, weight_kg, product_name, source_file):
    """
    ДОБАВЛЯЕТ новую строку в car-parts-db.db.

    UPDATE/REPLACE здесь намеренно не используется:
    старые результаты остаются в базе как история.
    """
    connection = get_parts_database()

    inserted_at = datetime.now().astimezone().isoformat(
        timespec="seconds"
    )

    cursor = connection.execute(
        """
        INSERT INTO parts (
            article,
            brand,
            weight_kg,
            product_name,
            source_file,
            inserted_at
        )
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (
            str(article or "").strip(),
            str(brand or "").strip(),
            weight_kg,
            str(product_name or "").strip(),
            str(source_file or "").strip(),
            inserted_at,
        ),
    )

    # Commit после КАЖДОЙ успешно распарсенной позиции.
    # Если дальше появится CAPTCHA/остановка, уже собранные строки сохранены.
    connection.commit()

    return cursor.lastrowid, inserted_at


atexit.register(close_parts_database)


# -----------------------------------------------------------------------------
# Доступ к Dexup через обычный Selenium WebDriver
# -----------------------------------------------------------------------------
#
# Почему эта версия надежнее предыдущей:
# - НЕ запускаем browser.exe вручную через subprocess;
# - НЕ подключаемся к браузеру через CDP/Playwright;
# - НЕ анализируем код завершения стартового процесса browser.exe;
# - Selenium сам запускает браузер и управляет его жизненным циклом;
# - Selenium Manager сам подбирает подходящий ChromeDriver / EdgeDriver.
#
# Сначала используется Google Chrome.
# Если Chrome не запускается или Dexup не отдает карточку, вторая попытка
# переключается на Microsoft Edge.
#
# Браузер НЕ headless: запускается как обычный Chrome/Edge, но окно скрывается средствами Windows.

DEXUP_HOME = "https://dexup.ru/"
DEXUP_WAIT = 25
DEXUP_PAGELOAD_TIMEOUT = 45
DEXUP_MAX_ATTEMPTS = 2

# Бережный режим после появления у ABCP защиты от автоматического парсинга.
DEXUP_DELAY_MIN = 2.0          # обычная пауза между товарами, сек.
DEXUP_DELAY_MAX = 3.0
DEXUP_BATCH_SIZE = 60          # каждые 60 позиций — длинная пауза
DEXUP_BATCH_PAUSE = 60        # 1 минута

# Если появилась CAPTCHA: сначала 5 минут, затем 15 минут.
DEXUP_CAPTCHA_COOLDOWNS = (300, 900)

BROWSER_ORDER = ("chrome", "edge")


class DexupProtectionStillActive(RuntimeError):
    """Защита Dexup осталась активной после предусмотренных пауз."""
    pass

PERSISTENT_PROFILE_DIRS = {
    "chrome": current_dir / "dexup_chrome_profile",
    "edge": current_dir / "dexup_edge_profile",
}

dexup_driver = None
dexup_browser_name = None
dexup_browser_index = 0
dexup_temp_profile_dirs = []


def _register_temp_profile(path):
    """Запоминает временный профиль для удаления после завершения скрипта."""
    path = Path(path)
    dexup_temp_profile_dirs.append(path)
    return path


def _cleanup_temp_profiles():
    """Удаляет только временные профили, созданные этим запуском."""
    for path in list(dexup_temp_profile_dirs):
        try:
            shutil.rmtree(path, ignore_errors=True)
        except Exception:
            pass


def _browser_display_name(browser_name):
    if browser_name == "chrome":
        return "Google Chrome"
    if browser_name == "edge":
        return "Microsoft Edge"
    return browser_name


def _make_browser_options(browser_name, profile_dir):
    """
    Создает Chromium-options для Chrome или Edge.

    Браузер запускается обычным НЕ-headless окном, которое затем скрывается.
    """
    if browser_name == "chrome":
        options = webdriver.ChromeOptions()
    elif browser_name == "edge":
        options = webdriver.EdgeOptions()
    else:
        raise ValueError(f"Неизвестный браузер: {browser_name}")

    options.add_argument(f"--user-data-dir={profile_dir}")
    options.add_argument("--lang=ru-RU")
    options.add_argument("--window-size=1280,900")
    options.add_argument("--window-position=-32000,-32000")
    options.add_argument("--disable-notifications")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")

    # Эти параметры были в ранее работавшей Selenium-версии.
    # Они не делают браузер headless.
    options.add_argument("--disable-blink-features=AutomationControlled")

    try:
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
    except Exception:
        pass

    return options


def _get_webdriver_process_ids(driver):
    """
    Возвращает PID chromedriver/msedgedriver и всех его дочерних процессов.
    Это позволяет не трогать обычные окна Chrome/Edge пользователя.
    """
    pids = set()

    try:
        service_process = driver.service.process
        if service_process is None:
            return pids

        root_pid = service_process.pid
        pids.add(root_pid)

        process = psutil.Process(root_pid)
        for child in process.children(recursive=True):
            pids.add(child.pid)

    except Exception:
        pass

    return pids


def _hide_webdriver_browser_windows(driver):
    """
    Скрывает только верхнеуровневые Windows-окна процессов,
    принадлежащих этому Selenium WebDriver.

    Сам браузер продолжает работать как обычный НЕ-headless Chrome/Edge.
    """
    if sys.platform != "win32" or driver is None:
        return

    pids = _get_webdriver_process_ids(driver)
    if not pids:
        return

    def callback(hwnd, _):
        try:
            _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
            if window_pid in pids and win32gui.IsWindow(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(callback, None)
    except Exception:
        pass


def _launch_browser_once(browser_name, profile_dir):
    """
    Один раз пытается запустить конкретный браузер с конкретным профилем.

    Драйвер вручную не задается: Selenium Manager подбирает его автоматически.
    """
    options = _make_browser_options(browser_name, profile_dir)

    if browser_name == "chrome":
        driver = webdriver.Chrome(options=options)
    elif browser_name == "edge":
        driver = webdriver.Edge(options=options)
    else:
        raise ValueError(f"Неизвестный браузер: {browser_name}")

    driver.set_page_load_timeout(DEXUP_PAGELOAD_TIMEOUT)

    # Окно уже создано за пределами экрана; теперь полностью скрываем его
    # средствами Windows. Браузер при этом остается обычным НЕ-headless.
    _hide_webdriver_browser_windows(driver)

    # Сохраняем то же поведение, которое помогало рабочей Chrome/Selenium версии.
    try:
        driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument",
            {
                "source": """
                    Object.defineProperty(navigator, 'webdriver', {
                        get: () => undefined
                    });
                """
            },
        )
    except Exception:
        pass

    return driver


def _launch_browser(browser_name):
    """
    Запускает браузер.

    Сначала пробует постоянный профиль.
    Если профиль занят/поврежден после прошлого запуска, автоматически
    пробует новый чистый временный профиль.
    """
    display_name = _browser_display_name(browser_name)
    persistent_profile = PERSISTENT_PROFILE_DIRS[browser_name]
    persistent_profile.mkdir(parents=True, exist_ok=True)

    errors = []

    # Попытка 1: постоянный профиль с cookies между запусками.
    try:
        print(f"Запускаем {display_name}...")
        return _launch_browser_once(browser_name, persistent_profile)
    except Exception as e:
        errors.append(f"постоянный профиль: {e}")
        print(f"{display_name}: не удалось запустить постоянный профиль.")
        print("Пробуем чистый временный профиль...")

    # Попытка 2: чистый профиль. Это обходит зависший/занятый user-data-dir.
    temp_profile = _register_temp_profile(
        tempfile.mkdtemp(prefix=f"dexup_{browser_name}_profile_")
    )

    try:
        return _launch_browser_once(browser_name, temp_profile)
    except Exception as e:
        errors.append(f"временный профиль: {e}")

    raise RuntimeError(
        f"{display_name} не удалось запустить. "
        + " | ".join(errors)
    )


def create_dexup_browser():
    """
    Запускает доступный браузер, начиная с текущего приоритета.

    Обычно:
      1) Chrome
      2) Edge
    """
    global dexup_driver, dexup_browser_name, dexup_browser_index

    last_error = None

    for index in range(dexup_browser_index, len(BROWSER_ORDER)):
        browser_name = BROWSER_ORDER[index]

        try:
            driver = _launch_browser(browser_name)

            dexup_driver = driver
            dexup_browser_name = browser_name
            dexup_browser_index = index

            # Первый заход на главную — cookies / состояние сайта.
            driver.get(DEXUP_HOME)
            _hide_webdriver_browser_windows(driver)

            try:
                WebDriverWait(driver, 15).until(
                    lambda d: d.execute_script("return document.readyState")
                    in ("interactive", "complete")
                )
            except TimeoutException:
                pass

            print(f"Dexup открыт через {_browser_display_name(browser_name)}.")
            return driver

        except Exception as e:
            last_error = e
            print("")
            print(
                f"Не удалось использовать "
                f"{_browser_display_name(browser_name)}: {e}"
            )

    raise RuntimeError(
        f"Не удалось запустить Chrome или Edge. Последняя ошибка: {last_error}"
    )


def get_dexup_browser():
    global dexup_driver

    if dexup_driver is None:
        return create_dexup_browser()

    # Проверяем, что сессия действительно жива.
    try:
        _ = dexup_driver.current_url
        return dexup_driver
    except Exception:
        close_dexup_browser()
        return create_dexup_browser()


def close_dexup_browser():
    global dexup_driver, dexup_browser_name

    driver = dexup_driver

    dexup_driver = None
    dexup_browser_name = None

    if driver is not None:
        try:
            driver.quit()
        except Exception:
            pass


atexit.register(close_dexup_browser)
atexit.register(_cleanup_temp_profiles)


def reset_dexup_browser(try_next_browser=False):
    """
    Перезапускает браузер.

    Если try_next_browser=True и сейчас использовался Chrome,
    следующая попытка идет через Edge.
    """
    global dexup_browser_index

    close_dexup_browser()
    time.sleep(1)

    if try_next_browser and dexup_browser_index < len(BROWSER_ORDER) - 1:
        dexup_browser_index += 1
        print(
            f"Повторная попытка через "
            f"{_browser_display_name(BROWSER_ORDER[dexup_browser_index])}..."
        )

    return get_dexup_browser()


def _sleep_with_progress(seconds, reason):
    """Ждет указанное время и раз в минуту показывает остаток."""
    seconds = int(seconds)
    if seconds <= 0:
        return

    print(f"{reason} Пауза {seconds // 60}:{seconds % 60:02d}.")
    remaining = seconds

    while remaining > 0:
        step = min(60, remaining)
        time.sleep(step)
        remaining -= step

        if remaining > 0:
            print(f"Осталось ждать {remaining // 60}:{remaining % 60:02d}...")


def _visible_captcha_present(driver):
    """
    Ищет именно видимую CAPTCHA/anti-bot проверку.
    Обычный подключенный на сайте скрипт reCAPTCHA сам по себе
    не считается блокировкой.
    """
    if driver is None:
        return False

    selectors = (
        "iframe[src*='recaptcha']",
        "iframe[title*='reCAPTCHA']",
        ".g-recaptcha",
        "textarea[name='g-recaptcha-response']",
        "iframe[src*='captcha']",
        "[id*='captcha']",
        "[class*='captcha']",
    )

    for selector in selectors:
        try:
            for element in driver.find_elements(By.CSS_SELECTOR, selector):
                try:
                    if element.is_displayed():
                        return True
                except Exception:
                    pass
        except Exception:
            pass

    try:
        if driver.find_elements(By.CSS_SELECTOR, ".goodsInfoDescr"):
            return False
    except Exception:
        pass

    try:
        html_lower = (driver.page_source or "").lower()
    except Exception:
        html_lower = ""

    strong_markers = (
        "подтвердите, что вы не робот",
        "подтвердите что вы не робот",
        "я не робот",
        "i'm not a robot",
        "recaptcha challenge",
    )

    return any(marker in html_lower for marker in strong_markers)


def _restart_after_captcha():
    """После CAPTCHA снова начинаем с Chrome, а не переключаемся сразу на Edge."""
    global dexup_browser_index

    close_dexup_browser()
    dexup_browser_index = 0
    time.sleep(1)
    return get_dexup_browser()


def _describe_bad_dexup_page(driver):
    """Короткая диагностика, если настоящая карточка не появилась."""
    if _visible_captcha_present(driver):
        return "Dexup показал CAPTCHA"

    try:
        title = (driver.title or "").strip()
    except Exception:
        title = ""

    try:
        html = driver.page_source or ""
    except Exception:
        html = ""

    html_lower = html.lower()

    if "403 forbidden" in html_lower or ">403<" in html_lower:
        return (
            f"{_browser_display_name(dexup_browser_name)} "
            f"получил страницу 403"
        )

    if "access denied" in html_lower:
        return (
            f"{_browser_display_name(dexup_browser_name)} "
            f"получил Access Denied"
        )

    if "captcha" in html_lower or "капча" in html_lower:
        return "Dexup показал CAPTCHA"

    if title:
        return f"карточка не появилась; title={title!r}"

    return "карточка товара не появилась"


def get_dexup_product_document(url):
    """
    Открывает полноценную карточку товара Dexup.

    Обычная ошибка может привести к fallback Chrome -> Edge.
    CAPTCHA обрабатывается иначе: запросы полностью прекращаются на cooldown,
    после чего повторяется ТА ЖЕ позиция. Если защита не снимается после
    предусмотренных пауз, весь цикл останавливается.
    """
    last_error = None
    last_url = url
    normal_attempt = 1
    captcha_cooldown_index = 0

    while normal_attempt <= DEXUP_MAX_ATTEMPTS:
        try:
            driver = get_dexup_browser()
            driver.get(url)
            _hide_webdriver_browser_windows(driver)

            WebDriverWait(driver, DEXUP_WAIT).until(
                lambda d: (
                    bool(d.find_elements(By.CSS_SELECTOR, ".goodsInfoDescr"))
                    or _visible_captcha_present(d)
                )
            )

            last_url = driver.current_url

            if _visible_captcha_present(driver):
                last_error = "Dexup показал CAPTCHA"
            elif "/parts/" not in last_url:
                last_error = f"редирект вне карточки товара: {last_url}"
            else:
                html = driver.page_source
                doc_dexup = BS(html, "html.parser")

                if doc_dexup.find(class_="goodsInfoDescr") is not None:
                    return doc_dexup

                last_error = "DOM загрузился, но goodsInfoDescr отсутствует"

        except TimeoutException:
            try:
                driver = get_dexup_browser()
                last_url = driver.current_url
                last_error = _describe_bad_dexup_page(driver)
            except Exception as e:
                last_error = f"таймаут и ошибка диагностики: {e}"

        except WebDriverException as e:
            last_error = (
                f"ошибка {_browser_display_name(dexup_browser_name)}"
                f"/WebDriver: "
                f"{e.msg if getattr(e, 'msg', None) else e}"
            )

        except Exception as e:
            last_error = str(e)

        # CAPTCHA: НЕ переключаемся сразу на Edge и НЕ продолжаем следующие товары.
        if last_error and "CAPTCHA" in last_error.upper():
            print("")
            print(f"DEXUP PROTECTION: {last_error}")
            print(f"URL: {last_url}")

            if captcha_cooldown_index >= len(DEXUP_CAPTCHA_COOLDOWNS):
                raise DexupProtectionStillActive(
                    "CAPTCHA осталась активной после всех предусмотренных пауз."
                )

            cooldown = DEXUP_CAPTCHA_COOLDOWNS[captcha_cooldown_index]
            captcha_cooldown_index += 1

            close_dexup_browser()

            _sleep_with_progress(
                cooldown,
                "Dexup включил антибот-защиту. Новые запросы временно остановлены."
            )

            print("Повторяем ту же позицию через Google Chrome...")
            _restart_after_captcha()
            continue

        # Обычная ошибка.
        print(
            f"DEXUP ERROR (attempt {normal_attempt}/{DEXUP_MAX_ATTEMPTS}): "
            f"{last_error}"
        )
        print(f"URL: {last_url}")

        if normal_attempt < DEXUP_MAX_ATTEMPTS:
            reset_dexup_browser(try_next_browser=True)
            time.sleep(2)

        normal_attempt += 1

    return None


# Функция для парсинга данных со страницы dexup
def parse_page_dexup(url):
    try:
        doc_dexup = get_dexup_product_document(url)
        if doc_dexup is None:
            return "Failed to retrieve page", None, None

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
            if property_tag and 'Вес, кг:' in property_tag.get_text(strip=True):
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

    except DexupProtectionStillActive:
        raise
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
    "SACHS YEREL" : "SACHS",
    "BLUE PRINT" : "Blue%20Print",
    "ERENBALATA" : "EREN%20BALATA",
    "HERTH+BUSS" : "H%2BB%20JAKOPARTS",
    #"MANNFILTER" : "MANN%20FILTER",
    "MANNFILTER" : "MANN-FILTER",
    "SCHMITZORG" : "SCHMITZ",
    "SSANG YONG" : "SSANG%20YONG",
    "VICTOR REI" : "VICTOR%20REINZ",
    "VİCTOR REİ" : "VICTOR%20REINZ",
    "HUTCHINSON" : "HUTCHINSON",
    "SSANGYONG" : "SSANG%20YONG",
    "BLUEPRINT" : "BLUEPRINT",
    "EUROREPAR" : "EUROREPAR",
    "GKNLOEBRO" : "GKN",
    "BORSEHUNG" : "BORSEHUNG",
    "HERTHBUSS" : "H%2BB",
    "KACMAZLAR" : "KAÇMAZLAR",
    "LEMFORDER" : "LEMFORDER",
    "NTN / SNR" : "NTN",
    "VICTORREI" : "VICTOR%20REINZ",
    "VİCTORREİ" : "VICTOR%20REINZ",
    "KRAFTVOLL" : "KRAFTVOLL",
    "CONTITECH" : "CONTITECH",
    "EURORAPE" : "EUROREPAR",
    "LMFORDER" : "LEMFORDER",
    "KOLBENSC" : "KOLBENSCHMIDT",
    "AUTOMEGA" : "AUTOMEGA",
    "BILSTEIN" : "BILSTEIN",
    "EUROREPA" : "EUROREPAR",
    "GOODYEAR" : "GOODYEAR",
    "HYD HOME" : "HYD%20HOME",
    "MBTRUCKS" : "MB%20TRUCKS",
    "PIERBURG" : "PIERBURG",
    "TEKNOROT" : "TEKNOROT",
    "VOLVOORG" : "VOLVO",
    "EUROBUMP" : "EUROBUMP",
    "CONTITEC" : "CONTITECH",
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
    "KOLBEN" : "KOLBENSCHMIDT",
    "KACMAZ" : "KACMAZLAR",
    "KAÇMAZ" : "KACMAZLAR",
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
    "AJUSA" : "AJUSA",
    "BANDO" : "BANDO",
    "BESER" : "BESER",
    "BOSCH" : "BOSCH",
    "CIFAM" : "CIFAM",
    "CONTI" : "CONTITECH",
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
    "VADEN" : "VADEN",
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
    "DEPO" : "DEPO",
    "FEBI" : "FEBI",
    "FEBİ" : "FEBI",
    "FORD" : "FORD",
    "HUCO" : "HUCO",
    "KALE" : "KALE",
    "UCEL" : "UC-EL",
    "VALS" : "VALEO",
    "VALE" : "VALEO",
    "VALA" : "VALEO",
    "VIKA" : "VIKA",
    "KAYA" : "KAYA",
    "KING" : "KING",
    "MANN" : "MANN-FILTER",
    "MARS" : "MARS",
    "MAIS" : "RENAULT",
    "MEHA" : "MEHA",
    "MİBA" : "MIBA",
    "TRSN" : "TIRSAN",
    "GRAF" : "GRAF",
    "ONKA" : "ONKA",
    "OPEL" : "OPEL",
    "SWAG" : "SWAG",
    "AIS" : "AISIN",
    "ASP" : "ASPOCK",
    "FOR" : "FORD",
    "AYF" : "AYFAR",
    "GMB" : "GMB",
    "BAN" : "BANDO",
    "BCH" : "BOSCH",
    "BER" : "BERU",
    "BIL" : "BILSTEIN",
    "BLU" : "BLUE%20PRINT",
    "BLP" : "BLUE%20PRINT",
    "BOS" : "BOSCH",
    "BSH" : "BOSCH",
    "BMW" : "BMW",
    "BRB" : "BREMBO",
    "BRE" : "BREMBO",
    "BRU" : "BERU",
    "BRS" : "BORSEHUNG",
    "BSC" : "BOSCH",
    "CIF" : "CIFAM",
    "CNT" : "CONTITECH",
    "CNR" : "CNR",
    "CTR" : "CTR",
    "CON" : "CONTITECH",
    "COR" : "CORTECO",
    "CRT" : "CORTECO",
    "DAY" : "DAYCO",
    "DEL" : "DELPHI",
    "DEN" : "DENSO",
    "DEG" : "DE-GA",
    "DEGA" : "DE-GA",
    "DNS" : "DENSO",
    "DPO" : "DEPO",
    "DOL" : "DOLZ",
    "ECE" : "ECEM",
    "ELR" : "ELRING",
    "ERA" : "ERA",
    "EUR" : "EUROBUMP",
    "FAC" : "FACET",
    "FBI" : "FEBI",
    "FEB" : "FEBI",
    "FLT" : "FILTRON",
    "FIA" : "Fiat%2FAlfa%2FLancia",
    "FOR" : "FORD",
    "FRD" : "FORD",
    "GAT" : "GATES",
    "GKN" : "GKN",
    "GLY" : "GLYCO",
    "GTS" : "GATES",
    "GSP" : "GSP",
    "GVA" : "GVA",
    "KSC" : "KOLBENSCHMIDT",
    "HEL" : "HELLA",
    "REMSA" : "REMSA",
    "HOL" : "HOLSET",
    "HLL" : "HELLA",
    "HNG" : "HENGST",
    "INA" : "INA",
    "KOL" : "KOLBENSCHMIDT",
    "KNG" : "KONGSBERG",
    "LEM" : "LEMFORDER",
    "LMF" : "LEMFORDER",
    "MAH" : "MAHLE",
    "MAI" : "RENAULT",
    "MAN" : "MANN-FILTER",
    "MAY" : "MAYSAN%20MANDO",
    "MER" : "MERCEDES-BENZ",
    "MHL" : "MAHLE",
    "MON" : "MONROE",
    "MTA" : "MTA",
    "MTL" : "METELLI",
    "NGK" : "NGK",
    "NIS" : "NISSENS",
    "NıS" : "NISSENS",
    "NRV" : "NARVA",
    "KRA" : "KRAFTVOLL",
    "OPL" : "OPEL",
    "OPT" : "OPTIMAL",
    "OSM" : "OSRAM",
    "OSR" : "OSRAM",
    "ULO" : "ULO",
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
    "TPR" : "TOPRAN",
    "TRW" : "TRW",
    "VAL" : "VALEO",
    "VCT" : "VICTOR%20REINZ",
    "VER" : "VERNET",
    "VIK" : "VIKA",
    "YEN" : "YENMAK",
    "DAF" : "DAF",
    "DYC" : "DAYCO",
    "FAE" : "FAE",
    "FAG" : "FAG",
    "FRJ" : "FIAT",
    "KRF" : "KRAFTVOLL",
    "KSC" : "KOLBENSCHMIDT",
    "LPR" : "LPR",
    "FMN" : "FMN",
    "FSE" : "FASE",
    "FTE" : "FTE",
    "GMB" : "GMB",
    "KLR" : "KALE",
    "KAL" : "KALE",
    "KYB" : "KYB",
    "KNG" : "KONGSBERG",
    "LUK" : "LUK",
    "MGA" : "AUTOMEGA",
    "MMA" : "MAGNETI%20MARELLI",
    "MND" : "MANDO",
    "MNN" : "MANN-FILTER",
    "MAP" : "MAPA",
    "MER" : "MERCEDES-BENZ",
    "MKS" : "MKS",
    "MGA" : "MGA",
    "NRF" : "NRF",
    "NTN" : "NTN",
    "PAY" : "PAYEN",
    "OES" : "OES",
    "GNS" : "GUNES",
    "IBR" : "IBRAS",
    "IVE" : "IVECO",
    "OTO" : "OTO",
    "POJ" : "PEUGEOT%20CITROEN",
    "PEU" : "PEUGEOT%20CITROEN",
    "PRB" : "PIERBURG",
    "PRG" : "PIERBURG",
    "PSA" : "PEUGEOT-CITROEN",
    "CAV" : "CAVO",
    "REN" : "RENAULT",
    "RYL" : "ROYAL",
    "MAR" : "MARS",
    "SKT" : "SKT",
    "SMP" : "SAMPART",
    "SWF" : "SWF",
    "TXT" : "TEXTAR",
    "TUR" : "TURTEL",
    "TIR" : "TIRSAN",
    "UFI" : "UFI",
    "WIN" : "WIN",
    "WOD" : "WOD",
    "WHL" : "WAHLER",
    "YTT" : "YTT",
    "YNM" : "YENMAK",
    "VAG" : "VAG",
    "ORJ" : "VAG",
    "K&B" : "K&B",
    #"ORJ" : "ORIGINAL",
    #"ORJ" : "PEUGEOT-CITROEN",
    "TX" : "TEXTAR",
    "BR" : "BERU",
    "GM" : "GENERAL MOTORS",
    "VR" : "VICTOR%20REINZ",
    "LF" : "LEMFORDER",
    "MB" : "MERCEDES%20BENZ",
    "ZF_" : "LEMFORDER",
    "ZFT" : "ZF"
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
print(f"Парсим сайт www.dexup.ru")
print("")

try:
    get_parts_database()
    existing_db_rows = db_connection.execute(
        "SELECT COUNT(*) FROM parts"
    ).fetchone()[0]
    print(f"База данных: {DATABASE_PATH}")
    print(f"Записей в базе до запуска: {existing_db_rows}")
except Exception as e:
    print("")
    print("ОШИБКА БАЗЫ ДАННЫХ.")
    print(f"Файл: {DATABASE_PATH}")
    print(f"Причина: {e}")
    print("Парсинг не запущен, чтобы результаты не остались только в Excel.")
    sys.exit(1)

print("")

# Запускаем браузер заранее, чтобы сообщения
# "Запускаем Google Chrome...", "Dexup открыт..." и стартовый fallback на Edge
# выводились до заголовка таблицы.
try:
    get_dexup_browser()
except Exception as e:
    print(f"ОШИБКА ЗАПУСКА БРАУЗЕРА: {e}")

print("")
print(f'{"Позиция".ljust(10):7}{"Артикул".ljust(20):15}{"Марка".ljust(20):15}{"Вес, кг".ljust(12):12}Наименование')

#if total_rows == 1:


#print(data)

# Проходим по строкам файла и парсим данные
stopped_by_dexup_protection = False
stopped_position = None

stopped_by_database_error = False
database_error_message = None

skipped_existing_name_count = 0

for row_index, row in enumerate(data):

    current_row = row_index + first_row

    # Нормальное непустое наименование в D => строку полностью пропускаем.
    # Пустое D, Error... или ...not found... => строку обрабатываем заново.
    existing_product_name = excel_value_to_string(
        sht.range(current_row, 4).value
    ).strip()

    if not product_name_cell_needs_fill(existing_product_name):
        skipped_existing_name_count += 1

        pos_index = (
            str(row_index + 1) + "/" + str(total_rows)
        ).ljust(10)

        raw_article_for_print = excel_value_to_string(row).upper()

        print(
            f'{pos_index:7}'
            f'{raw_article_for_print.ljust(20):15}'
            f'{"".ljust(20):15}'
            f'{"".ljust(12):12}'
            f'ПРОПУЩЕНО: в D уже есть нормальное наименование'
        )
        continue

    #if all(cell is None or cell == '' for cell in row):
     #   continue #пропускаем пустые строки

    #print(f"row_index:{row_index}, row:{row}")
    
    #raw_art = row.strip()
    raw_art = excel_value_to_string(row).upper() #в будущем надо будет использовать индексы если данные неодномерные

    # Пустую строку артикула не парсим и в БД не записываем.
    if not raw_art.strip():
        pos_index = (
            str(row_index + 1) + "/" + str(total_rows)
        ).ljust(10)

        print(
            f'{pos_index:7}'
            f'{"".ljust(20):15}'
            f'{"".ljust(20):15}'
            f'{"".ljust(12):12}'
            f'ПРОПУЩЕНО: пустой артикул'
        )
        continue

    #raw_art = str(row[0]).strip()  # все данные в row 
    #print(f"raw_art: {raw_art}" )
    new_art = raw_art
    marka = None
    #print(marka)
    #marka = str(row[1]).strip()  # марка из второго столбца

    original_proiz = ((sht.range(row_index + first_row, 12).value) or "").upper()   # or "" - чтобы не выдавал ошибку из-за None
    original_marka = ((sht.range(row_index + first_row, 13).value) or "").upper()
    #print(f"proiz: {proiz}, marka: {marka}")
    original_proiz = original_proiz.strip()
    original_marka = original_marka.strip() 


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

    if marka == "Fiat%2FAlfa%2FLancia":
        marka= "FIAT"


        
 

    # Очистка артикула
    art = clean_artikul(new_art)
    
    # Проверка наличия марки в словаре и замена, если необходимо
    #if marka in brand_replacement:
    #    marka = brand_replacement[marka]
    
    #Формирование URL и парсинг данных с dexup
    url_dexup = f"https://dexup.ru/parts/{marka}/{art}"
    
    # Вызов функции парсинга страницы dexup
    try:
        product_name_dexup, mass_dexup, material_dexup = parse_page_dexup(url_dexup)
    except DexupProtectionStillActive as e:
        stopped_by_dexup_protection = True
        stopped_position = row_index + 1

        print("")
        print("=" * 80)
        print("ОБРАБОТКА ОСТАНОВЛЕНА: Dexup продолжает показывать CAPTCHA.")
        print(f"Позиция: {stopped_position}/{total_rows}")
        print(f"Артикул: {art}")
        print(f"URL: {url_dexup}")
        print(f"Причина: {e}")
        print("Остальные позиции намеренно НЕ запрашиваются.")
        print("Позже запустите обработку снова, начиная с этой позиции.")
        print("=" * 80)
        break

    if "MAIS" in raw_art:                           # чтобы марка MAIS на сайте dexup искал по марке RENAULT,
        marka = "MAIS"                              # а сама марка MAIS осталась в пакинге.
    
    
    # Приведение данных к корректному регистру
    product_name_dexup = capitalize_first_letter(product_name_dexup)
    material_dexup = capitalize_first_letter(material_dexup)

    # decode marka, убирается лишнее из марки перед тем как вставлять в эксель файл
    marka = unquote(marka or "")

    pos_index = (str(row_index + 1) + '/' + str(total_rows)).ljust(10)    #номер позиции

    if mass_dexup is None:
        mass_for_print = ""
    else:
        mass_for_print = str(mass_dexup)

    print(f'{pos_index:7}{art.ljust(20):15}{marka.ljust(20):15}{mass_for_print.ljust(12):12}{product_name_dexup}')
    




    # Заполняем только разрешенные/пустые значения.
    current_product_name = sht.range(current_row, 4).value
    current_brand_l = sht.range(current_row, 12).value
    current_brand_m = sht.range(current_row, 13).value
    current_weight = sht.range(current_row, 16).value

    # D: пусто / Error... / ...not found... можно заменить,
    # но только если текущий результат парсинга успешный.
    if (
        product_name_cell_needs_fill(current_product_name)
        and is_successful_parse_result(
            article=art,
            brand=marka,
            product_name=product_name_dexup,
        )
    ):
        sht.range(current_row, 4).value = [
            product_name_dexup
        ]

    # L/M/P: существующие непустые значения никогда не перезаписываем.
    if not excel_value_to_string(current_brand_l).strip():
        sht.range(current_row, 12).value = marka

    if not excel_value_to_string(current_brand_m).strip():
        sht.range(current_row, 13).value = marka

    if not excel_value_to_string(current_weight).strip():
        sht.range(current_row, 16).value = mass_dexup

    # ----------------------------------------------------------------------
    # Накопительная SQLite-база.
    # В нее идут только успешно полученные карточки.
    # Каждый новый запуск/Excel-файл ДОБАВЛЯЕТ новые строки в car-parts-db.db.
    # ----------------------------------------------------------------------
    if is_successful_parse_result(
        article=art,
        brand=marka,
        product_name=product_name_dexup,
    ):
        try:
            add_part_to_database(
                article=art,
                brand=marka,
                weight_kg=mass_dexup,
                product_name=product_name_dexup,
                source_file=Path(wb_path).name,
            )
            db_inserted_count += 1

        except Exception as e:
            stopped_by_database_error = True
            database_error_message = str(e)

            print("")
            print("=" * 80)
            print("ОБРАБОТКА ОСТАНОВЛЕНА: ошибка записи в базу данных.")
            print(f"База: {DATABASE_PATH}")
            print(f"Позиция: {row_index + 1}/{total_rows}")
            print(f"Артикул: {art}")
            print(f"Причина: {database_error_message}")
            print(
                "Дальнейший парсинг остановлен, чтобы Excel и база "
                "не расходились."
            )
            print("=" * 80)
            break


    #sht.range(row_index + first_row, 4).value = [
    #    product_name_dexup,
    #    mass_dexup,
    #    decoded_marka,
    #    material_dexup,
    #    url_dexup,
    #]


    # Бережный режим: не отправляем карточки подряд слишком быстро.
    processed_count = row_index + 1

    if processed_count < total_rows:
        if processed_count % DEXUP_BATCH_SIZE == 0:
            _sleep_with_progress(
                DEXUP_BATCH_PAUSE,
                f"Обработано {processed_count} позиций. Плановая пауза."
            )
        else:
            time.sleep(random.uniform(DEXUP_DELAY_MIN, DEXUP_DELAY_MAX))
    
# Сохраняем обновленный Excel файл
#active_sheet.range("A1").value = data
print("")
if stopped_by_database_error:
    print("Обработка остановлена из-за ошибки базы данных.")
elif stopped_by_dexup_protection:
    print(
        f"Обработка остановлена на позиции {stopped_position}/{total_rows} "
        f"из-за активной защиты Dexup."
    )
else:
    print("Данные успешно сохранены в файл:", wb_path)

print(f"В базу данных добавлено записей: {db_inserted_count}")
print(
    f"Пропущено строк, где в D уже было нормальное наименование: "
    f"{skipped_existing_name_count}"
)
print(f"База данных: {DATABASE_PATH}")

t1 = time.time()
print("Процесс занял", round((t1 - t0)/60), "минут")
print("")
