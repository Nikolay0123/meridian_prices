import re
import time
import random
from datetime import datetime, timedelta, date
from typing import Optional

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service


# =========================
# НАСТРОЙКИ ПОЛЬЗОВАТЕЛЯ
# =========================
# В формате ГГГГ-ММ-ДД
START_DATE = "2026-03-20"
END_DATE = "2026-03-25"

# Телефоны отеля для уточнения (по ТЗ)
PHONE_1 = "+7 (916) 643-72-15"
PHONE_2 = "+7 (925) 337-72-07"

BASE_URL = "https://hotel-meridian.com/"
BOOKING_BASE_URL = "https://meridian.bookonline24.ru/"
ADULTS_COUNT = 1  # adultsCount в URL бронирования

# Список страниц категорий, которые вы указали.
# Именно с них извлекаем базовые цены для fallback-режима.
CATEGORY_PAGES = [
    "https://hotel-meridian.com/nomera/standartnyy-klassicheskiy2/",
    "https://hotel-meridian.com/nomera/standartnyy-uluchshennyy/",
    "https://hotel-meridian.com/nomera/standart-uluchshennyy-s-parkingom/",
    "https://hotel-meridian.com/nomera/standart-komfort/",
    "https://hotel-meridian.com/nomera/semeynyy-dvukhkomnatnyy/",
    "https://hotel-meridian.com/nomera/standart-komfort-uluchshennyy-plyus/",
]

# На странице виджета бронирования встречаются заголовки типа "ЗАЕЗД"/"ВЫЕЗД".
# Они не являются категориями номеров, поэтому исключаем их при попытках
# извлечь название категории.
EXCLUDED_TITLE_TOKENS = {
    "ЗАЕЗД",
    "ВЫЕЗД",
    "КОЛИЧЕСТВО ГОСТЕЙ",
    "ЗАБРОНИРОВАТЬ",
}

# Дополнительные отели: название, URL или шаблон, кнопка для клика (если нужна).
# cookie_button: нажать перед основной кнопкой при согласии на cookie.
# url_template: для подстановки дат используйте {dfrom} и {dto} в формате ДД-ММ-ГГГГ.
# button_in_iframe: искать кнопку внутри iframe (например Tilda).
# wait_seconds: дополнительная задержка после загрузки (для SPA).
HOTEL_SOURCES = [
    {
        "name": "Олимпийская",
        "type": "multi_page",
        "urls": [
            "https://olympik-hotel.ru/?view=page&id=2",
            "https://olympik-hotel.ru/?view=page&id=14",
            "https://olympik-hotel.ru/?view=page&id=7",
            "https://olympik-hotel.ru/?view=page&id=18",
        ],
    },
    {
        "name": "Freezone inn",
        "type": "button",
        "url": "https://www.freezone.net/hotel/",
        "button": "найти номер",
        "button_alt": ["Забронировать", "найти", "поиск", "Проверить наличие", "check availability"],
        "cookie_button": "согласен",
    },
    {
        "name": "Постоялый двор Русь",
        "type": "direct",
        "url": "https://booking-russ.otelms.com/booking/rooms",
        "wait_seconds": 6,
    },
    {
        "name": "Чехов API",
        "type": "date_button",
        "url_template": "https://chekhov-api.tilda.ws/booking?dfrom={dfrom}&dto={dto}&adults=1&padding=12&lang=ru&uid=53b6c90b-227a-48cf-8339-c85954fab29e",
        "button": "подобрать номер",
        "button_alt": ["Подобрать номер", "подобрать", "найти номер", "поиск номеров", "Проверить наличие", "Search", "Найти"],
        "button_in_iframe": True,
        "iframe_wait_seconds": 8,
        "cookie_button": "OK",
    },
    {
        "name": "Чеховский мини-отель",
        "type": "direct",
        "url": "https://hotelchehov.ru/",
    },
]


def human_sleep(a: float = 1.0, b: float = 2.2) -> None:
    """Случайная задержка для уменьшения нагрузки и для стабильности."""
    time.sleep(random.uniform(a, b))


def parse_ymd(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()


def format_for_picker(d: date) -> str:
    """Чаще всего виджеты на сайте принимают формат ДД.ММ.ГГГГ."""
    return d.strftime("%d.%m.%Y")


def format_for_booking_url(d: date) -> str:
    """Формат дат в URL бронирования: ДД.ММ.ГГГГ."""
    return d.strftime("%d.%m.%Y")


def format_date_dd_mm_yyyy(d: date) -> str:
    """Формат ДД-ММ-ГГГГ для URL (например Чехов API)."""
    return d.strftime("%d-%m-%Y")


def to_int_rub(text: str) -> Optional[int]:
    """
    Пытаемся извлечь число в рублях.
    Примеры входа: "от 14 000 рублей", "1 800,00 рублей".
    """
    if not text:
        return None

    # Схватим первое число и (опционально) дробную часть.
    # Учитываем пробелы/запятые в тысячах (34,000.00 или 34 000) и запятую/точку в дробной части.
    m = re.search(r"(\d[\d\s,]*)([.,]\d+)?\s*(?:руб|₽|рублей|RUB)", text, flags=re.IGNORECASE)
    if not m:
        # Если "руб"/RUB рядом нет, попробуем просто числа.
        m2 = re.search(r"(\d[\d\s,]*)([.,]\d+)?", text)
        if not m2:
            return None
        m = m2

    int_part = m.group(1).replace(" ", "").replace(",", "")
    dec_part = m.group(2) or ""

    if dec_part:
        # "1800,00" -> float
        val = float(int_part + dec_part.replace(",", "."))
        return int(round(val))
    return int(int_part)


def booking_url(from_date: date, to_date: date, adults_count: int = ADULTS_COUNT) -> str:
    """
    Генерируем URL бронирования, где выбираются даты и показываются цены по всем категориям.
    Пример: https://meridian.bookonline24.ru/?fromDate=22.03.2026&toDate=23.03.2026&adultsCount=1
    """
    fd = format_for_booking_url(from_date)
    td = format_for_booking_url(to_date)
    return f"{BOOKING_BASE_URL}?fromDate={fd}&toDate={td}&adultsCount={adults_count}"


def wait_booking_page_ready(driver: webdriver.Chrome, timeout_s: int = 30) -> None:
    """
    Ждем, пока страница бронирования отрисуется.
    Если попали на страницу техработ, это не упадет сразу — дальше парсер вернет пусто,
    и мы уйдем в fallback.
    """
    wait = WebDriverWait(driver, timeout_s)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    # Даем странице шанс дорисовать динамику.
    human_sleep(1.5, 2.7)


def click_check_availability_and_wait_prices(driver: webdriver.Chrome, timeout_s: int = 25) -> bool:
    """
    На странице бронирования цены появляются только после клика «Проверить наличие».
    Ищем кнопку/ссылку с таким текстом, кликаем и ждем появления цен.
    """
    btn = None
    # XPath: элемент (a, button, input, span, div), в тексте которого есть "проверить" и "наличие"
    for xpath in [
        "//a[contains(., 'проверить') and contains(., 'наличие')]",
        "//button[contains(., 'проверить') or contains(., 'наличие')]",
        "//*[@type='submit'][contains(@value, 'наличие') or contains(@value, 'проверить')]",
        "//a[contains(., 'наличие')]",
        "//button[contains(., 'наличие')]",
        "//*[contains(., 'Проверить наличие')]",
    ]:
        try:
            els = driver.find_elements(By.XPATH, xpath)
            for el in els:
                if el.is_displayed() and el.is_enabled():
                    btn = el
                    break
            if btn:
                break
        except NoSuchElementException:
            continue
    if not btn:
        return False

    try:
        human_sleep(0.5, 1.0)
        btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", btn)

    # Ждем загрузки: даем время на запрос и отрисовку цен
    human_sleep(2.5, 5.0)
    return True


def click_button_by_text(driver: webdriver.Chrome, button_text: str, wait_after_s: float = 2.5) -> bool:
    """
    Универсальный клик по кнопке/ссылке с заданным текстом (частичное совпадение).
    """
    if not button_text or not button_text.strip():
        return False
    t = button_text.strip().replace("'", " ")
    for xpath in [
        f"//a[contains(., '{t[:50]}')]",
        f"//button[contains(., '{t[:50]}')]",
        f"//*[@role='button'][contains(., '{t[:50]}')]",
        f"//*[@type='submit'][contains(@value, '{t[:30]}')]",
        f"//*[@type='button'][contains(@value, '{t[:30]}')]",
        f"//input[@type='submit' and contains(@value, '{t[:30]}')]",
        f"//*[contains(., '{t[:30]}')]",
    ]:
        try:
            els = driver.find_elements(By.XPATH, xpath)
            for el in els:
                if el.is_displayed() and el.is_enabled():
                    try:
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                        human_sleep(0.2, 0.4)
                    except Exception:
                        pass
                    try:
                        el.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", el)
                    human_sleep(wait_after_s, wait_after_s + 1.5)
                    return True
        except Exception:
            continue
    return False


def accept_cookie_if_present(driver: webdriver.Chrome, text: str = "согласен", wait_after_s: float = 1.0) -> bool:
    """
    Если на странице есть кнопка/ссылка согласия на cookie (например «Согласен»), нажимаем её.
    """
    if not text or not text.strip():
        return False
    t = text.strip().replace("'", " ")
    for xpath in [
        f"//a[contains(., '{t}')]",
        f"//button[contains(., '{t}')]",
        f"//*[@type='button'][contains(., '{t}')]",
        f"//*[contains(., 'cookie') and contains(., '{t[:10]}')]",
        f"//*[contains(., 'Cookie') and contains(., '{t[:10]}')]",
    ]:
        try:
            els = driver.find_elements(By.XPATH, xpath)
            for el in els:
                if el.is_displayed() and el.is_enabled():
                    try:
                        el.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", el)
                    human_sleep(wait_after_s, wait_after_s + 0.5)
                    return True
        except Exception:
            continue
    return False


def click_button_in_iframe_then_page(
    driver: webdriver.Chrome,
    button_text: str,
    wait_after_s: float = 3.0,
    button_alt: Optional[list] = None,
    iframe_wait_seconds: float = 4.0,
) -> bool:
    """
    Сначала ждём загрузки iframe, ищем кнопку во всех iframe (виджет Tilda/Bnovo), затем на основной странице.
    button_alt — список альтернативных текстов кнопки для перебора.
    """
    human_sleep(iframe_wait_seconds, iframe_wait_seconds + 1.0)
    texts_to_try = [button_text]
    if button_alt:
        texts_to_try = [button_text] + [t for t in button_alt if t and t != button_text]

    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    for ifr in iframes:
        try:
            driver.switch_to.frame(ifr)
            for txt in texts_to_try:
                if click_button_by_text(driver, txt, wait_after_s=wait_after_s):
                    driver.switch_to.default_content()
                    return True
        except Exception:
            pass
        driver.switch_to.default_content()

    for txt in texts_to_try:
        if click_button_by_text(driver, txt, wait_after_s=wait_after_s):
            return True
    return False


def scrape_prices_from_body_text(driver: webdriver.Chrome) -> list:
    """
    Фолбэк для страниц, где цены не находятся стандартными селекторами.
    Ищем строки в `body.text`, где встречается валюта, парсим цену и
    пытаемся взять название из ближайшей строки выше.
    """
    try:
        body_text = driver.find_element(By.TAG_NAME, "body").text or ""
    except NoSuchElementException:
        body_text = ""

    lines = [ln.strip() for ln in body_text.splitlines() if ln.strip()]
    if not lines:
        return []

    results = []
    seen_categories = set()

    def has_currency(s: str) -> bool:
        s_low = (s or "").lower()
        return ("руб" in s_low) or ("₽" in s_low) or ("RUB" in (s or "").upper())

    for i, ln in enumerate(lines):
        if not has_currency(ln):
            continue
        price = to_int_rub(ln)
        if price is None:
            continue

        cat = None
        # Ищем ближайшую "категорию" в предыдущих 5 строках.
        for j in range(max(0, i - 5), i):
            prev = lines[j]
            if has_currency(prev):
                continue
            if "номер" in prev.lower():
                cat = prev
                break

        if not cat:
            # Если нет "Номер", берём предыдущую строку (или запасной вариант).
            cat = lines[i - 1] if i > 0 else "Номер"

        if cat in seen_categories:
            continue
        seen_categories.add(cat)
        results.append({"category_name": cat, "price": price})

    return results


def scrape_prices_from_page_source_html(driver: webdriver.Chrome) -> list:
    """
    Экстренный фолбэк: достаём цены прямо из page_source (после удаления тегов).
    Используется, когда body.text не отдаёт нужный текст.
    """
    html = driver.page_source or ""
    if not html:
        return []

    # Удаляем теги, чтобы цена/валюта оказались рядом в "plain" тексте.
    plain = re.sub(r"<[^>]+>", " ", html)
    plain = " ".join(plain.split())

    # Пытаемся вытащить все суммы с валютой.
    price_pat = re.compile(r"(\d[\d\s,]*)([.,]\d+)?\s*(руб|₽|рублей|RUB)", flags=re.IGNORECASE)

    results = []
    seen_prices = set()
    for m in price_pat.finditer(plain):
        chunk = m.group(0)
        price = to_int_rub(chunk)
        if price is None or price in seen_prices:
            continue
        seen_prices.add(price)
        results.append({"category_name": "Номер", "price": price})

    return results


def scrape_prices_generic(driver: webdriver.Chrome) -> list:
    """
    Универсальный сбор категорий и цен со страницы (та же логика, что у Meridian).
    Возвращает список {"category_name": str, "price": int}.
    """
    # 1) Пробуем собрать цены в текущем (основном) DOM.
    results = scrape_prices_from_booking_page(driver)
    if results:
        return results

    # 2) Иногда виджет бронирования (Bnovo/OtelMS) рендерит цены в iframe.
    #    Тогда в основном DOM может не быть ни "руб", ни "₽", ни "RUB".
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    for iframe in iframes:
        try:
            driver.switch_to.frame(iframe)
            iframe_results = scrape_prices_from_booking_page(driver)
            if iframe_results:
                return iframe_results
        except Exception:
            pass
        finally:
            try:
                driver.switch_to.default_content()
            except Exception:
                pass

    # 3) Последние фолбэки: извлекаем цены напрямую.
    body_results = scrape_prices_from_body_text(driver)
    if body_results:
        return body_results

    page_source_results = scrape_prices_from_page_source_html(driver)
    if page_source_results:
        return page_source_results

    return results


def scrape_one_category_per_page(driver: webdriver.Chrome) -> Optional[dict]:
    """
    Одна страница — одна категория (например Олимпийская: каждая ссылка id=2,14,7,18 — отдельная категория).
    Сначала ищем блок, в котором есть цена (руб), затем в этом блоке — название категории (заголовок или первая строка).
    """
    try:
        body_text = driver.find_element(By.TAG_NAME, "body").text or ""
    except NoSuchElementException:
        return None
    price = to_int_rub(body_text)
    if price is None:
        return None

    name = None
    # Ищем элемент с ценой (руб), поднимаемся к контейнеру и в нём ищем заголовок/название категории
    try:
        price_els = driver.find_elements(
            By.XPATH,
            "//*[contains(., 'руб') or contains(., '₽')]",
        )
        for el in price_els:
            txt = (el.text or "").strip()
            if to_int_rub(txt) != price:
                continue
            # Контейнер: div, section, article, main
            for tag in ["div", "section", "article", "main"]:
                try:
                    block = el.find_element(By.XPATH, f"./ancestor::{tag}[1]")
                    block_text = (block.text or "").strip()
                    if not block_text or "руб" not in block_text.lower():
                        continue
                    # В блоке ищем заголовок (h1, h2, h3) или первую осмысленную строку
                    for sel in ["h1", "h2", "h3", "h4"]:
                        try:
                            h = block.find_element(By.CSS_SELECTOR, sel)
                            t = (h.text or "").strip()
                            if t and len(t) >= 2 and t.upper() not in EXCLUDED_TITLE_TOKENS and "руб" not in t.lower():
                                name = t
                                break
                        except NoSuchElementException:
                            continue
                    if name:
                        break
                    name = extract_room_name_from_block_text(block_text)
                    if name:
                        break
                except NoSuchElementException:
                    continue
            if name:
                break
    except Exception:
        pass

    # Fallback: заголовки страницы по порядку (часто первый — логотип, второй — категория)
    if not name:
        for sel in ["h2", "h3", "h1", "h4"]:
            try:
                els = driver.find_elements(By.CSS_SELECTOR, sel)
                for el in els:
                    t = (el.text or "").strip()
                    if t and len(t) >= 2 and t.upper() not in EXCLUDED_TITLE_TOKENS and "руб" not in t.lower():
                        name = t
                        break
            except NoSuchElementException:
                continue
            if name:
                break

    if not name:
        name = "Номер"
    return {"category_name": name, "price": price}


def extract_room_name_from_block_text(block_text: str) -> Optional[str]:
    """
    Пытаемся вытащить название категории из текста блока:
    берем первую "осмысленную" строку без служебных слов.
    """
    if not block_text:
        return None

    lines = [ln.strip() for ln in block_text.splitlines() if ln.strip()]
    for ln in lines[:15]:
        upper = ln.upper()
        if upper in EXCLUDED_TITLE_TOKENS:
            continue
        if "ЗАЕЗД" in upper or "ВЫЕЗД" in upper or "КОЛИЧЕСТВО" in upper:
            continue
        if "РУБ" in upper or "RUB" in ln.upper() or "₽" in ln:
            continue
        # отсечем слишком короткое
        if len(ln) < 3:
            continue
        return ln
    return None


def scrape_prices_from_booking_page(driver: webdriver.Chrome) -> list:
    """
    Собираем цены по всем категориям со страницы meridian.bookonline24.ru.

    Так как разметка может быть динамической/меняться, используем устойчивую эвристику:
    - находим "листовые" элементы, содержащие руб/₽
    - для каждого элемента поднимаемся к ближайшему контейнеру и извлекаем из него название
    - дедуплицируем по названию, берем первую найденную цену
    """
    # Быстрый признак техработ (на случай, если страница реально не отдает контент)
    try:
        body_text = driver.find_element(By.TAG_NAME, "body").text
        if "updating our service" in body_text.lower():
            return []
    except NoSuchElementException:
        return []

    # Ищем элементы с ценой.
    # На OtelMS цена лежит внутри `span` с дочерними элементами (например help-icon),
    # поэтому ограничение `not(*)` ломает извлечение (элемент не становится "leaf node").
    # Учитываем руб/₽ (кириллица) и RUB (латиница, например OtelMS).
    price_els = driver.find_elements(
        By.XPATH,
        # normalize case for Cyrillic "РУБ" -> "руб"
        "//*[not(self::script) and not(self::style) and (contains(translate(., 'РУБ', 'руб'), 'руб') or contains(., '₽') or contains(., 'RUB'))]",
    )

    results = {}
    for el in price_els:
        txt = (el.text or "").strip()
        # Сильно ограничим мусор: если нет цифр рядом с валютой — to_int_rub вернет None,
        # но это уже фильтрация на раннем этапе (ускоряет и повышает стабильность).
        if not txt or not any(ch.isdigit() for ch in txt):
            continue
        price = to_int_rub(txt)
        if price is None:
            continue

        # Поднимаемся к контейнеру (ограничим глубину, чтобы не схватить всю страницу)
        block = None
        try:
            block = el.find_element(By.XPATH, "./ancestor::*[self::div or self::li or self::section][1]")
        except NoSuchElementException:
            block = None

        block_text = ""
        if block is not None:
            block_text = (block.text or "").strip()
        else:
            # fallback: используем текст родителя
            try:
                parent = el.find_element(By.XPATH, "./parent::*")
                block_text = (parent.text or "").strip()
            except NoSuchElementException:
                block_text = txt

        # Пытаемся найти "заголовок" внутри контейнера
        name = None
        if block is not None:
            for sel in ["h1", "h2", "h3", "h4", "h5", "h6"]:
                hs = block.find_elements(By.CSS_SELECTOR, sel)
                for h in hs:
                    ht = (h.text or "").strip()
                    if not ht:
                        continue
                    hu = ht.upper()
                    if hu in EXCLUDED_TITLE_TOKENS:
                        continue
                    if "ЗАЕЗД" in hu or "ВЫЕЗД" in hu or "КОЛИЧЕСТВО" in hu:
                        continue
                    name = ht
                    break
                if name:
                    break

        if not name:
            name = extract_room_name_from_block_text(block_text)

        if not name:
            continue

        # Дедуп по названию: берем первую (обычно минимальная/основная цена)
        if name not in results:
            results[name] = price

    # Превращаем в список
    out = [{"category_name": k, "price": v} for k, v in results.items()]
    return out


def extract_category_title(body_text: str) -> Optional[str]:
    """
    На страницах категорий встречается фраза:
    Номер первой категории "Стандарт Классический"
    """
    m = re.search(r'Номер первой категории\s*[\"“](.*?)[\"”]', body_text)
    if m:
        return m.group(1).strip()
    return None


def get_category_title_from_dom(driver: webdriver.Chrome) -> Optional[str]:
    """
    Надежно достаем название категории из DOM.

    Используем ключевую фразу "Номер первой категории ..." и избегаем
    попадания в заголовки виджета типа "ЗАЕЗД".
    """
    # 1) Ищем элемент, который содержит ключевую фразу.
    try:
        el = driver.find_element(By.XPATH, "//*[contains(normalize-space(.), 'Номер первой категории')]")
        t = (el.text or "").strip()
        if t:
            # Варианты кавычек: "..." или «...»
            m = re.search(r'Номер первой категории\s*[\"“”](.*?)[\"“”]', t)
            if m:
                return m.group(1).strip()
            m2 = re.search(r'Номер первой категории.*?[«](.*?)[»]', t)
            if m2:
                return m2.group(1).strip()
            # На случай одинарных кавычек
            m3 = re.search(r"Номер первой категории\s*['‘](.*?)[’']", t)
            if m3:
                return m3.group(1).strip()
    except NoSuchElementException:
        pass

    # 2) fallback: заголовки, но фильтруем мусорные 'ЗАЕЗД/ВЫЕЗД/Количество'
    for sel in ["h1", "h2", "h3", "h4"]:
        els = driver.find_elements(By.CSS_SELECTOR, sel)
        for e in els:
            txt = (e.text or "").strip()
            if not txt:
                continue
            txt_norm = " ".join(txt.split())
            txt_upper = txt_norm.upper()
            if txt_upper in EXCLUDED_TITLE_TOKENS:
                continue
            if "ЗАЕЗД" in txt_upper or "ВЫЕЗД" in txt_upper or "КОЛИЧЕСТВО" in txt_upper:
                continue
            # Категории обычно длиннее, чем служебные подписи.
            if len(txt_norm) >= 3:
                return txt_norm

    return None


def parse_base_price_for_guest(body_text: str, preferred_guest: int = 1) -> Optional[int]:
    """
    На страницах категорий цены часто заданы словами:
    - 1 гость - 5400 рублей, 2 гостя - 5400 рублей...
    - 1-2 гостя- 5500 рублей, 3 гостя - 6500 рублей
    - 5 гостей – 10 600 рублей

    В fallback-режиме нам нужна "базовая" цена по категории.
    Сначала пытаемся выбрать цену, которая относится к preferred_guest (по умолчанию 1).
    Если точного совпадения нет — берем первую подходящую сумму "руб".
    """
    text = " ".join(body_text.split())

    # Диапазон: "1-2 гостя ... 5500 руб"
    range_pat = re.compile(
        r"(\d+)\s*[-–—]\s*(\d+)\s*гост[а-я]*[^0-9]*([\d][\d\s]*)\s*руб",
        flags=re.IGNORECASE,
    )
    for a_str, b_str, price_str in range_pat.findall(text):
        a = int(a_str)
        b = int(b_str)
        if a <= preferred_guest <= b:
            return int(price_str.replace(" ", ""))

    # Точное значение: "1 гость ... 5400 руб"
    exact_pat = re.compile(
        r"(\d+)\s*гост[а-я]*[^0-9]*([\d][\d\s]*)\s*руб",
        flags=re.IGNORECASE,
    )
    for guest_str, price_str in exact_pat.findall(text):
        if int(guest_str) == preferred_guest:
            return int(price_str.replace(" ", ""))

    # Если не нашли для preferred_guest, пробуем взять "N гостей ... цена"
    # (частая ситуация: для категории задана цена только для 5 гостей).
    guest_any_pat = re.compile(
        r"(\d+)\s*гост[а-я]*[^0-9]*([\d][\d\s]*)\s*руб",
        flags=re.IGNORECASE,
    )
    for _guest_str, price_str in guest_any_pat.findall(text):
        return int(price_str.replace(" ", ""))

    # Если не нашли для preferred_guest — берем первую сумму рядом с рублями.
    any_pat = re.compile(r"([\d][\d\s]*)\s*руб", flags=re.IGNORECASE)
    m = any_pat.search(text)
    if m:
        return int(m.group(1).replace(" ", ""))

    return None


def extract_categories_from_category_pages(driver: webdriver.Chrome, preferred_guest: int = 1) -> list:
    """
    Извлекаем категории и их базовые цены именно с указанных страниц категорий.
    """
    categories = []
    wait = WebDriverWait(driver, 20)

    for page_url in CATEGORY_PAGES:
        driver.get(page_url)
        human_sleep(1.0, 2.0)

        # Подождем появления заголовка/контента страницы.
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

        body_text = driver.find_element(By.TAG_NAME, "body").text

        # Сначала надежно пытаемся извлечь название из DOM по ключевой фразе.
        title = get_category_title_from_dom(driver)
        if not title:
            # fallback: по тексту страницы (старый метод)
            title = extract_category_title(body_text)

        base_price = parse_base_price_for_guest(body_text, preferred_guest=preferred_guest)

        if not title or base_price is None:
            print(f"ВНИМАНИЕ: не удалось извлечь базовую цену/название для {page_url}")
            continue

        categories.append(
            {
                "category_name": title,
                "base_price": base_price,
                "category_url": page_url,
            }
        )

    return categories


def set_picker_date(
    driver: webdriver.Chrome,
    date_in: date,
    date_out: date,
    guests: int = 1,
) -> None:
    """
    У виджета бронирования на страницах категорий присутствуют:
    - input[name="date-in"]
    - input[name="date-out"]
    - b#zaezddatein, i#zaezdmonthin, i#zaezdyearin
    - b#viezddatein, i#viezdmonthin, i#viezdyearin

    Мы выставляем и текстовые элементы, и значение инпутов.
    """
    in_text = format_for_picker(date_in)
    out_text = format_for_picker(date_out)

    # Инпуты даты
    date_in_input = driver.find_element(By.CSS_SELECTOR, 'input[name="date-in"]')
    date_out_input = driver.find_element(By.CSS_SELECTOR, 'input[name="date-out"]')

    # Внутренние элементы, которые отображаются в календаре (в HTML они присутствуют всегда)
    # Форматируем день/месяц как 2-значные, чтобы совпадало с версткой.
    in_day = f"{date_in.day:02d}"
    in_month = f"{date_in.month:02d}"
    in_year = f"{date_in.year:04d}"

    out_day = f"{date_out.day:02d}"
    out_month = f"{date_out.month:02d}"
    out_year = f"{date_out.year:04d}"

    # Устанавливаем значения через JS (так обычно стабильнее, чем clear()+send_keys в кастомных виджетах)
    driver.execute_script(
        """
        const input = arguments[0];
        const val = arguments[1];
        input.value = val;
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        """,
        date_in_input,
        in_text,
    )
    driver.execute_script(
        """
        const input = arguments[0];
        const val = arguments[1];
        input.value = val;
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        """,
        date_out_input,
        out_text,
    )

    # Экранные значения (отображаемые части виджета)
    for el_id, value in [
        ("zaezddatein", in_day),
        ("zaezdmonthin", in_month),
        ("zaezdyearin", in_year),
        ("viezddatein", out_day),
        ("viezdmonthin", out_month),
        ("viezdyearin", out_year),
    ]:
        try:
            el = driver.find_element(By.ID, el_id)
            driver.execute_script("arguments[0].textContent = arguments[1];", el, value)
        except NoSuchElementException:
            # Если каких-то элементов нет (редко, но бывает) — не падаем, просто пропускаем.
            pass

    # Гости (виджет хранит default и считает цену в зависимости от количества)
    try:
        guests_input = driver.find_element(By.ID, "kolichestvogosteiin")
        driver.execute_script(
            "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change', { bubbles: true }));",
            guests_input,
            f"{guests:02d}",
        )
    except NoSuchElementException:
        pass


def try_extract_exact_price_from_booking_page(driver: webdriver.Chrome) -> Optional[int]:
    """
    Пытаемся извлечь точную цену на странице бронирования внешнего сервиса.

    Так как точная разметка внешнего сервиса может меняться, здесь используется
    безопасный эвристический поиск: на странице ищем первое вхождение "руб/₽"
    и парсим первое число рядом с валютой.
    """
    human_sleep(1.0, 2.5)

    try:
        body_text = driver.find_element(By.TAG_NAME, "body").text
    except NoSuchElementException:
        return None

    # Ищем подстроку вида "xxxx руб" или "34,000.00 RUB"
    m = re.search(r"(\d[\d\s,]*)([.,]\d+)?\s*(руб|₽|рублей|RUB)", body_text, flags=re.IGNORECASE)
    if not m:
        return None

    int_part = m.group(1).replace(" ", "").replace(",", "")
    dec_part = m.group(2) or ""
    if dec_part:
        val = float(int_part + dec_part.replace(",", "."))
        return int(round(val))
    return int(int_part)


def open_booking_new_tab_and_parse_price(
    driver: webdriver.Chrome,
    category_url: str,
    date_in: date,
    date_out: date,
    guests: int = 1,
) -> Optional[int]:
    """
    Открываем страницу категории, выставляем даты/гостей, кликаем на "check-avail"
    (ссылка на bookonline24), парсим цену и возвращаемся назад.
    """
    driver.get(category_url)
    human_sleep(1.0, 2.0)

    wait = WebDriverWait(driver, 20)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="date-in"]')))

    set_picker_date(driver, date_in=date_in, date_out=date_out, guests=guests)
    human_sleep(0.8, 1.5)

    # Клик по ссылке проверки наличия / онлайн-бронирование.
    # По HTML-контексту виджета: div.avail-sec -> li.check-avail -> a.lnk-default
    check_link = driver.find_element(By.CSS_SELECTOR, "div.avail-sec li.check-avail a.lnk-default")

    before_handles = set(driver.window_handles)
    check_link.click()
    human_sleep(1.2, 2.0)

    # Если открылось в новой вкладке — используем ее, иначе парсим текущую страницу.
    after_handles = set(driver.window_handles)
    new_handles = list(after_handles - before_handles)
    if new_handles:
        driver.switch_to.window(new_handles[0])
        try:
            return try_extract_exact_price_from_booking_page(driver)
        finally:
            driver.close()
            driver.switch_to.window(list(before_handles)[0])
            human_sleep(0.6, 1.2)
    else:
        # Та же вкладка: парсим текущую страницу.
        return try_extract_exact_price_from_booking_page(driver)


def main() -> None:
    start = parse_ymd(START_DATE)
    end = parse_ymd(END_DATE)
    if end < start:
        raise ValueError("END_DATE не может быть меньше START_DATE")

    dates = []
    cur = start
    while cur <= end:
        dates.append(cur)
        cur += timedelta(days=1)

    # =========================
    # Запуск Chrome в видимом режиме
    # =========================
    chrome_service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=chrome_service)
    driver.maximize_window()

    try:
        # Достаем категории и базовые цены с конкретных страниц категорий (как вы просили).
        categories = extract_categories_from_category_pages(driver, preferred_guest=1)
        if not categories:
            raise RuntimeError("Не удалось извлечь категории и базовые цены с страниц категорий.")

        rows = []

        # Основной источник теперь — единая страница бронирования, где есть все категории.
        # Для каждой даты (1 ночь) собираем список категорий/цен и пишем в таблицу.
        for d in dates:
            date_in = d
            date_out = d + timedelta(days=1)  # 1 ночь

            url = booking_url(date_in, date_out, adults_count=ADULTS_COUNT)
            print(f"\nОткрываю бронирование: {url}")
            driver.get(url)
            wait_booking_page_ready(driver, timeout_s=35)

            # Цены появляются только после клика «Проверить наличие»
            if not click_check_availability_and_wait_prices(driver, timeout_s=25):
                print("Кнопка «Проверить наличие» не найдена, пробуем парсить без клика.")

            scraped = scrape_prices_from_booking_page(driver)

            if scraped:
                for item in scraped:
                    rows.append(
                        {
                            "Дата": date_in.strftime("%Y-%m-%d"),
                            "Отель": "Отель Меридиан",
                            "Категория номера": item["category_name"],
                            "Стоимость (руб)": item["price"],
                            "Примечание": "",
                        }
                    )
            else:
                # Если не смогли вытащить цены (например, техработы/блокировка) — fallback:
                print(
                    "ВНИМАНИЕ: не удалось извлечь цены со страницы бронирования. "
                    "Будут использованы базовые цены со страниц категорий отеля."
                )
                for cat in categories:
                    rows.append(
                        {
                            "Дата": date_in.strftime("%Y-%m-%d"),
                            "Отель": "Отель Меридиан",
                            "Категория номера": cat["category_name"],
                            "Стоимость (руб)": cat["base_price"],
                            "Примечание": "* (цена базовая, требуется уточнение). Уточнить: "
                            + PHONE_1
                            + "; "
                            + PHONE_2,
                        }
                    )

            human_sleep(1.0, 2.0)

        # Дополнительные отели из HOTEL_SOURCES
        for hotel in HOTEL_SOURCES:
            hotel_name = hotel["name"]
            date_in = None
            date_out = None
            for d in dates:
                date_in = d
                date_out = d + timedelta(days=1)
                if hotel["type"] == "multi_page":
                    for page_url in hotel["urls"]:
                        print(f"\n[{hotel_name}] {date_in} -> {page_url}")
                        driver.get(page_url)
                        human_sleep(1.2, 2.5)
                        one = scrape_one_category_per_page(driver)
                        if one:
                            rows.append(
                                {
                                    "Дата": date_in.strftime("%Y-%m-%d"),
                                    "Отель": hotel_name,
                                    "Категория номера": one["category_name"],
                                    "Стоимость (руб)": one["price"],
                                    "Примечание": "",
                                }
                            )
                        human_sleep(0.8, 1.5)
                elif hotel["type"] == "date_button":
                    dfrom = format_date_dd_mm_yyyy(date_in)
                    dto = format_date_dd_mm_yyyy(date_out)
                    url = hotel["url_template"].format(dfrom=dfrom, dto=dto)
                    print(f"\n[{hotel_name}] {date_in} -> {url[:80]}...")
                    driver.get(url)
                    human_sleep(1.5, 2.5)
                    if hotel.get("cookie_button"):
                        if accept_cookie_if_present(driver, hotel["cookie_button"], wait_after_s=1.0):
                            human_sleep(0.8, 1.5)
                    if hotel.get("button"):
                        if hotel.get("button_in_iframe"):
                            if not click_button_in_iframe_then_page(
                                driver,
                                hotel["button"],
                                wait_after_s=3.0,
                                button_alt=hotel.get("button_alt"),
                                iframe_wait_seconds=hotel.get("iframe_wait_seconds", 5),
                            ):
                                print(f"  Кнопка «{hotel['button']}» не найдена (в т.ч. в iframe).")
                        else:
                            if not click_button_by_text(driver, hotel["button"], wait_after_s=3.0):
                                for alt in hotel.get("button_alt") or []:
                                    if click_button_by_text(driver, alt, wait_after_s=3.0):
                                        break
                                else:
                                    print(f"  Кнопка «{hotel['button']}» не найдена.")
                    scraped = scrape_prices_generic(driver)
                    for item in scraped:
                        rows.append(
                            {
                                "Дата": date_in.strftime("%Y-%m-%d"),
                                "Отель": hotel_name,
                                "Категория номера": item["category_name"],
                                "Стоимость (руб)": item["price"],
                                "Примечание": "",
                            }
                        )
                    human_sleep(1.0, 2.0)
                elif hotel["type"] == "button":
                    print(f"\n[{hotel_name}] {date_in} -> {hotel['url']}")
                    driver.get(hotel["url"])
                    human_sleep(1.5, 2.5)
                    if hotel.get("cookie_button"):
                        if accept_cookie_if_present(driver, hotel["cookie_button"], wait_after_s=1.0):
                            human_sleep(0.8, 1.5)
                    clicked = click_button_by_text(driver, hotel["button"], wait_after_s=3.0)
                    if not clicked:
                        for alt in hotel.get("button_alt") or []:
                            if click_button_by_text(driver, alt, wait_after_s=3.0):
                                clicked = True
                                break
                    if not clicked:
                        # виджет бронирования может быть в iframe (например Freezone)
                        clicked = click_button_in_iframe_then_page(
                            driver,
                            hotel["button"],
                            wait_after_s=3.0,
                            button_alt=hotel.get("button_alt"),
                            iframe_wait_seconds=4,
                        )
                    if not clicked:
                        print(f"  Кнопка «{hotel['button']}» не найдена.")
                    scraped = scrape_prices_generic(driver)
                    for item in scraped:
                        rows.append(
                            {
                                "Дата": date_in.strftime("%Y-%m-%d"),
                                "Отель": hotel_name,
                                "Категория номера": item["category_name"],
                                "Стоимость (руб)": item["price"],
                                "Примечание": "",
                            }
                        )
                    human_sleep(1.0, 2.0)
                else:
                    # direct (в т.ч. Постоялый двор Русь — даём время на загрузку SPA)
                    print(f"\n[{hotel_name}] {date_in} -> {hotel['url']}")
                    driver.get(hotel["url"])
                    extra_wait = hotel.get("wait_seconds") or 1.5
                    human_sleep(extra_wait, extra_wait + 1.5)
                    # Ждём появления цен на странице (для otelms и других SPA)
                    try:
                        WebDriverWait(driver, 15).until(
                            lambda d: "руб" in ((d.find_element(By.TAG_NAME, "body").text) or "").lower()
                            or "₽" in (d.find_element(By.TAG_NAME, "body").text or "")
                            or "RUB" in (d.find_element(By.TAG_NAME, "body").text or "")
                        )
                    except (TimeoutException, NoSuchElementException):
                        pass
                    human_sleep(1.0, 2.0)
                    scraped = scrape_prices_generic(driver)
                    for item in scraped:
                        rows.append(
                            {
                                "Дата": date_in.strftime("%Y-%m-%d"),
                                "Отель": hotel_name,
                                "Категория номера": item["category_name"],
                                "Стоимость (руб)": item["price"],
                                "Примечание": "",
                            }
                        )
                    human_sleep(1.0, 2.0)

        df = pd.DataFrame(
            rows,
            columns=["Дата", "Отель", "Категория номера", "Стоимость (руб)", "Примечание"],
        )
        out_name = f"prices_all_{START_DATE}_{END_DATE}.xlsx"
        out_path = out_name  # сохраняем в текущую папку запуска/репозиторий Cursor
        # pandas для .xlsx требует openpyxl (если не указан другой engine)
        try:
            import openpyxl  # noqa: F401
        except ImportError:
            raise SystemExit("Ошибка: не установлен модуль `openpyxl`. Установите: pip install openpyxl")

        df.to_excel(out_path, index=False)
        print(f"\nГотово. Excel сохранен: {out_path}")
        print(df.head(10))

    finally:
        driver.quit()


if __name__ == "__main__":
    main()

