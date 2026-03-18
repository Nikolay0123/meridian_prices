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


def human_sleep(a: float = 1.0, b: float = 2.2) -> None:
    """Случайная задержка для уменьшения нагрузки и для стабильности."""
    time.sleep(random.uniform(a, b))


def parse_ymd(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()


def format_for_picker(d: date) -> str:
    """Чаще всего виджеты на сайте принимают формат ДД.ММ.ГГГГ."""
    return d.strftime("%d.%m.%Y")


def to_int_rub(text: str) -> Optional[int]:
    """
    Пытаемся извлечь число в рублях.
    Примеры входа: "от 14 000 рублей", "1 800,00 рублей".
    """
    if not text:
        return None

    # Схватим первое число и (опционально) дробную часть.
    # Учитываем пробелы в тысячах и запятую в дробной части.
    m = re.search(r"(\d[\d\s]*)([.,]\d+)?\s*(?:руб|₽|рублей)", text, flags=re.IGNORECASE)
    if not m:
        # Если "руб" рядом нет, попробуем просто числа.
        m2 = re.search(r"(\d[\d\s]*)([.,]\d+)?", text)
        if not m2:
            return None
        m = m2

    int_part = m.group(1).replace(" ", "")
    dec_part = m.group(2) or ""

    if dec_part:
        # "1800,00" -> float
        val = float(int_part + dec_part.replace(",", "."))
        return int(round(val))
    return int(int_part)


def extract_category_title(body_text: str) -> Optional[str]:
    """
    На страницах категорий встречается фраза:
    Номер первой категории "Стандарт Классический"
    """
    m = re.search(r'Номер первой категории\s*[\"“](.*?)[\"”]', body_text)
    if m:
        return m.group(1).strip()
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
        title = extract_category_title(body_text)
        if not title:
            # fallback: по заголовкам
            for sel in ["h1", "h2", "h3", "h4"]:
                try:
                    el = driver.find_element(By.CSS_SELECTOR, sel)
                    t = el.text.strip()
                    if t:
                        title = t
                        break
                except NoSuchElementException:
                    pass

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

    # Ищем подстроку вида "xxxx руб"
    m = re.search(r"(\d[\d\s]*)([.,]\d+)?\s*(руб|₽|рублей)", body_text, flags=re.IGNORECASE)
    if not m:
        return None

    int_part = m.group(1).replace(" ", "")
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

        # Если точные цены не извлекаются для категории, мы пометим это в примечаниях.
        exact_supported = {}

        for cat in categories:
            cat_name = cat["category_name"]
            cat_url = cat["category_url"]
            base_price = cat["base_price"]
            if not base_price:
                continue

            print(f"Категория: {cat_name} | базовая цена: {base_price} руб")

            for d in dates:
                date_in = d
                date_out = d + timedelta(days=1)  # 1 ночь

                exact_price = None
                if exact_supported.get(cat_name, True):
                    try:
                        exact_price = open_booking_new_tab_and_parse_price(
                            driver=driver,
                            category_url=cat_url,
                            date_in=date_in,
                            date_out=date_out,
                            guests=1,
                        )
                    except (TimeoutException, NoSuchElementException, Exception) as e:
                        # Любая нештатная ситуация считаем "точные цены не достались"
                        # и падаем в fallback.
                        exact_price = None

                note = ""
                cost = base_price

                if exact_price is not None:
                    cost = exact_price
                else:
                    # В fallback-режиме добавляем ТЗ-пометку и телефоны.
                    note = "* (цена базовая, требуется уточнение). Уточнить: " + PHONE_1 + "; " + PHONE_2
                    exact_supported[cat_name] = False
                    print(
                        f"ВНИМАНИЕ: точные цены не удалось извлечь для категории '{cat_name}' на {date_in}. "
                        f"Подставлена базовая цена {base_price} руб."
                    )
                    # Даем немного времени, чтобы пользователь в видимом Chrome мог заметить что происходит.

                rows.append(
                    {
                        "Дата": date_in.strftime("%Y-%m-%d"),
                        "Категория номера": cat_name,
                        "Стоимость (руб)": cost,
                        "Примечание": note,
                    }
                )

                human_sleep(0.8, 1.6)

        df = pd.DataFrame(rows, columns=["Дата", "Категория номера", "Стоимость (руб)", "Примечание"])
        out_name = f"prices_meridian_{START_DATE}_{END_DATE}.xlsx"
        out_path = out_name  # сохраняем в текущую папку запуска/репозиторий Cursor

        df.to_excel(out_path, index=False)
        print(f"\nГотово. Excel сохранен: {out_path}")
        print(df.head(10))

    finally:
        driver.quit()


if __name__ == "__main__":
    main()

