from collections import Counter, defaultdict
from datetime import datetime
from io import BytesIO
from pathlib import Path, PurePosixPath
import math
import re
import unicodedata
from zipfile import ZIP_DEFLATED, ZipFile

import openpyxl
import pandas as pd
try:
    from pandastable import Table, TableModel
except ImportError:
    Table = TableModel = None

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


UNIT_COLUMNS = [
    "SKU",
    "Артикул",
    "Текущая цена",
    "Заказано товаров, шт",
    "Доставлено товаров, шт",
    "Выручка",
    "Баллы за скидки",
    "Программы партнёров",
    "Вознаграждение Ozon",
    "Эквайринг",
    "Обработка отправления",
    "Логистика",
    "Доставка до места выдачи",
    "Стоимость размещения",
    "Обработка возврата",
    "Обратная логистика",
    "Утилизация",
    "Дополнительная обработка ОВХ",
    "Операционные ошибки",
    "Оплата за клик",
    "Оплата за заказ",
    "Звёздные товары",
    "Платный бренд",
    "Прибыль за период",
]

UNIT_ALIASES = {
    "SKU": ("sku", "номер sku"),
    "Артикул": ("артикул", "номер артикула"),
    "Текущая цена": ("текущая цена",),
    "Заказано товаров, шт": ("заказано товаров шт",),
    "Доставлено товаров, шт": ("доставлено товаров шт",),
    "Выручка": ("выручка",),
    "Баллы за скидки": ("баллы за скидки",),
    "Программы партнёров": ("программы партнеров",),
    "Вознаграждение Ozon": ("вознаграждение ozon",),
    "Эквайринг": ("эквайринг",),
    "Обработка отправления": ("обработка отправления",),
    "Логистика": ("логистика",),
    "Доставка до места выдачи": ("доставка до места выдачи",),
    "Стоимость размещения": ("стоимость размещения",),
    "Обработка возврата": ("обработка возврата",),
    "Обратная логистика": ("обратная логистика", "обработка логистика"),
    "Утилизация": ("утилизация",),
    "Дополнительная обработка ОВХ": ("дополнительная обработка овх",),
    "Операционные ошибки": ("операционные ошибки", "операционная ошибка"),
    "Оплата за клик": ("оплата за клик",),
    "Оплата за заказ": ("оплата за заказ",),
    "Звёздные товары": ("звездные товары",),
    "Платный бренд": ("платный бренд",),
    "Прибыль за период": ("прибыль за период",),
}

STOCK_COLUMNS = [
    "Кластер",
    "Склад",
    "Доступно к продаже",
    "Брак, доступный к вывозу с поставки",
]

STOCK_ALIASES = {
    "Артикул": ("артикул", "номер артикула"),
    "SKU": ("sku", "номер sku"),
    "Кластер": ("кластер",),
    "Склад": ("склад",),
    "Доступно к продаже": ("доступно к продаже",),
    "Брак, доступный к вывозу с поставки": (
        "брак доступный к вывозу с поставки",
    ),
}

REDEMPTION_ALIASES = {
    "Артикул": ("артикул",),
    "Уровень выкупа товара": (
        "уровень выкупа товара за последние 50 заказов",
        "уровень выкупа товара",
    ),
}

FINAL_COLUMNS = UNIT_COLUMNS.copy()
FINAL_COLUMNS.insert(FINAL_COLUMNS.index("Выручка") + 1, "Реальная цена")
FINAL_COLUMNS.extend(STOCK_COLUMNS)
FINAL_COLUMNS.append("Уровень выкупа товара")

REPORT_TITLES = {
    "unit": "Юнит-экономика",
    "stock": "Управление остатками",
    "redemption": "Уровень выкупа",
}

WB_BASE_COLUMNS = [
    "Бренд",
    "Предмет",
    "Сезон",
    "Коллекция",
    "Наименование",
    "Артикул продавца",
    "Артикул WB",
    "Баркод",
    "Размер",
    "Контракт",
    "Склад",
]

WB_BASE_ALIASES = {
    "Бренд": ("бренд",),
    "Предмет": ("предмет",),
    "Сезон": ("сезон",),
    "Коллекция": ("коллекция",),
    "Наименование": ("наименование", "название"),
    "Артикул продавца": ("артикул продавца",),
    "Артикул WB": ("артикул wb",),
    "Баркод": ("баркод",),
    "Размер": ("размер",),
    "Контракт": ("контракт",),
    "Склад": ("склад",),
}

WB_SUPPLIER_METRICS = [
    "Заказано, шт.",
    "Сумма заказов минус комиссия WB, руб.",
    "Выкупили, шт.",
    "К перечислению за товар, руб.",
    "Текущий остаток, шт.",
]

WB_SUPPLIER_ALIASES = {
    "Заказано, шт.": ("шт",),
    "Сумма заказов минус комиссия WB, руб.": (
        "сумма заказов минус комиссия wb руб",
    ),
    "Выкупили, шт.": ("выкупили шт",),
    "К перечислению за товар, руб.": (
        "к перечислению за товар руб",
    ),
    "Текущий остаток, шт.": ("текущий остаток шт",),
}

WB_FINANCE_METRICS = [
    "Итог, ₽",
    "Продажи, ₽",
    "Продажи, шт",
    "Возвраты, ₽",
    "Возвраты, шт",
    "Логистика, ₽",
    "Штрафы, ₽",
    "Комиссия WB, ₽",
    "Эквайринг, ₽",
    "Потери, подмены и товары с дефектами, ₽",
    "Доплаты, ₽",
    "Программа лояльности, ₽",
]

WB_REPORT_TITLES = {
    "supplier": "Товары по складам",
    "finance": "Доходы и расходы",
}

WB_PRODUCT_KEY_FIELDS = (
    "Артикул WB",
    "Артикул продавца",
    "Баркод",
)

WB_RESULT_COLUMNS = [
    "Наименование",
    "Артикул продавца",
    "Артикул WB",
    "Баркод",
    "Заказано, шт",
    "Выкуплено, шт",
    "Невыкуплено, шт",
    "Выкуп, %",
    "Заказы, ₽",
    "К перечислению, ₽",
    "Средний заказ, ₽",
    "Выплата за выкуп, ₽",
    "Остаток, шт",
    "Продажи, шт",
    "Продажи, ₽",
    "Средняя продажа, ₽",
    "Возвраты, шт",
    "Возвраты, ₽",
    "Логистика, ₽",
    "Комиссия WB, ₽",
    "Эквайринг, ₽",
    "Штрафы и потери, ₽",
    "Расходы WB, ₽",
    "Расходы на продажу, ₽",
    "Расходы WB, %",
    "Результат после WB, ₽",
    "Остаётся после WB, %",
    "Продаж в день",
    "Запас, дней",
    "Продажи ранее, ₽",
    "Динамика продаж, %",
    "Результат ранее, ₽",
    "Динамика результата, %",
    "Заказано ранее, шт",
    "Динамика заказов, %",
    "Выкуплено ранее, шт",
    "Динамика выкупов, %",
]

MAX_ZIP_EXCEL_SIZE = 100 * 1024 * 1024


def normalize_text(value):
    if value is None:
        return ""

    text = unicodedata.normalize("NFKC", str(value))
    text = text.replace("\xa0", " ").replace("_x000D_", " ")
    text = text.replace("ё", "е").replace("Ё", "Е")
    return " ".join(text.split()).strip().lower()


def normalize_header(value):
    text = normalize_text(value)
    return re.sub(r"[^0-9a-zа-я]+", " ", text).strip()


def display_identifier(value):
    if value is None:
        return ""

    if isinstance(value, bool):
        return str(value)

    if isinstance(value, int):
        return str(value)

    if isinstance(value, float):
        if math.isnan(value):
            return ""
        if value.is_integer():
            return str(int(value))

    return unicodedata.normalize("NFKC", str(value)).strip()


def identifier_keys(value):
    value = display_identifier(value)
    if not value:
        return []

    parts = re.split(r"[,;\n]+", value)
    keys = []

    for part in parts:
        key = normalize_text(part)
        if key and key not in keys:
            keys.append(key)

    full_key = normalize_text(value)
    if full_key and full_key not in keys:
        keys.append(full_key)

    return keys


def clean_value(value):
    if value is None:
        return None

    if isinstance(value, float):
        if math.isnan(value):
            return None
        if value.is_integer():
            return int(value)
        return round(value, 2)

    return value


def number_value(value):
    if value is None or value == "":
        return None

    if isinstance(value, bool):
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace("\xa0", " ").replace(" ", "")
    text = text.replace(",", ".")

    try:
        return float(text)
    except ValueError:
        return None


def calculate_real_price(record):
    revenue = number_value(record.get("Выручка"))
    delivered = number_value(record.get("Доставлено товаров, шт"))

    if revenue is None or delivered in (None, 0):
        return None

    result = round(revenue / delivered, 2)
    if result.is_integer():
        return int(result)
    return result


def format_cell_value(cell):
    value = cell.value

    if (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and "%" in str(cell.number_format)
    ):
        percentage = round(float(value) * 100, 2)
        if percentage.is_integer():
            return f"{int(percentage)}%"
        return f"{percentage:g}%"

    return clean_value(value)


def _repair_alignment_attributes(xml_text):
    valid_horizontal = {
        "general",
        "left",
        "center",
        "right",
        "fill",
        "justify",
        "centercontinuous",
        "distributed",
    }
    valid_vertical = {"top", "center", "bottom", "justify", "distributed"}

    def replace_horizontal(match):
        value = match.group(1)
        lowered = value.lower()
        if lowered not in valid_horizontal:
            lowered = "general"
        if lowered == "centercontinuous":
            lowered = "centerContinuous"
        return f'horizontal="{lowered}"'

    def replace_vertical(match):
        value = match.group(1).lower()
        if value not in valid_vertical:
            value = "bottom"
        return f'vertical="{value}"'

    xml_text = re.sub(
        r'horizontal="([^"]+)"',
        replace_horizontal,
        xml_text,
    )
    xml_text = re.sub(
        r'vertical="([^"]+)"',
        replace_vertical,
        xml_text,
    )
    return xml_text


def rewind_source(file_source):
    if hasattr(file_source, "seek"):
        file_source.seek(0)


def load_workbook_safely(file_source):
    if isinstance(file_source, (str, Path)):
        suffix=Path(file_source).suffix.lower()
        if suffix=='.xls':
            import xlrd
            old=xlrd.open_workbook(str(file_source)); new=openpyxl.Workbook();new.remove(new.active)
            try:
                for sheet in old.sheets():
                    ws=new.create_sheet(sheet.name)
                    for row in sheet.get_rows():
                        ws.append([xlrd.xldate_as_datetime(c.value,old.datemode) if c.ctype==xlrd.XL_CELL_DATE else c.value for c in row])
            finally: old.release_resources()
            return new
        if suffix in ('.csv','.tsv'):
            import csv
            payload=Path(file_source).read_bytes()
            for encoding in ('utf-8-sig','utf-16','cp1251'):
                try: text=payload.decode(encoding);break
                except UnicodeError:continue
            try: dialect=csv.Sniffer().sniff(text[:10000],delimiters=';,\t')
            except csv.Error: dialect=csv.excel_tab if suffix=='.tsv' else csv.excel
            new=openpyxl.Workbook()
            for row in csv.reader(text.splitlines(),dialect):new.active.append(row)
            return new
    rewind_source(file_source)
    try:
        return openpyxl.load_workbook(
            file_source,
            data_only=True,
            read_only=False,
        )
    except ValueError as error:
        if "stylesheet" not in str(error).lower():
            raise

    repaired_file = BytesIO()
    rewind_source(file_source)

    with ZipFile(file_source, "r") as source_zip:
        with ZipFile(repaired_file, "w", ZIP_DEFLATED) as target_zip:
            for member in source_zip.infolist():
                content = source_zip.read(member.filename)

                if member.filename == "xl/styles.xml":
                    xml_text = content.decode("utf-8")
                    xml_text = _repair_alignment_attributes(xml_text)
                    content = xml_text.encode("utf-8")

                target_zip.writestr(member, content)

    repaired_file.seek(0)
    workbook = openpyxl.load_workbook(
        repaired_file,
        data_only=True,
        read_only=False,
    )
    workbook._repaired_source = repaired_file
    return workbook


def close_workbook(workbook):
    repaired_source = getattr(workbook, "_repaired_source", None)
    workbook.close()
    if repaired_source is not None:
        repaired_source.close()


def row_header_keys(worksheet, row_number):
    return {
        normalize_header(worksheet.cell(row_number, column).value)
        for column in range(1, worksheet.max_column + 1)
        if worksheet.cell(row_number, column).value is not None
    }


def contains_exact(keys, *aliases):
    normalized_aliases = {normalize_header(alias) for alias in aliases}
    return bool(keys & normalized_aliases)


def detect_report(workbook):
    max_scan_rows = 30

    for worksheet in workbook.worksheets:
        for row_number in range(
            1,
            min(worksheet.max_row, max_scan_rows) + 1,
        ):
            keys = row_header_keys(worksheet, row_number)

            if (
                contains_exact(keys, "SKU")
                and contains_exact(keys, "Артикул")
                and contains_exact(keys, "Текущая цена")
                and contains_exact(keys, "Выручка")
                and contains_exact(keys, "Прибыль за период")
            ):
                return "unit", worksheet, row_number

    for worksheet in workbook.worksheets:
        for row_number in range(
            1,
            min(worksheet.max_row, max_scan_rows) + 1,
        ):
            keys = row_header_keys(worksheet, row_number)

            if (
                contains_exact(keys, "Артикул")
                and contains_exact(
                    keys,
                    "Уровень выкупа товара за последние 50 заказов",
                    "Уровень выкупа товара",
                )
                and contains_exact(keys, "Заказано товаров")
            ):
                return "redemption", worksheet, row_number

    for worksheet in workbook.worksheets:
        for row_number in range(
            1,
            min(worksheet.max_row, max_scan_rows) + 1,
        ):
            keys = row_header_keys(worksheet, row_number)

            if (
                contains_exact(keys, "SKU")
                and contains_exact(keys, "Артикул")
                and contains_exact(keys, "Кластер")
                and contains_exact(keys, "Склад")
            ):
                return "stock", worksheet, row_number

    return None


def find_column(worksheet, rows, aliases):
    alias_keys = {normalize_header(alias) for alias in aliases}

    for column in range(1, worksheet.max_column + 1):
        for row_number in rows:
            key = normalize_header(
                worksheet.cell(row_number, column).value
            )
            if key in alias_keys:
                return column

    return None


def valid_identifier(value):
    key = normalize_header(display_identifier(value))
    invalid_values = {
        "",
        "8",
        "артикул",
        "sku",
        "номер артикула",
        "номер sku",
        "нередактируемое",
    }
    return key not in invalid_values


def extract_unit_rows(worksheet, header_row):
    columns = {
        output_name: find_column(
            worksheet,
            (header_row,),
            aliases,
        )
        for output_name, aliases in UNIT_ALIASES.items()
    }

    if columns["SKU"] is None and columns["Артикул"] is None:
        return []

    rows = []

    for row_number in range(header_row + 1, worksheet.max_row + 1):
        sku = (
            worksheet.cell(row_number, columns["SKU"]).value
            if columns["SKU"]
            else None
        )
        article = (
            worksheet.cell(row_number, columns["Артикул"]).value
            if columns["Артикул"]
            else None
        )

        if not valid_identifier(sku) and not valid_identifier(article):
            continue

        record = {}

        for output_name in UNIT_COLUMNS:
            column = columns.get(output_name)
            value = (
                worksheet.cell(row_number, column).value
                if column is not None
                else None
            )

            if output_name in ("SKU", "Артикул"):
                value = display_identifier(value)
            else:
                value = clean_value(value)

            record[output_name] = value

        rows.append(record)

    return rows


def extract_stock_rows(worksheet, header_row):
    header_rows = range(
        header_row,
        min(header_row + 4, worksheet.max_row + 1),
    )
    columns = {
        output_name: find_column(
            worksheet,
            header_rows,
            aliases,
        )
        for output_name, aliases in STOCK_ALIASES.items()
    }

    if columns["SKU"] is None and columns["Артикул"] is None:
        return []

    rows = []

    for row_number in range(header_row + 1, worksheet.max_row + 1):
        sku = (
            worksheet.cell(row_number, columns["SKU"]).value
            if columns["SKU"]
            else None
        )
        article = (
            worksheet.cell(row_number, columns["Артикул"]).value
            if columns["Артикул"]
            else None
        )

        if not valid_identifier(sku) and not valid_identifier(article):
            continue

        record = {
            "SKU": display_identifier(sku),
            "Артикул": display_identifier(article),
        }

        for output_name in STOCK_COLUMNS:
            column = columns.get(output_name)
            value = (
                worksheet.cell(row_number, column).value
                if column is not None
                else None
            )
            record[output_name] = clean_value(value)

        rows.append(record)

    return rows


def extract_redemption_rows(worksheet, header_row):
    columns = {
        output_name: find_column(
            worksheet,
            (header_row,),
            aliases,
        )
        for output_name, aliases in REDEMPTION_ALIASES.items()
    }

    if columns["Артикул"] is None:
        return []

    rows = []

    for row_number in range(header_row + 1, worksheet.max_row + 1):
        article_cell = worksheet.cell(row_number, columns["Артикул"])

        if not valid_identifier(article_cell.value):
            continue

        level_column = columns.get("Уровень выкупа товара")
        level = (
            format_cell_value(worksheet.cell(row_number, level_column))
            if level_column is not None
            else None
        )

        rows.append(
            {
                "Артикул": display_identifier(article_cell.value),
                "Уровень выкупа товара": level,
            }
        )

    return rows


def build_index(rows, field):
    index = defaultdict(list)

    for row_index, record in enumerate(rows):
        for key in identifier_keys(record.get(field)):
            index[key].append((row_index, record))

    return index


def find_matches(record, sku_index, article_index):
    matches = []
    seen = set()

    for key in identifier_keys(record.get("SKU")):
        for row_index, matched_record in sku_index.get(key, []):
            if row_index not in seen:
                seen.add(row_index)
                matches.append((row_index, matched_record))

    if matches:
        return matches

    for key in identifier_keys(record.get("Артикул")):
        for row_index, matched_record in article_index.get(key, []):
            if row_index not in seen:
                seen.add(row_index)
                matches.append((row_index, matched_record))

    return matches


def find_redemption(record, redemption_index):
    for key in identifier_keys(record.get("Артикул")):
        matches = redemption_index.get(key, [])
        if matches:
            return matches[0]
    return None


def empty_result_record():
    return {column: None for column in FINAL_COLUMNS}


def combine_ozon_rows(unit_rows, stock_rows, redemption_rows):
    stock_by_sku = build_index(stock_rows, "SKU")
    stock_by_article = build_index(stock_rows, "Артикул")
    redemption_by_article = build_index(redemption_rows, "Артикул")

    used_stock_rows = set()
    used_redemption_rows = set()
    result = []

    def append_result(unit_record=None, stock_record=None):
        combined = empty_result_record()

        if unit_record:
            combined.update(unit_record)

        if stock_record:
            if not combined.get("SKU"):
                combined["SKU"] = stock_record.get("SKU")
            if not combined.get("Артикул"):
                combined["Артикул"] = stock_record.get("Артикул")
            for column in STOCK_COLUMNS:
                combined[column] = stock_record.get(column)

        redemption_match = find_redemption(
            combined,
            redemption_by_article,
        )
        if redemption_match:
            redemption_index, redemption_record = redemption_match
            used_redemption_rows.add(redemption_index)
            combined["Уровень выкупа товара"] = redemption_record.get(
                "Уровень выкупа товара"
            )

        combined["Реальная цена"] = calculate_real_price(combined)
        result.append(combined)

    for unit_record in unit_rows:
        matches = find_matches(
            unit_record,
            stock_by_sku,
            stock_by_article,
        )

        if not matches:
            append_result(unit_record=unit_record)
            continue

        for stock_index, stock_record in matches:
            used_stock_rows.add(stock_index)
            append_result(
                unit_record=unit_record,
                stock_record=stock_record,
            )

    for stock_index, stock_record in enumerate(stock_rows):
        if stock_index not in used_stock_rows:
            append_result(stock_record=stock_record)

    for redemption_index, redemption_record in enumerate(redemption_rows):
        if redemption_index in used_redemption_rows:
            continue

        combined = empty_result_record()
        combined["Артикул"] = redemption_record.get("Артикул")
        combined["Уровень выкупа товара"] = redemption_record.get(
            "Уровень выкупа товара"
        )
        result.append(combined)

    dataframe = pd.DataFrame(result, columns=FINAL_COLUMNS)

    for column in dataframe.columns:
        if column in ("SKU", "Артикул"):
            values = [display_identifier(value) for value in dataframe[column]]
        else:
            values = [clean_value(value) for value in dataframe[column]]
        dataframe[column] = pd.Series(values, dtype=object)

    return dataframe


def process_ozon_files(file_paths):
    unit_rows = []
    stock_rows = []
    redemption_rows = []
    report_counts = Counter()
    ignored_files = []
    errors = []

    for file_path in file_paths:
        workbook = None

        try:
            workbook = load_workbook_safely(file_path)
            detected = detect_report(workbook)

            if detected is None:
                ignored_files.append(Path(file_path).name)
                continue

            report_type, worksheet, header_row = detected
            report_counts[report_type] += 1

            if report_type == "unit":
                unit_rows.extend(extract_unit_rows(worksheet, header_row))
            elif report_type == "stock":
                stock_rows.extend(extract_stock_rows(worksheet, header_row))
            elif report_type == "redemption":
                redemption_rows.extend(
                    extract_redemption_rows(worksheet, header_row)
                )

        except Exception as error:
            errors.append(f"{Path(file_path).name}: {error}")
        finally:
            if workbook is not None:
                close_workbook(workbook)

    if not report_counts:
        details = "\n".join(errors[:5])
        if details:
            details = f"\n\n{details}"
        raise ValueError(
            "Не удалось распознать ни одного отчёта Ozon."
            f"{details}"
        )

    dataframe = combine_ozon_rows(
        unit_rows,
        stock_rows,
        redemption_rows,
    )

    return dataframe, report_counts, ignored_files, errors


def make_period(start_date, end_date, fallback_label=None):
    if start_date is not None and end_date is not None:
        label = (
            f"{start_date.strftime('%d.%m.%Y')}–"
            f"{end_date.strftime('%d.%m.%Y')}"
        )
        key = f"{start_date.isoformat()}_{end_date.isoformat()}"
    else:
        label = fallback_label or "Период без даты"
        key = f"unknown:{normalize_text(label)}"

    return {
        "start": start_date,
        "end": end_date,
        "label": label,
        "key": key,
    }


def period_from_text(value):
    text = normalize_text(value)
    patterns = (
        (
            r"\bс\s+(\d{2}\.\d{2}\.\d{4})\s+по\s+"
            r"(\d{2}\.\d{2}\.\d{4})",
            "%d.%m.%Y",
        ),
        (
            r"\bс\s+(\d{4}-\d{2}-\d{2})\s+по\s+"
            r"(\d{4}-\d{2}-\d{2})",
            "%Y-%m-%d",
        ),
    )

    for pattern, date_format in patterns:
        match = re.search(pattern, text)
        if match:
            start_date = datetime.strptime(
                match.group(1),
                date_format,
            ).date()
            end_date = datetime.strptime(
                match.group(2),
                date_format,
            ).date()
            return make_period(start_date, end_date)

    return None


def extract_workbook_period(workbook):
    for worksheet in workbook.worksheets:
        for row_number in range(1, min(worksheet.max_row, 12) + 1):
            for column in range(1, min(worksheet.max_column, 8) + 1):
                period = period_from_text(
                    worksheet.cell(row_number, column).value
                )
                if period is not None:
                    return period
    return None


def detect_wb_report(workbook):
    max_scan_rows = 25

    for worksheet in workbook.worksheets:
        for row_number in range(
            1,
            min(worksheet.max_row, max_scan_rows) + 1,
        ):
            keys = row_header_keys(worksheet, row_number)
            if (
                contains_exact(keys, "Артикул продавца")
                and contains_exact(keys, "Артикул WB")
                and contains_exact(keys, "Склад")
                and contains_exact(keys, "Текущий остаток, шт.")
                and contains_exact(keys, "Выкупили, шт.")
            ):
                return "supplier", worksheet, row_number

    for worksheet in workbook.worksheets:
        for row_number in range(
            1,
            min(worksheet.max_row, max_scan_rows) + 1,
        ):
            keys = row_header_keys(worksheet, row_number)
            if (
                contains_exact(keys, "Артикул продавца")
                and contains_exact(keys, "Артикул WB")
                and contains_exact(keys, "Итог, ₽")
                and contains_exact(keys, "Продажи, ₽")
                and contains_exact(keys, "Комиссия WB, ₽")
            ):
                return "finance", worksheet, row_number

    return None


def wb_base_record(worksheet, row_number, columns):
    identifier_columns = {
        "Сезон",
        "Коллекция",
        "Артикул продавца",
        "Артикул WB",
        "Баркод",
        "Размер",
        "Контракт",
    }
    record = {}

    for output_name in WB_BASE_COLUMNS:
        column = columns.get(output_name)
        value = (
            worksheet.cell(row_number, column).value
            if column is not None
            else None
        )
        if output_name in identifier_columns:
            value = display_identifier(value)
        else:
            value = clean_value(value)
        record[output_name] = value

    return record


def extract_wb_supplier_rows(worksheet, header_row, period):
    base_columns = {
        output_name: find_column(
            worksheet,
            (header_row,),
            aliases,
        )
        for output_name, aliases in WB_BASE_ALIASES.items()
    }
    metric_columns = {
        output_name: find_column(
            worksheet,
            (header_row,),
            aliases,
        )
        for output_name, aliases in WB_SUPPLIER_ALIASES.items()
    }
    rows = []

    for row_number in range(header_row + 1, worksheet.max_row + 1):
        base = wb_base_record(worksheet, row_number, base_columns)
        if not any(
            valid_identifier(base.get(field))
            for field in WB_PRODUCT_KEY_FIELDS
        ):
            continue

        metrics = {}
        for output_name in WB_SUPPLIER_METRICS:
            column = metric_columns.get(output_name)
            value = (
                worksheet.cell(row_number, column).value
                if column is not None
                else None
            )
            metrics[output_name] = clean_value(value)

        rows.append(
            {
                "base": base,
                "period": period,
                "metrics": metrics,
            }
        )

    return rows


def extract_wb_finance_rows(worksheet, header_row, current_period):
    base_columns = {
        output_name: find_column(
            worksheet,
            (header_row,),
            aliases,
        )
        for output_name, aliases in WB_BASE_ALIASES.items()
    }
    current_columns = {}
    previous_columns = {}

    for metric in WB_FINANCE_METRICS:
        current_columns[metric] = find_column(
            worksheet,
            (header_row,),
            (metric,),
        )
        previous_columns[metric] = find_column(
            worksheet,
            (header_row,),
            (f"{metric} (предыдущий период)",),
        )

    rows = []

    for row_number in range(header_row + 1, worksheet.max_row + 1):
        base = wb_base_record(worksheet, row_number, base_columns)
        if not any(
            valid_identifier(base.get(field))
            for field in WB_PRODUCT_KEY_FIELDS
        ):
            continue

        current_metrics = {}
        previous_metrics = {}

        for metric in WB_FINANCE_METRICS:
            current_column = current_columns.get(metric)
            previous_column = previous_columns.get(metric)
            current_metrics[metric] = clean_value(
                worksheet.cell(row_number, current_column).value
                if current_column is not None
                else None
            )
            previous_metrics[metric] = clean_value(
                worksheet.cell(row_number, previous_column).value
                if previous_column is not None
                else None
            )

        rows.append(
            {
                "base": base,
                "current_period": current_period,
                "current_metrics": current_metrics,
                "previous_metrics": previous_metrics,
            }
        )

    return rows


def safe_zip_excel_sources(file_path):
    sources = []

    with ZipFile(file_path, "r") as archive:
        for member in archive.infolist():
            member_path = PurePosixPath(member.filename.replace("\\", "/"))
            suffix = member_path.suffix.lower()

            if member.is_dir() or suffix not in (".xlsx", ".xlsm"):
                continue
            if member_path.is_absolute() or ".." in member_path.parts:
                raise ValueError("В ZIP найден небезопасный путь к файлу")
            if member.flag_bits & 0x1:
                raise ValueError("В ZIP найден Excel-файл с паролем")
            if member.file_size > MAX_ZIP_EXCEL_SIZE:
                raise ValueError("Excel-файл внутри ZIP слишком большой")

            content = BytesIO(archive.read(member))
            display_name = f"{Path(file_path).name} / {member_path.name}"
            sources.append((display_name, content, True))

    if not sources:
        raise ValueError("В ZIP не найдено файлов .xlsx или .xlsm")

    return sources


def wb_excel_sources(file_path):
    if Path(file_path).suffix.lower() == ".zip":
        return safe_zip_excel_sources(file_path)
    return [(Path(file_path).name, file_path, False)]


def period_sort_key(period):
    if period.get("start") is not None:
        return 0, period["start"], period.get("end")
    return 1, period.get("label", ""), period.get("key", "")


def previous_period_for(current_period, periods):
    if current_period and current_period.get("start") is not None:
        candidates = [
            period
            for period in periods
            if period.get("end") is not None
            and period["end"] < current_period["start"]
        ]
        if candidates:
            return max(candidates, key=lambda period: period["end"])

    label = "Предыдущий период"
    if current_period is not None:
        label = f"Предыдущий период перед {current_period['label']}"
    return make_period(None, None, label)


def aggregate_metric(old_value, new_value):
    if new_value is None or new_value == "":
        return old_value
    if old_value is None or old_value == "":
        return clean_value(new_value)

    old_number = number_value(old_value)
    new_number = number_value(new_value)
    if old_number is not None and new_number is not None:
        return clean_value(old_number + new_number)
    return new_value


def calculated_ratio(numerator, denominator, percent=False):
    numerator = number_value(numerator)
    denominator = number_value(denominator)
    if numerator is None or denominator in (None, 0):
        return None
    multiplier = 100 if percent else 1
    return clean_value(numerator / denominator * multiplier)


def calculated_change(current_value, previous_value):
    current_value = number_value(current_value)
    previous_value = number_value(previous_value)
    if current_value is None or previous_value in (None, 0):
        return None
    return clean_value(
        (current_value - previous_value) / abs(previous_value) * 100
    )


def positive_cost(value):
    value = number_value(value)
    if value is None:
        return None
    return clean_value(max(-value, 0))


def combine_costs(*values):
    costs = [positive_cost(value) for value in values]
    available = [value for value in costs if value is not None]
    if not available:
        return None
    return clean_value(sum(available))


def report_period_days(period):
    if (
        period is None
        or period.get("start") is None
        or period.get("end") is None
    ):
        return None
    return (period["end"] - period["start"]).days + 1


def combine_wb_reports(supplier_rows, finance_rows):
    product_bases = {}
    alias_indexes = {field: {} for field in WB_PRODUCT_KEY_FIELDS}
    next_product_id = 0

    def get_product_id(base):
        nonlocal next_product_id
        product_id = None

        for field in WB_PRODUCT_KEY_FIELDS:
            for key in identifier_keys(base.get(field)):
                if key in alias_indexes[field]:
                    product_id = alias_indexes[field][key]
                    break
            if product_id is not None:
                break

        if product_id is None:
            product_id = next_product_id
            next_product_id += 1
            product_bases[product_id] = {
                column: None for column in WB_BASE_COLUMNS
            }

        product_base = product_bases[product_id]
        for column in WB_BASE_COLUMNS:
            value = base.get(column)
            if value not in (None, "") and product_base.get(column) in (None, ""):
                product_base[column] = value

        for field in WB_PRODUCT_KEY_FIELDS:
            for key in identifier_keys(base.get(field)):
                alias_indexes[field][key] = product_id

        return product_id

    periods = {}
    product_supplier = defaultdict(dict)
    product_finance = defaultdict(dict)

    for row in supplier_rows:
        base = row["base"]
        period = row["period"]
        periods[period["key"]] = period
        product_id = get_product_id(base)
        target = product_supplier[(product_id, period["key"])]

        for metric, value in row["metrics"].items():
            target[metric] = aggregate_metric(target.get(metric), value)

    supplier_periods = list(periods.values())

    for row in finance_rows:
        base = row["base"]
        product_id = get_product_id(base)
        current_period = row["current_period"]
        previous_period = previous_period_for(
            current_period,
            supplier_periods,
        )
        periods[current_period["key"]] = current_period
        periods[previous_period["key"]] = previous_period

        for period, metrics in (
            (current_period, row["current_metrics"]),
            (previous_period, row["previous_metrics"]),
        ):
            target = product_finance[(product_id, period["key"])]
            for metric, value in metrics.items():
                target[metric] = aggregate_metric(target.get(metric), value)

    dated_periods = sorted(
        [period for period in periods.values() if period.get("start")],
        key=period_sort_key,
    )
    if dated_periods:
        current_period = dated_periods[-1]
        previous_period = (
            dated_periods[-2]
            if len(dated_periods) > 1
            else previous_period_for(current_period, periods.values())
        )
    else:
        sorted_periods = sorted(periods.values(), key=period_sort_key)
        current_period = sorted_periods[-1]
        previous_period = (
            sorted_periods[-2]
            if len(sorted_periods) > 1
            else make_period(None, None, "Предыдущий период")
        )

    period_days = report_period_days(current_period)
    result = []

    for product_id, base in product_bases.items():
        current_supplier = product_supplier.get(
            (product_id, current_period["key"]),
            {},
        )
        previous_supplier = product_supplier.get(
            (product_id, previous_period["key"]),
            {},
        )
        current_finance = product_finance.get(
            (product_id, current_period["key"]),
            {},
        )
        previous_finance = product_finance.get(
            (product_id, previous_period["key"]),
            {},
        )

        ordered = current_supplier.get("Заказано, шт.")
        redeemed = current_supplier.get("Выкупили, шт.")
        order_amount = current_supplier.get(
            "Сумма заказов минус комиссия WB, руб."
        )
        payout = current_supplier.get("К перечислению за товар, руб.")
        stock = current_supplier.get("Текущий остаток, шт.")

        previous_ordered = previous_supplier.get("Заказано, шт.")
        previous_redeemed = previous_supplier.get("Выкупили, шт.")

        sales_units = current_finance.get("Продажи, шт")
        sales_amount = current_finance.get("Продажи, ₽")
        returns_units_raw = current_finance.get("Возвраты, шт")
        returns_amount_raw = current_finance.get("Возвраты, ₽")
        result_after_wb = current_finance.get("Итог, ₽")
        previous_sales = previous_finance.get("Продажи, ₽")
        previous_result = previous_finance.get("Итог, ₽")

        ordered_number = number_value(ordered)
        redeemed_number = number_value(redeemed)
        unredeemed = None
        if ordered_number is not None and redeemed_number is not None:
            unredeemed = clean_value(max(ordered_number - redeemed_number, 0))

        logistics = positive_cost(current_finance.get("Логистика, ₽"))
        commission = positive_cost(current_finance.get("Комиссия WB, ₽"))
        acquiring = positive_cost(current_finance.get("Эквайринг, ₽"))
        fines_and_losses = combine_costs(
            current_finance.get("Штрафы, ₽"),
            current_finance.get("Потери, подмены и товары с дефектами, ₽"),
        )

        sales_number = number_value(sales_amount)
        returns_number = number_value(returns_amount_raw)
        result_number = number_value(result_after_wb)
        wb_expenses = None
        if sales_number is not None and result_number is not None:
            net_sales = sales_number + (returns_number or 0)
            wb_expenses = clean_value(net_sales - result_number)

        sales_per_day = calculated_ratio(sales_units, period_days)
        stock_days = calculated_ratio(stock, sales_per_day)

        row = {
            "Наименование": base.get("Наименование"),
            "Артикул продавца": base.get("Артикул продавца"),
            "Артикул WB": base.get("Артикул WB"),
            "Баркод": base.get("Баркод"),
            "Заказано, шт": ordered,
            "Выкуплено, шт": redeemed,
            "Невыкуплено, шт": unredeemed,
            "Выкуп, %": calculated_ratio(redeemed, ordered, percent=True),
            "Заказы, ₽": order_amount,
            "К перечислению, ₽": payout,
            "Средний заказ, ₽": calculated_ratio(order_amount, ordered),
            "Выплата за выкуп, ₽": calculated_ratio(payout, redeemed),
            "Остаток, шт": stock,
            "Продажи, шт": sales_units,
            "Продажи, ₽": sales_amount,
            "Средняя продажа, ₽": calculated_ratio(sales_amount, sales_units),
            "Возвраты, шт": (
                clean_value(abs(number_value(returns_units_raw)))
                if number_value(returns_units_raw) is not None
                else None
            ),
            "Возвраты, ₽": (
                clean_value(abs(returns_number))
                if returns_number is not None
                else None
            ),
            "Логистика, ₽": logistics,
            "Комиссия WB, ₽": commission,
            "Эквайринг, ₽": acquiring,
            "Штрафы и потери, ₽": fines_and_losses,
            "Расходы WB, ₽": wb_expenses,
            "Расходы на продажу, ₽": calculated_ratio(
                wb_expenses,
                sales_units,
            ),
            "Расходы WB, %": calculated_ratio(
                wb_expenses,
                sales_amount,
                percent=True,
            ),
            "Результат после WB, ₽": result_after_wb,
            "Остаётся после WB, %": calculated_ratio(
                result_after_wb,
                sales_amount,
                percent=True,
            ),
            "Продаж в день": sales_per_day,
            "Запас, дней": stock_days,
            "Продажи ранее, ₽": previous_sales,
            "Динамика продаж, %": calculated_change(
                sales_amount,
                previous_sales,
            ),
            "Результат ранее, ₽": previous_result,
            "Динамика результата, %": calculated_change(
                result_after_wb,
                previous_result,
            ),
            "Заказано ранее, шт": previous_ordered,
            "Динамика заказов, %": calculated_change(
                ordered,
                previous_ordered,
            ),
            "Выкуплено ранее, шт": previous_redeemed,
            "Динамика выкупов, %": calculated_change(
                redeemed,
                previous_redeemed,
            ),
        }
        result.append(row)

    result.sort(
        key=lambda row: (
            normalize_text(row.get("Наименование")),
            normalize_text(row.get("Артикул WB")),
        )
    )
    dataframe = pd.DataFrame(result, columns=WB_RESULT_COLUMNS)
    for column in dataframe.columns:
        if column in ("Артикул продавца", "Артикул WB", "Баркод"):
            values = [display_identifier(value) for value in dataframe[column]]
        else:
            values = [clean_value(value) for value in dataframe[column]]
        dataframe[column] = pd.Series(values, dtype=object)

    return dataframe, [previous_period["label"], current_period["label"]]


def process_wb_files(file_paths):
    supplier_rows = []
    finance_rows = []
    report_counts = Counter()
    ignored_files = []
    errors = []
    unknown_period_number = 0

    for file_path in file_paths:
        try:
            sources = wb_excel_sources(file_path)
        except Exception as error:
            errors.append(f"{Path(file_path).name}: {error}")
            continue

        for display_name, file_source, should_close in sources:
            workbook = None
            try:
                workbook = load_workbook_safely(file_source)
                detected = detect_wb_report(workbook)
                if detected is None:
                    ignored_files.append(display_name)
                    continue

                report_type, worksheet, header_row = detected
                report_counts[report_type] += 1
                period = extract_workbook_period(workbook)
                if period is None:
                    unknown_period_number += 1
                    period = make_period(
                        None,
                        None,
                        f"Период {unknown_period_number}",
                    )

                if report_type == "supplier":
                    supplier_rows.extend(
                        extract_wb_supplier_rows(
                            worksheet,
                            header_row,
                            period,
                        )
                    )
                elif report_type == "finance":
                    finance_rows.extend(
                        extract_wb_finance_rows(
                            worksheet,
                            header_row,
                            period,
                        )
                    )

            except Exception as error:
                errors.append(f"{display_name}: {error}")
            finally:
                if workbook is not None:
                    close_workbook(workbook)
                if should_close:
                    file_source.close()

    if not report_counts:
        details = "\n".join(errors[:5])
        if details:
            details = f"\n\n{details}"
        raise ValueError(
            "Не удалось распознать ни одного отчёта Wildberries."
            f"{details}"
        )

    dataframe, periods = combine_wb_reports(
        supplier_rows,
        finance_rows,
    )
    return dataframe, report_counts, ignored_files, errors, periods


class MarketplaceViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Помощник маркетплейсов")
        self.root.geometry("1280x720")
        self.root.minsize(900, 550)

        self.ozon_df = None
        self.wb_df = None
        self.ozon_table = None
        self.ozon_model = None
        self.wb_table = None
        self.wb_model = None

        self.style = ttk.Style()
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.ozon_tab = ttk.Frame(self.notebook, padding=10)
        self.wildberries_tab = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(self.ozon_tab, text="Ozon")
        self.notebook.add(self.wildberries_tab, text="Wildberries")

        self.build_ozon_tab()
        self.build_wildberries_tab()

    def build_ozon_tab(self):
        title = ttk.Label(
            self.ozon_tab,
            text="Раздел Ozon",
            font=("Arial", 16, "bold"),
        )
        title.pack(anchor="w")

        description = ttk.Label(
            self.ozon_tab,
            text=(
                "Выберите все отчёты сразу. Программа определит их по "
                "заголовкам и объединит товары по SKU, а если SKU нет — "
                "по артикулу. Названия файлов могут быть любыми."
            ),
            wraplength=1050,
        )
        description.pack(anchor="w", pady=(4, 10))

        controls = ttk.Frame(self.ozon_tab)
        controls.pack(fill="x")

        load_button = ttk.Button(
            controls,
            text="Загрузить файлы Ozon",
            command=self.open_ozon_files,
        )
        load_button.pack(side="left")

        ttk.Label(
            controls,
            text="Поиск по SKU или артикулу:",
        ).pack(side="left", padx=(20, 6))

        self.search_entry = ttk.Entry(controls, width=28)
        self.search_entry.pack(side="left")
        self.search_entry.bind("<Return>", lambda event: self.search_products())

        search_button = ttk.Button(
            controls,
            text="Найти",
            command=self.search_products,
        )
        search_button.pack(side="left", padx=6)

        reset_button = ttk.Button(
            controls,
            text="Показать всё",
            command=self.reset_search,
        )
        reset_button.pack(side="left")

        self.ozon_status = ttk.Label(
            self.ozon_tab,
            text="Файлы ещё не загружены",
        )
        self.ozon_status.pack(anchor="w", pady=(10, 8))

        self.ozon_table_frame = ttk.Frame(self.ozon_tab)
        self.ozon_table_frame.pack(fill="both", expand=True)

    def build_wildberries_tab(self):
        title = ttk.Label(
            self.wildberries_tab,
            text="Раздел Wildberries",
            font=("Arial", 16, "bold"),
        )
        title.pack(anchor="w")

        description = ttk.Label(
            self.wildberries_tab,
            text=(
                "Выберите отчёты Wildberries и ZIP-архивы сразу. Программа "
                "распознает их по заголовкам, определит периоды и объединит "
                "товары по артикулу WB, артикулу продавца или баркоду."
            ),
            wraplength=1050,
        )
        description.pack(anchor="w", pady=(4, 10))

        explanation = ttk.Label(
            self.wildberries_tab,
            text=(
                "Заказы, выкупы и остатки берутся из отчётов по товарам. "
                "Продажи, расходы и результат — из отчёта «Доходы и расходы». "
                "Результат после WB рассчитан до вычета себестоимости товара."
            ),
            wraplength=1050,
        )
        explanation.pack(anchor="w", pady=(0, 10))

        controls = ttk.Frame(self.wildberries_tab)
        controls.pack(fill="x")

        load_button = ttk.Button(
            controls,
            text="Загрузить файлы Wildberries",
            command=self.open_wb_files,
        )
        load_button.pack(side="left")

        ttk.Label(
            controls,
            text="Поиск по артикулу, баркоду или названию:",
        ).pack(side="left", padx=(20, 6))

        self.wb_search_entry = ttk.Entry(controls, width=28)
        self.wb_search_entry.pack(side="left")
        self.wb_search_entry.bind(
            "<Return>",
            lambda event: self.search_wb_products(),
        )

        search_button = ttk.Button(
            controls,
            text="Найти",
            command=self.search_wb_products,
        )
        search_button.pack(side="left", padx=6)

        reset_button = ttk.Button(
            controls,
            text="Показать всё",
            command=self.reset_wb_search,
        )
        reset_button.pack(side="left")

        self.wb_status = ttk.Label(
            self.wildberries_tab,
            text="Файлы ещё не загружены",
        )
        self.wb_status.pack(anchor="w", pady=(10, 8))

        self.wb_summary = ttk.Label(
            self.wildberries_tab,
            text="",
            font=("Arial", 10, "bold"),
            wraplength=1200,
        )
        self.wb_summary.pack(anchor="w", pady=(0, 8))

        self.wb_table_frame = ttk.Frame(self.wildberries_tab)
        self.wb_table_frame.pack(fill="both", expand=True)

    def open_ozon_files(self):
        if Table is None:
            messagebox.showerror("Ozon", "Для Ozon установите pandastable: python -m pip install pandastable")
            return
        file_paths = filedialog.askopenfilenames(
            title="Выберите все Excel-файлы Ozon",
            filetypes=[
                ("Excel-файлы", "*.xlsx *.XLSX *.xlsm *.XLSM"),
                ("Все файлы", "*.*"),
            ],
        )

        if not file_paths:
            return

        self.ozon_status.config(text="Чтение и сопоставление файлов...")
        self.root.update_idletasks()

        try:
            dataframe, counts, ignored, errors = process_ozon_files(file_paths)

            if dataframe.empty:
                messagebox.showerror(
                    "Ошибка",
                    "Отчёты распознаны, но строки с товарами не найдены.",
                )
                self.ozon_status.config(text="Данные не найдены")
                return

            self.ozon_df = dataframe
            self.search_entry.delete(0, tk.END)
            self.show_ozon_table(self.ozon_df)

            recognized = ", ".join(
                f"{REPORT_TITLES[report_type]}: {count}"
                for report_type, count in counts.items()
            )
            status = (
                f"Распознано — {recognized}. "
                f"Строк в таблице: {len(self.ozon_df)}"
            )
            self.ozon_status.config(text=status)

            warnings = []
            if ignored:
                warnings.append(
                    "Не распознаны и пропущены:\n- " + "\n- ".join(ignored)
                )
            if errors:
                warnings.append(
                    "Не удалось прочитать:\n- " + "\n- ".join(errors[:10])
                )
            if warnings:
                messagebox.showwarning(
                    "Файлы обработаны с предупреждениями",
                    "\n\n".join(warnings),
                )

        except Exception as error:
            self.ozon_status.config(text="Ошибка обработки файлов")
            messagebox.showerror(
                "Ошибка",
                f"Не удалось собрать отчёты Ozon:\n\n{error}",
            )

    def search_products(self):
        if self.ozon_df is None:
            messagebox.showwarning(
                "Предупреждение",
                "Сначала загрузите файлы Ozon",
            )
            return

        search_text = self.search_entry.get().strip()
        if not search_text:
            self.show_ozon_table(self.ozon_df)
            return

        sku_mask = self.ozon_df["SKU"].astype(str).str.contains(
            search_text,
            case=False,
            na=False,
            regex=False,
        )
        article_mask = self.ozon_df["Артикул"].astype(str).str.contains(
            search_text,
            case=False,
            na=False,
            regex=False,
        )
        self.show_ozon_table(
            self.ozon_df[sku_mask | article_mask].copy()
        )

    def reset_search(self):
        if self.ozon_df is None:
            return
        self.search_entry.delete(0, tk.END)
        self.show_ozon_table(self.ozon_df)

    def show_ozon_table(self, dataframe):
        if self.ozon_table is not None:
            self.ozon_table.destroy()
            self.ozon_table = None

        display_dataframe = dataframe.reset_index(drop=True).copy()
        self.ozon_model = TableModel(dataframe=display_dataframe)
        self.ozon_table = Table(
            self.ozon_table_frame,
            model=self.ozon_model,
            showtoolbar=False,
            showstatusbar=False,
        )
        self.ozon_table.adjustColumnWidths = lambda *args, **kwargs: None
        self.ozon_table.floatprecision = 2

        for column in display_dataframe.columns:
            self.ozon_table.columnwidths[column] = 135

        self.ozon_table.show()
        self.ozon_status.config(
            text=(
                f"Показано строк: {len(display_dataframe)} "
                f"из {len(self.ozon_df)}"
            )
        )

    def open_wb_files(self):
        file_paths = filedialog.askopenfilenames(
            title="Выберите файлы Wildberries и ZIP-архивы",
            filetypes=[
                (
                    "Wildberries",
                    "*.xlsx *.XLSX *.xlsm *.XLSM *.zip *.ZIP",
                ),
                ("Excel-файлы", "*.xlsx *.XLSX *.xlsm *.XLSM"),
                ("ZIP-архивы", "*.zip *.ZIP"),
                ("Все файлы", "*.*"),
            ],
        )

        if not file_paths:
            return

        self.wb_status.config(text="Чтение и сопоставление файлов...")
        self.root.update_idletasks()

        try:
            dataframe, counts, ignored, errors, periods = process_wb_files(
                file_paths
            )

            if dataframe.empty:
                messagebox.showerror(
                    "Ошибка",
                    "Отчёты распознаны, но строки с товарами не найдены.",
                )
                self.wb_status.config(text="Данные не найдены")
                return

            self.wb_df = dataframe
            self.wb_search_entry.delete(0, tk.END)
            self.show_wb_table(self.wb_df)

            recognized = ", ".join(
                f"{WB_REPORT_TITLES[report_type]}: {count}"
                for report_type, count in counts.items()
            )
            period_text = ", ".join(periods)
            self.wb_status.config(
                text=(
                    f"Распознано — {recognized}. Периоды: {period_text}. "
                    f"Строк в таблице: {len(self.wb_df)}"
                )
            )

            warnings = []
            if ignored:
                warnings.append(
                    "Не распознаны и пропущены:\n- " + "\n- ".join(ignored)
                )
            if errors:
                warnings.append(
                    "Не удалось прочитать:\n- " + "\n- ".join(errors[:10])
                )
            if warnings:
                messagebox.showwarning(
                    "Файлы обработаны с предупреждениями",
                    "\n\n".join(warnings),
                )

        except Exception as error:
            self.wb_status.config(text="Ошибка обработки файлов")
            messagebox.showerror(
                "Ошибка",
                f"Не удалось собрать отчёты Wildberries:\n\n{error}",
            )

    def search_wb_products(self):
        if self.wb_df is None:
            messagebox.showwarning(
                "Предупреждение",
                "Сначала загрузите файлы Wildberries",
            )
            return

        search_text = self.wb_search_entry.get().strip()
        if not search_text:
            self.show_wb_table(self.wb_df)
            return

        mask = pd.Series(False, index=self.wb_df.index)
        for column in (
            "Артикул продавца",
            "Артикул WB",
            "Баркод",
            "Наименование",
        ):
            mask = mask | self.wb_df[column].astype(str).str.contains(
                search_text,
                case=False,
                na=False,
                regex=False,
            )

        self.show_wb_table(self.wb_df[mask].copy())

    def reset_wb_search(self):
        if self.wb_df is None:
            return
        self.wb_search_entry.delete(0, tk.END)
        self.show_wb_table(self.wb_df)

    def show_wb_table(self, dataframe):
        if self.wb_table is not None:
            self.wb_table.destroy()
            self.wb_table = None

        display_dataframe = dataframe.reset_index(drop=True).copy()
        self.wb_model = TableModel(dataframe=display_dataframe)
        self.wb_table = Table(
            self.wb_table_frame,
            model=self.wb_model,
            showtoolbar=False,
            showstatusbar=False,
        )
        self.wb_table.adjustColumnWidths = lambda *args, **kwargs: None
        self.wb_table.floatprecision = 2

        for column in display_dataframe.columns:
            width = 135
            if column == "Наименование":
                width = 260
            elif column in ("Артикул продавца", "Артикул WB", "Баркод"):
                width = 155
            elif (
                "Динамика" in column
                or "Результат" in column
                or "Штрафы" in column
            ):
                width = 175
            self.wb_table.columnwidths[column] = width

        self.wb_table.show()
        self.update_wb_summary(display_dataframe)
        self.wb_status.config(
            text=(
                f"Показано строк: {len(display_dataframe)} "
                f"из {len(self.wb_df)}"
            )
        )

    def update_wb_summary(self, dataframe):
        def total(column):
            values = pd.to_numeric(dataframe[column], errors="coerce")
            return float(values.fillna(0).sum())

        def formatted(value):
            return f"{value:,.0f}".replace(",", " ")

        ordered = total("Заказано, шт")
        redeemed = total("Выкуплено, шт")
        sales = total("Продажи, ₽")
        expenses = total("Расходы WB, ₽")
        result = total("Результат после WB, ₽")
        stock = total("Остаток, шт")
        redemption_rate = redeemed / ordered * 100 if ordered else 0
        result_rate = result / sales * 100 if sales else 0

        self.wb_summary.config(
            text=(
                f"ИТОГО: товаров — {len(dataframe)} | "
                f"заказано — {formatted(ordered)} шт | "
                f"выкуп — {redemption_rate:.1f}% | "
                f"продажи — {formatted(sales)} ₽ | "
                f"расходы WB — {formatted(expenses)} ₽ | "
                f"результат после WB — {formatted(result)} ₽ "
                f"({result_rate:.1f}%) | "
                f"остаток — {formatted(stock)} шт"
            )
        )



# WB: исходные периодные итоги не распределяются пропорционально дням.
from datetime import timedelta
import hashlib
import threading


def wb_date(value):
    text = str(value or '').strip()
    for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%d.%m.%Y'):
        try:
            return datetime.strptime(text[:10], fmt).date()
        except ValueError:
            pass
    return None


def wb_period(value):
    dates = re.findall(r'\d{4}-\d{2}-\d{2}|\d{2}[.-]\d{2}[.-]\d{4}', str(value))
    if len(dates) >= 2:
        a, b = map(wb_date, dates[:2])
        if a and b and a <= b:
            return a, b
    return None


def wb_label(p):
    return f'{p[0]:%d.%m.%Y} — {p[1]:%d.%m.%Y}'


def wb_read_table(ws):
    rows = list(ws.values)
    for i, row in enumerate(rows[:10]):
        if 'Артикул WB' in row or '№ отчета' in row:
            headers = [str(v or '').strip() for v in row]
            return [(i + j + 2, dict(zip(headers, r))) for j, r in enumerate(rows[i+1:]) if any(v is not None for v in r)]
    return []


def wb_clean_frame(records):
    df = pd.DataFrame(records)
    if df.empty:
        return df
    # Нули скрываются только в отображении; исходные записи сохранены.
    keep = []
    for c in df:
        s = df[c].dropna()
        if s.empty or s.astype(str).str.strip().isin(['', '-']).all():
            continue
        n = pd.to_numeric(s, errors='coerce')
        if n.notna().all() and n.eq(0).all() and c not in ('Артикул WB', 'Артикул продавца'):
            continue
        keep.append(c)
    return df[keep]


def analyze_wb(file_paths, keep_zero=False):
    groups = defaultdict(list)
    warnings, audit, coverage = [], [], []
    seen_files = set()
    for path in file_paths:
        for name, source, close in wb_excel_sources(path):
            w = None
            try:
                if hasattr(source, 'read'):
                    source.seek(0); payload = source.read(); source.seek(0)
                else:
                    payload = Path(source).read_bytes()
                digest = hashlib.sha256(payload).hexdigest()
                if digest in seen_files:
                    warnings.append(f'{name}: повторный файл пропущен'); continue
                seen_files.add(digest)
                w = load_workbook_safely(source)
                p = prev = None
                if 'Общая информация' in w.sheetnames:
                    for r in w['Общая информация'].values:
                        q = wb_period(' '.join(str(x or '') for x in r))
                        if q:
                            if 'прошл' in str(r[0]).lower() or 'предыдущ' in str(r[0]).lower(): prev = q
                            elif p is None: p = q
                found = False
                mapping = {'Товары':'Воронка','Промосервисы':'Промо','Детальная информация':'Остатки: детали',
                           'Остатки по дням':'Остатки: дни','Все заказы':'Лента','Детальные показатели':'Поиск'}
                for ws in w:
                    rows = wb_read_table(ws)
                    if not rows: continue
                    headers = rows[0][1]
                    if '№ отчета' in headers: kind = 'Финансы'
                    elif 'Комиссия WB, ₽' in headers and 'Итог, ₽' in headers: kind = 'Доходы и расходы'
                    elif 'Текущий остаток, шт.' in headers and 'Выкупили, шт.' in headers: kind = 'Поставщик'
                    else: kind = mapping.get(ws.title)
                    if p is None:
                        p = next((wb_period(' '.join(str(v or '') for v in r)) for r in list(ws.values)[:10] if wb_period(' '.join(str(v or '') for v in r))),None)
                    if not kind:
                        warnings.append(f'{name}/{ws.title}: тип отчёта пока не поддержан; строки сохранены в исходных')
                        for lineno,data in rows:audit.append({'Тип':'Не поддержано','Источник':f'{name} / {ws.title} / строка {lineno}',**data})
                        found=True
                        continue
                    if kind != 'Финансы' and p is None:
                        warnings.append(f'{name}/{ws.title}: не найден период, лист пропущен'); continue
                    found = True
                    if p and kind not in ('Финансы','Промо','Остатки: детали'):
                        coverage.append((p,kind))
                    for lineno, data in rows:
                        origin = f'{name} / {ws.title} / строка {lineno}'
                        groups[kind].append((p, prev, data, origin))
                        audit.append({'Тип':kind,'Источник':origin,**data})
                if not found: warnings.append(f'{name}: формат не распознан')
            except Exception as exc:
                warnings.append(f'{name}: {exc}')
            finally:
                if w is not None: close_workbook(w)
                if close: source.close()
    products = {}
    def identity(d):
        return display_identifier(d.get('Артикул WB')) or 'seller:' + display_identifier(d.get('Артикул продавца'))
    def base(d):
        return {'Артикул WB':display_identifier(d.get('Артикул WB')),
                'Артикул продавца':display_identifier(d.get('Артикул продавца')),
                'SKU':display_identifier(d.get('SKU')),
                'Наименование':d.get('Название', d.get('Наименование', ''))}
    meta = {'Артикул WB','Артикул продавца','SKU','Название','Предмет','Бренд','Размер','Склад'}
    def product(p,d):
        key = (p,identity(d))
        return products.setdefault(key,{'Период':wb_label(p),**base(d),'Источники':''})
    # Детали остатков содержат несколько складов/размеров одного товара.
    stock_groups = defaultdict(list)
    seen_stock = set()
    for p, prev, d, origin in groups['Остатки: детали']:
        key = (p, identity(d), str(d.get('Размер')), str(d.get('Склад')))
        if key in seen_stock:
            warnings.append(origin + ': повтор склада/размера за период пропущен')
            continue
        seen_stock.add(key)
        stock_groups[(p, identity(d))].append((d, origin))
    groups['Остатки: детали'] = []
    for (p, pid), records in stock_groups.items():
        merged = dict(records[0][0])
        if len(records) > 1:
            for c in merged:
                if c in meta:
                    continue
                values = [d.get(c) for d, _ in records if d.get(c) is not None]
                nums = [number_value(v) for v in values]
                if ('шт' in c or 'на сумму' in c or 'Стоимость остатков' in c) and all(v is not None for v in nums):
                    merged[c] = sum(nums)
                elif len(set(map(str, values))) > 1:
                    merged[c] = None
            merged['Процент выкупа'] = None
        groups['Остатки: детали'].append((p, None, merged, '; '.join(o for _, o in records)))
    # В старых отчётах один товар разбит по складам/баркодам/размерам.
    # Сначала сводим эти строки, затем присоединяем к полному периоду.
    for kind in ('Поставщик','Доходы и расходы'):
        chunks = defaultdict(list); seen = {}
        for p,prev,d,origin in groups[kind]:
            key=(p,identity(d),str(d.get('Размер')),str(d.get('Склад')),str(d.get('Баркод')))
            if key in seen:
                if seen[key] != d: warnings.append(origin+': конфликт версии, сохранена первая; обе версии есть в исходных')
                continue
            seen[key]=d;chunks[(p,identity(d))].append((prev,d,origin))
        groups[kind]=[]
        for (p,pid),items in chunks.items():
            merged=dict(items[0][1])
            for c in merged:
                if c in meta or c in ('Баркод','Контракт','Сезон','Коллекция','Наименование'):continue
                vals=[number_value(d.get(c)) for _,d,_ in items]
                if all(v is not None for v in vals):merged[c]=sum(vals)
            groups[kind].append((p,items[0][0],merged,'; '.join(x[2] for x in items)))
    for kind in ('Воронка','Промо','Остатки: детали','Поставщик','Доходы и расходы'):
        for p,prev,d,origin in groups[kind]:
            row = product(p,d)
            row['Источники'] += ('; ' if row['Источники'] else '') + origin
            for c,v in d.items():
                if c in meta or c in ('Баркод','Контракт','Сезон','Коллекция','Наименование') or v is None: continue
                target = row
                if 'предыдущий период' in c:
                    if not prev:
                        target[f'{kind} | {c} [даты не указаны]']=v
                        continue
                    target = product(prev,d)
                    target['Источники'] = origin + ' (предыдущий период)'
                    c = c.replace(' (предыдущий период)','').replace('Выкупы, шт','Выкупили, шт')
                key = f'{kind} | {c}'
                if key in target and target[key] != v:
                    warnings.append(f'Конфликт {target["Период"]}, {identity(d)}, {key}: сохранено первое значение; оба есть в исходных')
                else: target[key] = v
    def ratio(row, label, a, b, factor=1):
        a,b = number_value(a),number_value(b)
        if a is not None and b not in (None,0): row['Расчёт | '+label] = round(a/b*factor,2)
    for (p,_),row in products.items():
        get=lambda c: row.get('Воронка | '+c)
        n, redeemed, cancelled = map(number_value,(get('Заказали товаров, шт'),get('Выкупили, шт'),get('Отменили, шт')))
        ratio(row,'Заказов в день',n,(p[1]-p[0]).days+1)
        ratio(row,'Средняя сумма заказа, ₽',get('Заказали на сумму, ₽'),n)
        ratio(row,'Средняя сумма выкупа, ₽',get('Выкупили на сумму, ₽'),redeemed)
        ratio(row,'Выкуплено / заказано, %',redeemed,n,100)
        ratio(row,'Отменено / заказано, %',cancelled,n,100)
        ratio(row,'Конверсия переход → заказ, %',n,get('Переходы в карточку'),100)
        if None not in (n,redeemed,cancelled): row['Расчёт | Незавершённые заказы, шт'] = n-redeemed-cancelled
        if get('Заказали товаров, шт') is not None and row.get('Остатки: детали | Заказали, шт') is not None:
            row['Расчёт | Расхождение заказов между отчётами, шт'] = n-number_value(row['Остатки: детали | Заказали, шт'])
    # Только источники с дневной детализацией можно разрезать по границам периодов.
    cuts=sorted({x for p,k in coverage for x in (p[0],p[1]+timedelta(days=1))})
    segments=[(a,b-timedelta(days=1)) for a,b in zip(cuts,cuts[1:]) if any(p[0]<=a<=p[1] for p,k in coverage)]
    dayrows={}; order_ids={}; daily_stock={}
    def segment(day):
        return next((p for p in segments if p[0]<=day<=p[1]),(day,day))
    def daily(p,d):
        return dayrows.setdefault((p,identity(d)),{'Период':wb_label(p),**base(d)})
    for p,prev,d,origin in groups['Лента']:
        key=d.get('ID заказа')
        if not key:
            warnings.append(origin+': отсутствует ID заказа; запись доступна только в исходных'); continue
        old=order_ids.get(key)
        if old is None or str(d.get('Дата текущего статуса'))>str(old[2].get('Дата текущего статуса')):
            order_ids[key]=(p,prev,d,origin)
    for p,prev,d,origin in order_ids.values():
        day=wb_date(d.get('Дата оформления заказа'))
        if not day: warnings.append(origin+': неверная дата заказа'); continue
        row=daily(segment(day),d)
        for c,v in [('Лента | Заказов, шт',1),('Лента | Сумма заказов, ₽',number_value(d.get('Цена со скидкой продавца'))),('Лента | '+str(d.get('Статус заказа'))+', шт',1)]:
            if v is not None: row[c]=row.get(c,0)+v
    for p,prev,d,origin in groups['Остатки: дни']:
        for c,v in d.items():
            day=wb_date(c)
            if not day: continue
            key=(identity(d),str(d.get('Размер')),str(d.get('Склад')),day)
            if key in daily_stock and daily_stock[key][0]!=v: warnings.append(origin+': конфликт дневного остатка, сохранено первое значение')
            else: daily_stock[key]=(v,d)
    by_product=defaultdict(dict)
    for (pid,size,warehouse,day),(v,d) in daily_stock.items():
        n=number_value(v)
        if n is None: continue
        pp=segment(day); daily(pp,d)
        series=by_product[(pp,pid)]
        series[day]=series.get(day,0)+n
    for key,series in by_product.items():
        row=dayrows[key]; dates=sorted(series); values=[series[x] for x in dates]
        row.update({'Остатки | На начало, шт':values[0],'Остатки | На конец, шт':values[-1],
                    'Расчёт | Средний остаток, шт':round(sum(values)/len(values),2),
                    'Расчёт | Дней с нулевым остатком':sum(v==0 for v in values),
                    'Расчёт | Изменение остатка, шт':values[-1]-values[0],
                    'Расчёт | Дней наблюдения остатков':len(values)})
    for (p,pid),row in dayrows.items():
        active=sorted({k for q,k in coverage if q[0]<=p[0] and q[1]>=p[1]})
        ratio(row,'Заказов в день',row.get('Лента | Заказов, шт'),(p[1]-p[0]).days+1)
        ratio(row,'Выкуплено / заказы ленты, %',row.get('Лента | Выкуплен, шт'),row.get('Лента | Заказов, шт'),100)
        ratio(row,'Средняя сумма заказа ленты, ₽',row.get('Лента | Сумма заказов, ₽'),row.get('Лента | Заказов, шт'))
        if 'Лента' in active and 'Остатки: дни' in active:
            ratio(row,'Запас по темпу заказов, дней (оценка)',row.get('Остатки | На конец, шт'),row.get('Расчёт | Заказов в день'))
    search=[]; seen_search=set()
    for p,prev,d,origin in groups['Поиск']:
        for previous,q in ((False,p),(True,prev)):
            if q is None: continue
            row={'Период':wb_label(q),**base(d),'Поисковый запрос':d.get('Поисковый запрос'),'Источник':origin}
            for c,v in d.items():
                if c in meta or c=='Поисковый запрос':continue
                if ('предыдущий период' in c)==previous:
                    row['Поиск | '+c.replace(' (предыдущий период)','')]=v
            signature=(q,identity(d),d.get('Поисковый запрос'))
            if signature in seen_search: warnings.append(origin+': повтор товара/запроса/периода пропущен');continue
            seen_search.add(signature)
            ratio(row,'Конверсия переход → заказ, %',row.get('Поиск | Заказали, шт'),row.get('Поиск | Переходы в карточку'),100)
            search.append(row)
    finance=[]; seen_reports=set()
    for _,_,d,origin in groups['Финансы']:
        key=(d.get('№ отчета'),d.get('Юридическое лицо'),d.get('Валюта'))
        if key in seen_reports: warnings.append(origin+': повтор номера отчёта пропущен');continue
        seen_reports.add(key)
        a,b=wb_date(d.get('Дата начала')),wb_date(d.get('Дата конца'))
        row={'Период':wb_label((a,b)) if a and b else 'Не определён','Уровень':'Весь магазин',**{'Финансы | '+c:v for c,v in d.items()},'Источник':origin}
        ratio(row,'К оплате / продажи, %',d.get('Итого к оплате'),d.get('Продажа'),100)
        ratio(row,'Доставка / продажи, %',d.get('Стоимость доставки'),d.get('Продажа'),100)
        finance.append(row)
    if not audit: raise ValueError('Ни один поддерживаемый отчёт не прочитан. '+ '\n'.join(warnings))
    product_records=[r for key,r in sorted(products.items(),key=lambda item:item[0])]
    segment_records=[r for key,r in sorted(dayrows.items(),key=lambda item:(-sum(q[0]<=item[0][0][0] and q[1]>=item[0][0][1] for q,k in coverage),item[0]))]
    # Убираем полностью неактивные товары, но не товары с остатком или историей.
    def active(row):
        vals=[number_value(v) for c,v in row.items() if '|' in c and not any(t in c.lower() for t in ('рейтинг','цена','процент','конверсия','локальные','дней наблюдения'))]
        return any(v is not None and v!=0 for v in vals)
    result = {'Товары по периодам':pd.DataFrame(product_records) if keep_zero else wb_clean_frame([r for r in product_records if active(r)]),
            'Пересечения дат':pd.DataFrame(segment_records) if keep_zero else wb_clean_frame([r for r in segment_records if active(r)]),
            'Поисковые запросы':pd.DataFrame(search) if keep_zero else wb_clean_frame(search),'Финансы магазина':pd.DataFrame(finance) if keep_zero else wb_clean_frame(finance),
            'Исходные данные':pd.DataFrame(audit)}
    if warnings: result['Замечания'] = pd.DataFrame({'Сообщение':warnings})
    return result, warnings


class MarketplaceViewer(MarketplaceViewer):
    def build_wildberries_tab(self):
        self.wb_paths=[]; self.wb_views={}
        ttk.Label(self.wildberries_tab,text='Wildberries — автоматическое объединение',font=('Arial',16,'bold')).pack(anchor='w')
        ttk.Label(self.wildberries_tab,text='Исходные колонки подписаны названием отчёта. Добавленные — «Расчёт». Пустое значение означает отсутствие данных.\nСовпадающие полные периоды объединены по артикулу WB. Дневные данные разделены по пересечениям. Финансы — по всему магазину.',wraplength=1150).pack(anchor='w',pady=8)
        bar=ttk.Frame(self.wildberries_tab);bar.pack(fill='x')
        self.wb_load=ttk.Button(bar,text='Добавить отчёты',command=self.open_wb_files);self.wb_load.pack(side='left')
        ttk.Button(bar,text='Очистить',command=self.clear_wb).pack(side='left',padx=5)
        self.wb_search_entry=ttk.Entry(bar,width=30);self.wb_search_entry.pack(side='left',padx=5)
        self.wb_search_entry.bind('<Return>',lambda e:self.render_wb())
        ttk.Button(bar,text='Найти',command=self.render_wb).pack(side='left')
        ttk.Button(bar,text='Сбросить поиск',command=self.reset_wb_search).pack(side='left',padx=5)
        ttk.Button(bar,text='Экспорт CSV',command=self.export_wb_csv).pack(side='right')
        self.wb_status=ttk.Label(self.wildberries_tab,text='Выберите один или несколько файлов');self.wb_status.pack(anchor='w',pady=8)
        self.wb_tabs=ttk.Notebook(self.wildberries_tab);self.wb_tabs.pack(fill='both',expand=True)
    def clear_wb(self):
        if getattr(self,'wb_busy',False): return
        self.wb_paths=[];self.wb_views={};self.render_wb();self.wb_status.config(text='Данные очищены')
    def open_wb_files(self):
        paths=filedialog.askopenfilenames(filetypes=[('Отчёты WB','*.xlsx *.xlsm *.zip')])
        if not paths:return
        candidate=list(dict.fromkeys(self.wb_paths+list(paths)))
        self.wb_busy=True;self.wb_load.config(state='disabled');self.wb_status.config(text='Чтение и объединение…')
        box={}
        def work():
            try:box['result']=analyze_wb(candidate)
            except Exception as e:box['error']=str(e)
        threading.Thread(target=work,daemon=True).start()
        def poll():
            if not box:self.root.after(100,poll);return
            self.wb_busy=False;self.wb_load.config(state='normal')
            if 'error' in box:
                self.wb_status.config(text='Ошибка загрузки');messagebox.showerror('Wildberries',box['error']);return
            self.wb_views,warnings=box['result'];self.wb_paths=candidate
            self.render_wb();self.wb_status.config(text=f'Файлов: {len(candidate)}. Нулевые колонки скрыты. Исходные записи сохранены. Замечаний: {len(warnings)}')
            if warnings:messagebox.showwarning('Замечания при объединении','\n'.join(warnings[:25]))
        self.root.after(100,poll)
    def reset_wb_search(self):
        self.wb_search_entry.delete(0,tk.END);self.render_wb()
    def render_wb(self):
        for tab in self.wb_tabs.winfo_children():tab.destroy()
        query=self.wb_search_entry.get().strip().lower();self.wb_export_views={}
        for title,df in self.wb_views.items():
            if query and not df.empty:
                cols=[c for c in ('Артикул WB','Артикул продавца','Наименование','Название','Поисковый запрос') if c in df]
                if cols:df=df[df[cols].fillna('').astype(str).apply(lambda s:s.str.lower().str.contains(query,regex=False)).any(axis=1)]
            self.wb_export_views[title]=df
            frame=ttk.Frame(self.wb_tabs);self.wb_tabs.add(frame,text=f'{title} ({len(df)})')
            tree=ttk.Treeview(frame,columns=list(range(len(df.columns))),show='headings')
            y=ttk.Scrollbar(frame,orient='vertical',command=tree.yview);x=ttk.Scrollbar(frame,orient='horizontal',command=tree.xview)
            tree.configure(yscrollcommand=y.set,xscrollcommand=x.set)
            tree.grid(row=0,column=0,sticky='nsew');y.grid(row=0,column=1,sticky='ns');x.grid(row=1,column=0,sticky='ew');frame.rowconfigure(0,weight=1);frame.columnconfigure(0,weight=1)
            for i,c in enumerate(df.columns):tree.heading(i,text=c);tree.column(i,width=220 if '|' in c else 170,minwidth=70,stretch=False)
            tree.tag_configure('even',background='#edf3f8');tree.tag_configure('odd',background='#ffffff')
            previous=None;block=0
            for values in df.itertuples(index=False,name=None):
                period=values[0] if values else ''
                if period!=previous:block+=1;previous=period
                tree.insert('', 'end',values=['' if pd.isna(v) else f'{v:g}' if isinstance(v,float) else str(v) for v in values],tags=('even' if block%2==0 else 'odd',))
    def export_wb_csv(self):
        if not self.wb_views:return
        i=self.wb_tabs.index(self.wb_tabs.select());title=list(self.wb_export_views)[i]
        path=filedialog.asksaveasfilename(initialfile=title+'.csv',defaultextension='.csv',filetypes=[('CSV для Excel','*.csv')])
        if path:
            try:self.wb_export_views[title].to_csv(path,index=False,sep=';',encoding='utf-8-sig');messagebox.showinfo('Экспорт','Файл сохранён')
            except Exception as exc:messagebox.showerror('Экспорт',str(exc))


if __name__ == "__main__":
    root = tk.Tk()
    app = MarketplaceViewer(root)
    root.mainloop()
