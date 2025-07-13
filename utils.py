import pandas as pd
import time
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Font

def find_column_index(columns, names):
    """Ищет в списке columns индекс колонки с любым из имён в names."""
    lower_names = [n.lower() for n in names]
    for idx, col in enumerate(columns):
        if str(col).strip().lower() in lower_names:
            return idx
    return None

def find_quantity_column(df):
    """Ищет первую из первых 10 колонок, чьё имя похоже на 'кол-во' и содержит числа."""
    candidates = ['кол-во', 'количество', 'кол']
    for idx, col in enumerate(df.columns[:10]):
        low = str(col).strip().lower()
        if low in candidates:
            sample = df[col].dropna().head(10)
            if sample.apply(pd.to_numeric, errors='coerce').notna().any():
                return idx
    return None

def save_workbook_with_retries(wb, filename, retries=5, delay=3):
    """Пытается сохранить книгу, если файл занят — ждёт и повторяет."""
    for attempt in range(1, retries + 1):
        try:
            wb.save(filename)
            return True
        except PermissionError:
            time.sleep(delay)
    return False

def format_sticker_cell(cell):
    """Отформатировать ячейку '№ Стикера': последние 4 символа — жирные + размер шрифта +1."""
    value = str(cell.value or "")
    if len(value) < 4:
        return
    main, last4 = value[:-4].rstrip(), value[-4:]
    cell.value = f"{main} {last4}"
    cell.font = cell.font.copy(bold=True, size=cell.font.size + 1)
