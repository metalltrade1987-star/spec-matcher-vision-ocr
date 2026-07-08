"""
Этап 1: Парсинг Excel и создание первых двух таблиц.
"""

import pandas as pd
import re
from pathlib import Path

EXCEL_FILE = Path("КП 1пров.xls")
OUTPUT_FILE = Path("Сверка_спецификаций.xlsx")


def parse_excel(filepath: Path) -> pd.DataFrame:
    """Парсит Excel, извлекает систему, наименование, количество."""
    print(f"[INFO] Читаю Excel: {filepath}")
    
    df_raw = pd.read_excel(filepath, engine='xlrd', header=None)
    print(f"[INFO] Всего строк в Excel: {len(df_raw)}")
    
    clean_rows = []
    current_system = "Не определено"
    
    # Паттерн для определения системы (П1, В2, М3, ККБ, Автоматика и т.д.)
    system_pattern = re.compile(r'^(П\d|В\d|М\d|ККБ|Автоматика)', re.IGNORECASE)
    
    # Слова-маркеры заголовков/пропускаемых строк
    skip_words = [
        "номенклатура", "артикул", "код обору", "ед. изм",
        "кол-во", "цена", "сумма", "валюта", "итоги",
        "ндс", "стоимость", "предложение", "р-климат",
        "наименование", "техническая характеристика"
    ]
    
    for idx, row in df_raw.iterrows():
        # Получаем все непустые ячейки строки
        non_empty_cells = [str(cell).strip() for cell in row if pd.notna(cell) and str(cell).strip()]
        if not non_empty_cells:
            continue
        
        full_row_text = " ".join(non_empty_cells).lower()
        
        # Пропускаем строки-заголовки
        if any(w in full_row_text for w in skip_words):
            continue
        
        # Проверяем, является ли строка заголовком системы
        is_system_row = False
        for cell_text in non_empty_cells:
            if system_pattern.match(cell_text):
                # Извлекаем название системы (до слова "Расход" или конца)
                current_system = cell_text.split('Расход')[0].strip()
                is_system_row = True
                break
        
        if is_system_row:
            continue
        
        # Ищем количество (число от 1 до 500)
        qty_found = None
        name_found = ""
        art_found = ""
        
        for cell in row:
            val = str(cell).strip() if pd.notna(cell) else ""
            if not val or val == "nan":
                continue
            
            # Пытаемся найти количество
            clean_val = val.replace(',', '.').replace(' ', '')
            try:
                num = float(clean_val)
                if '.' not in clean_val and 1 <= num <= 500:
                    qty_found = num
            except ValueError:
                pass
            
            # Ищем наименование (длинная строка с кириллицей)
            if len(val) > 10 and re.search(r'[а-яА-Я]', val) and not name_found:
                name_found = val
            # Ищем артикул (латиница + цифры, короткий)
            elif (re.match(r'^[A-Za-z0-9\-\.\/\(\)]+$', val)
                  and len(val) > 3 and not val.isdigit() and not art_found):
                art_found = val
        
        # Валидация: должно быть наименование и количество
        if not name_found or qty_found is None or len(name_found) < 5:
            continue
        
        clean_rows.append({
            'system': current_system,
            'name': name_found,
            'article_raw': art_found,
            'qty': int(qty_found)
        })
    
    print(f"[INFO] Извлечено позиций: {len(clean_rows)}")
    print(f"[INFO] Системы: {sorted(set(r['system'] for r in clean_rows))}")
    
    return pd.DataFrame(clean_rows)


def create_summary_table(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Группирует одинаковые позиции, суммирует количества."""
    print(f"\n[INFO] Создаю сводную таблицу (группировка)...")
    
    # Нормализуем названия для группировки
    df = df_raw.copy()
    df['name_norm'] = df['name'].str.lower().str.strip()
    
    # Группируем по нормализованному имени
    grouped = df.groupby('name_norm').agg({
        'name': 'first',           # Берём первое (оригинальное) название
        'article_raw': 'first',    # Берём первый артикул
        'qty': 'sum',              # Суммируем количества
        'system': lambda x: ', '.join(sorted(set(x)))  # Объединяем системы
    }).reset_index()
    
    # Сортируем по количеству (убывание)
    grouped = grouped.sort_values('qty', ascending=False)
    
    print(f"[INFO] Уникальных позиций после группировки: {len(grouped)}")
    
    return grouped[['system', 'name', 'article_raw', 'qty']]


def save_to_excel(df_raw: pd.DataFrame, df_summary: pd.DataFrame, output_path: Path):
    """Сохраняет две таблицы в Excel."""
    print(f"\n[INFO] Сохраняю в {output_path}...")
    
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        # Лист 1: Сырые данные
        df_raw.to_excel(writer, index=False, sheet_name='Сырые данные КП')
        ws1 = writer.sheets['Сырые данные КП']
        ws1.autofilter(0, 0, len(df_raw), len(df_raw.columns) - 1)
        for i, col in enumerate(df_raw.columns):
            max_len = max(df_raw[col].astype(str).map(len).max(), len(col)) + 2
            ws1.set_column(i, i, min(max_len, 50))
        
        # Лист 2: Сводная по КП
        df_summary.to_excel(writer, index=False, sheet_name='Сводная по КП')
        ws2 = writer.sheets['Сводная по КП']
        ws2.autofilter(0, 0, len(df_summary), len(df_summary.columns) - 1)
        for i, col in enumerate(df_summary.columns):
            max_len = max(df_summary[col].astype(str).map(len).max(), len(col)) + 2
            ws2.set_column(i, i, min(max_len, 50))
    
    print(f"[OK] Сохранено: {output_path}")


def main():
    print("=" * 60)
    print("ЭТАП 1: Парсинг Excel")
    print("=" * 60)
    
    # 1. Парсим Excel
    df_raw = parse_excel(EXCEL_FILE)
    
    if df_raw.empty:
        print("[ERROR] Не удалось извлечь данные из Excel!")
        return
    
    # 2. Создаём сводную таблицу
    df_summary = create_summary_table(df_raw)
    
    # 3. Сохраняем
    save_to_excel(df_raw, df_summary, OUTPUT_FILE)
    
    print(f"\n{'=' * 60}")
    print("ГОТОВО!")
    print(f"{'=' * 60}")
    print(f"Лист 1 'Сырые данные КП': {len(df_raw)} строк")
    print(f"Лист 2 'Сводная по КП': {len(df_summary)} строк")
    print(f"\nОткрой файл: explorer.exe {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
