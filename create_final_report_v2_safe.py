"""
Этап 3 V2 SAFE: Откат + защитные механизмы + лист unmatched.
Внедрены все замечания из code review.
"""

import pandas as pd
import json
import sys
from pathlib import Path
from rapidfuzz import process, fuzz
from collections import defaultdict

# === КОНФИГУРАЦИЯ ===
CONFIG = {
    'excel_input': 'Сверка_спецификаций.xlsx',
    'pdf_json': 'pdf_all_parsed.json',
    'output_file': 'Сверка_результат.xlsx',  # ОТДЕЛЬНЫЙ выходной файл!
    'fuzzy_threshold': 90,
    'sheet_excel': 'Сводная по КП',
    'sheet_result': 'Сверка с PDF',
    'sheet_unmatched': 'Не найдено в Excel',
}


def load_excel_summary(filepath: Path) -> pd.DataFrame:
    """Безопасная загрузка Excel с проверкой дубликатов."""
    try:
        df = pd.read_excel(filepath, sheet_name=CONFIG['sheet_excel'], engine='openpyxl')
        print(f"[INFO] Загружено {len(df)} позиций из Excel")
    except FileNotFoundError:
        print(f"[ERROR] Файл не найден: {filepath}")
        sys.exit(1)
    except Exception as e:
        print(f"[ERROR] Ошибка чтения Excel: {e}")
        sys.exit(1)
    
    # Проверка на дубликаты названий
    names_lower = df['name'].str.lower().str.strip()
    dup_count = len(names_lower) - len(set(names_lower))
    if dup_count > 0:
        print(f"[WARNING] Обнаружены дубликаты названий: {dup_count} шт.")
        print("[WARNING] При матчинге количество PDF может 'прилипнуть' только к последней строке!")
    
    return df


def load_pdf_data(filepath: Path) -> list:
    """Безопасная загрузка JSON."""
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            data = json.load(f)
        print(f"[INFO] Загружено {len(data)} позиций из PDF")
        return data
    except FileNotFoundError:
        print(f"[ERROR] Файл не найден: {filepath}")
        sys.exit(1)
    except json.JSONDecodeError as e:
        print(f"[ERROR] Битый JSON: {e}")
        sys.exit(1)


def match_pdf_to_excel(excel_df: pd.DataFrame, pdf_items: list) -> tuple:
    """
    PDF → Excel. Для каждой позиции PDF ищем ОДНО лучшее совпадение.
    Tie-breaker: если fuzzy score равен, приоритет у совпадения по артикулу.
    Возвращает: (matches_dict, unmatched_list)
    """
    print(f"\n[INFO] Fuzzy matching: PDF → Excel (порог {CONFIG['fuzzy_threshold']}%)...")
    
    excel_names = [row['name'].lower().strip() for _, row in excel_df.iterrows()]
    excel_articles = [str(row.get('article_raw', '')).lower().strip() for _, row in excel_df.iterrows()]
    
    matches = defaultdict(lambda: {'qty': 0, 'count': 0})
    unmatched = []
    
    matched_count = 0
    
    for item in pdf_items:
        pdf_name = item['name'].lower().strip()
        pdf_article = str(item.get('article', '')).lower().strip()
        pdf_qty = item['qty']
        
        # Ищем лучшее совпадение
        result = process.extractOne(
            pdf_name,
            excel_names,
            scorer=fuzz.WRatio,
            score_cutoff=CONFIG['fuzzy_threshold']
        )
        
        if result:
            best_score = result[1]
            best_idx = result[2]
            
            # TIE-BREAKER: Если есть другие кандидаты с таким же score,
            # проверяем, есть ли среди них совпадение по артикулу
            if pdf_article:
                candidates = process.extract(
                    pdf_name,
                    excel_names,
                    scorer=fuzz.WRatio,
                    score_cutoff=best_score,
                    limit=5
                )
                for cand_score, cand_idx, _ in candidates:
                    if cand_score == best_score and pdf_article in excel_articles[cand_idx]:
                        best_idx = cand_idx
                        break
            
            excel_name_key = excel_df.iloc[best_idx]['name'].lower().strip()
            matches[excel_name_key]['qty'] += pdf_qty
            matches[excel_name_key]['count'] += 1
            matched_count += 1
        else:
            unmatched.append({
                'Наименование PDF': item['name'],
                'Артикул PDF': item.get('article', ''),
                'Количество': pdf_qty,
                'Страница': item.get('page', ''),
                'Причина': f'Fuzzy score < {CONFIG["fuzzy_threshold"]}%'
            })
    
    print(f"[INFO] PDF позиций: совпало {matched_count}, не совпало {len(unmatched)}")
    
    # Преобразуем matches в формат для отчета
    report_matches = {}
    for idx, row in excel_df.iterrows():
        name_key = row['name'].lower().strip()
        report_matches[name_key] = {
            'pdf_qty': matches[name_key]['qty'],
            'match_count': matches[name_key]['count']
        }
    
    return report_matches, unmatched


def color_discrepancy(val):
    """Цветовая кодировка расхождений."""
    if val > 0:
        return 'background-color: #ffcccc'  # Красный: КП > PDF
    elif val < 0:
        return 'background-color: #ccffcc'  # Зеленый: PDF > КП
    return ''


def create_final_excel(excel_df: pd.DataFrame, matches: dict, 
                       unmatched: list, output_path: Path):
    """Создание Excel с тремя листами: сводка, сверка, unmatched."""
    print(f"\n[INFO] Создаю финальный Excel: {output_path}...")
    
    # Лист 1: Сверка с PDF
    comparison_data = []
    for idx, row in excel_df.iterrows():
        excel_name = row['name'].lower().strip()
        excel_qty = row['qty']
        
        m = matches.get(excel_name, {'pdf_qty': 0, 'match_count': 0})
        
        comparison_data.append({
            'Наименование': row['name'],
            'Система': row.get('system', ''),
            'Кол-во КП': excel_qty,
            'Кол-во PDF': m['pdf_qty'],
            'Разница': excel_qty - m['pdf_qty'],
            'Найдено вхождений': m['match_count']
        })
    
    df_comparison = pd.DataFrame(comparison_data)
    df_comparison = df_comparison.sort_values('Разница', ascending=False)
    
    # Лист 2: Не найдено в Excel
    df_unmatched = pd.DataFrame(unmatched) if unmatched else pd.DataFrame(
        columns=['Наименование PDF', 'Артикул PDF', 'Количество', 'Страница', 'Причина']
    )
    
    # Запись в Excel
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Лист сверки со стилизацией
            styled = df_comparison.style.map(color_discrepancy, subset=['Разница'])
            styled.to_excel(writer, index=False, sheet_name=CONFIG['sheet_result'])
            
            ws = writer.sheets[CONFIG['sheet_result']]
            ws.auto_filter.ref = ws.dimensions
            for i, col in enumerate(df_comparison.columns):
                max_len = max(df_comparison[col].astype(str).map(len).max(), len(col)) + 2
                ws.column_dimensions[ws.cell(row=1, column=i+1).column_letter].width = min(max_len, 60)
            
            # Лист unmatched
            df_unmatched.to_excel(writer, index=False, sheet_name=CONFIG['sheet_unmatched'])
            ws2 = writer.sheets[CONFIG['sheet_unmatched']]
            ws2.auto_filter.ref = ws2.dimensions
            
        print(f"[OK] Сохранено в {output_path}")
    except PermissionError:
        print(f"[ERROR] Файл {output_path} открыт в Excel! Закройте его и повторите.")
        sys.exit(1)
    
    # Статистика
    print(f"\n[СТАТИСТИКА]")
    print(f"  Всего позиций КП: {len(df_comparison)}")
    exact = len(df_comparison[df_comparison['Разница'] == 0])
    diff = len(df_comparison[df_comparison['Разница'] != 0])
    zero_pdf = len(df_comparison[df_comparison['Кол-во PDF'] == 0])
    print(f"  Совпало точно: {exact}")
    print(f"  С расхождениями: {diff}")
    print(f"  Не найдено в PDF: {zero_pdf}")
    print(f"  Позиций PDF без совпадения: {len(unmatched)}")


def main():
    print(f"{'='*60}")
    print("ЭТАП 3 V2 SAFE: Защитный матчинг + лист unmatched")
    print(f"{'='*60}")
    
    excel_df = load_excel_summary(Path(CONFIG['excel_input']))
    pdf_items = load_pdf_data(Path(CONFIG['pdf_json']))
    
    matches, unmatched = match_pdf_to_excel(excel_df, pdf_items)
    
    create_final_excel(excel_df, matches, unmatched, Path(CONFIG['output_file']))
    
    print(f"\n{'='*60}")
    print("ГОТОВО!")
    print(f"{'='*60}")
    print(f"Откройте: explorer.exe {CONFIG['output_file']}")


if __name__ == "__main__":
    main()
