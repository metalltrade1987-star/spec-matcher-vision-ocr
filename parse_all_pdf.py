"""
Этап 2: Парсинг всех страниц PDF через vision-модель.
"""

import base64
import json
import requests
import re
from pathlib import Path
from pdf2image import convert_from_path
import io
import time

PDF_PATH = Path("project_spec.pdf")
OLLAMA_URL = "http://localhost:11434/api/generate"
MODEL_NAME = "qwen2.5vl:7b"

# Все страницы с таблицами (найденные ранее)
TABLE_PAGES = [73, 80, 85, 90, 95, 100, 105, 110, 114, 116, 118, 122, 124, 
               152, 153, 154, 155, 156, 157, 158, 159, 160, 161, 162, 163, 
               164, 165, 166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 
               176, 177, 178, 179, 180, 181, 182, 183, 184, 185]


def extract_json(text):
    """Извлекает JSON-массив из текста ответа."""
    start = text.find("[")
    end = text.rfind("]")
    if start != -1 and end != -1 and end > start:
        return text[start:end+1]
    return ""


def fix_cyrillic_latin(text):
    """Заменяет латинские буквы на кириллицу в русских словах"""
    replacements = {
        'p': 'р', 'y': 'у', 'o': 'о', 'c': 'с', 'e': 'е',
        'x': 'х', 'M': 'М', 'T': 'Т', 'A': 'А', 'K': 'К',
        'H': 'Н', 'P': 'Р', 'B': 'В', 'E': 'Е', '6': 'б',
        'r': 'г', 'a': 'а', 'u': 'и'
    }
    words = text.split()
    fixed_words = []
    for w in words:
        if re.search(r'[а-яА-Я]', w):
            new_w = ''.join(replacements.get(ch, ch) for ch in w)
            fixed_words.append(new_w)
        else:
            fixed_words.append(w)
    return ' '.join(fixed_words)


def parse_page(page_num, dpi=250):
    """Парсит одну страницу PDF через vision-модель."""
    print(f"\n[INFO] Страница {page_num}...", end=" ", flush=True)
    
    try:
        images = convert_from_path(PDF_PATH, first_page=page_num, last_page=page_num, dpi=dpi)
        img = images[0]
        
        buffer = io.BytesIO()
        img.save(buffer, format="PNG")
        img_b64 = base64.b64encode(buffer.getvalue()).decode("utf-8")
        
        prompt = f"""Ты — система оптического распознавания символов (OCR). Твоя задача — СКОПИРОВАТЬ текст из таблицы на странице {page_num}.

КРИТИЧЕСКИ ВАЖНЫЕ ПРАВИЛА:
1. КОПИРУЙ БУКВА В БУКВУ. НЕ ИСПРАВЛЯЙ слова. НЕ ДОДУМЫВАЙ смысл.
2. ЗАПРЕЩЕНО заменять слова на синонимы.
3. ЕДИНИЦЫ ИЗМЕРЕНИЯ: Внимательно смотри на колонку "Единица измерения". Там может быть "м.п.", "м", "шт", "компл". Используй ТОЧНО то, что написано.
4. КОЛИЧЕСТВО: Если число дробное (например 37,333), округли до целого. Если ячейка пустая — пропусти эту строку.
5. ЗАГОЛОВКИ: Пропускай строки, где в колонке "Количество" нет числа.

ФОРМАТ ОТВЕТА (строго JSON):
[{{"name": "точный текст из колонки Наименование", "article": "текст из колонки Тип/марка", "qty": 10, "unit": "м.п."}}]

Начинай ответ сразу с "["."""

        payload = {
            "model": MODEL_NAME,
            "prompt": prompt,
            "images": [img_b64],
            "stream": False,
            "options": {
                "temperature": 0.0,
                "num_predict": 4096,
                "num_ctx": 32768
            }
        }
        
        response = requests.post(OLLAMA_URL, json=payload, timeout=300)
        
        if response.status_code != 200:
            print(f"ERROR {response.status_code}")
            return []
        
        result = response.json()
        answer = result.get("response", "")
        eval_duration = result.get("eval_duration", 0) / 1e9
        
        json_str = extract_json(answer)
        if not json_str:
            print(f"WARN (нет JSON)")
            return []
            
        items = json.loads(json_str)
        
        # Постобработка
        clean_items = []
        for item in items:
            name = item.get("name", "").strip()
            qty_raw = item.get("qty", 0)
            unit = item.get("unit", "шт").strip()
            
            try:
                qty = int(float(str(qty_raw).replace(',', '.')))
                if not (1 <= qty <= 500):
                    continue
            except:
                continue
                
            name = fix_cyrillic_latin(name)
            article = fix_cyrillic_latin(item.get("article", ""))
            
            clean_items.append({
                "name": name,
                "article": article,
                "qty": qty,
                "unit": unit,
                "page": page_num
            })
        
        print(f"OK ({len(clean_items)} поз, {eval_duration:.1f}с)")
        return clean_items
        
    except Exception as e:
        print(f"ERROR: {e}")
        return []


def main():
    print(f"{'='*60}")
    print(f"ПАРСИНГ PDF: {len(TABLE_PAGES)} страниц")
    print(f"{'='*60}")
    
    all_items = []
    start_time = time.time()
    
    for i, page_num in enumerate(TABLE_PAGES, 1):
        items = parse_page(page_num)
        all_items.extend(items)
        
        # Прогресс
        if i % 10 == 0:
            elapsed = time.time() - start_time
            print(f"\n[ПРОГРЕСС] {i}/{len(TABLE_PAGES)} страниц, {elapsed:.0f} сек")
    
    elapsed = time.time() - start_time
    print(f"\n{'='*60}")
    print(f"ИТОГО: {len(all_items)} позиций из {len(TABLE_PAGES)} страниц за {elapsed:.0f} сек")
    print(f"{'='*60}")
    
    # Сохраняем
    output_file = Path("pdf_all_parsed.json")
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(all_items, f, ensure_ascii=False, indent=2)
    
    print(f"[OK] Сохранено в {output_file}")
    
    # Показываем примеры
    print(f"\n[DEBUG] Первые 10 позиций:")
    for item in all_items[:10]:
        print(f"  {item['qty']:3d} {item['unit']:5s} | {item['name'][:50]}")


if __name__ == "__main__":
    main()
