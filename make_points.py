import json
import os
import openpyxl
import random

current_dir = os.path.dirname(os.path.abspath(__file__))
input_file = os.path.join(current_dir, "Лист Microsoft Excel.xlsx")
output_file = os.path.join(current_dir, "points.js")

points_data = []

try:
    wb = openpyxl.load_workbook(input_file, data_only=True)
    sheet = wb.active

    for row_idx, row in enumerate(sheet.iter_rows(values_only=True)):
        if len(row) < 4 or row[2] is None or row[3] is None:
            continue
            
        point_id = str(row[0]) if row[0] is not None else str(row_idx + 1)
        address = str(row[1]) if row[1] is not None else "Адрес неизвестен"
        
        try:
            lat = float(str(row[2]).replace(',', '.')) + random.uniform(-0.00015, 0.00015)
            lon = float(str(row[3]).replace(',', '.')) + random.uniform(-0.00015, 0.00015)
        except ValueError:
            continue 

        points_data.append({
            "id": point_id,
            "name": f"Точка {point_id}",
            "address": address,
            "coords": [lat, lon]
        })

    # Сжимаем всё в одну плотную строку
    js_content = f"const pointsData = {json.dumps(points_data, ensure_ascii=False)};"

    with open(output_file, mode='w', encoding='utf-8') as f:
        f.write(js_content)

    print(f"Готово! Обработано точек: {len(points_data)}")

except FileNotFoundError:
    print("ОШИБКА: Файл 'Лист Microsoft Excel.xlsx' не найден!")
except Exception as e:
    print(f"Произошла ошибка: {e}")