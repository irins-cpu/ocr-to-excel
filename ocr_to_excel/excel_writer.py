# ocr_to_excel/excel_writer.py

import pandas as pd
import os
from datetime import datetime

def save_to_excel(data_dict, output_folder='output'):
    """
    Сохраняет распознанный текст в Excel.
    Каждый файл -> отдельная строка, весь текст в одной колонке.
    """
    rows = []
    for filename, lines in data_dict.items():
        row = {
            'Файл': filename,
            'Текст (распознанный)': '\n'.join(lines)
        }
        rows.append(row)

    df = pd.DataFrame(rows)

    # Убедимся, что папка существует
    os.makedirs(output_folder, exist_ok=True)

    # Имя файла с текущей датой и временем
    now = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    output_path = os.path.join(output_folder, f'ocr_result_{now}.xlsx')

    df.to_excel(output_path, index=False)
    return output_path
