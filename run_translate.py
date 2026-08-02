#!/usr/bin/env python3
import logging
import sys
from pathlib import Path
sys.path.append(Path(__file__).parent)
from translator import ExcelTranslator

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    stream=sys.stdout,
    force=True
)

file_path = r"C:\Users\evgeniy.bogdanov\Documents\Python\Excell_Translator\eclass_16.xlsx"
input_path = Path(file_path)
output_path = input_path.parent / f"{input_path.stem}_ru{input_path.suffix}"

translator = ExcelTranslator()

selected_sheets = {
    "ECLASS 16.0": ["Название EN", "Определение"]
}

def progress_callback(sheet_name, progress, message):
    print(f"[{progress:.1f}%] {message}", flush=True)

try:
    success = translator.process_excel_file(
        str(file_path),
        str(output_path),
        selected_sheets,
        progress_callback=progress_callback
    )
    if success:
        stats = translator.get_translation_stats()
        print("\n=== ИТОГОВАЯ СТАТИСТИКА ===", flush=True)
        print(f"Переведено: {stats['translated']}", flush=True)
        print(f"Из кэша: {stats['cached']}", flush=True)
        print(f"Ошибки: {stats['failed']}", flush=True)
        print(f"Повторы: {stats['retries']}", flush=True)
        print(f"Время: {stats['elapsed_time']:.1f} сек", flush=True)
        print(f"Скорость: {stats['requests_per_second']:.1f} запр/сек", flush=True)
    else:
        print("Перевод завершился с ошибкой", flush=True)
except Exception as e:
    print(f"КРИТИЧЕСКАЯ ОШИБКА: {e}", flush=True)
    import traceback
    traceback.print_exc()
