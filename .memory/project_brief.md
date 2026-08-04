# Excell_Translator — Project Brief
> Updated: 2026-08-04

## Overview
Excel Translator v3.0 — перевод Excel-файлов на русский язык с графическим интерфейсом. Стабильный переводчик на базе deep-translator, умное кэширование, контроль лимитов.

## Stack
- **Language**: Python (pandas, openpyxl, deep-translator, requests, tqdm)
- **UI**: GUI (PyInstaller .exe)
- **Entry**: `main.py`

## Structure
- `main.py`, `translator.py`, `batch_translate.py`, `run_translate.py`
- Утилиты: calc_remaining.py, check_cache.py, check_progress.py, debug.py, test_batch.py
- `eclass_16.xlsx`, `eclass_16_ru.xlsx` — данные ECLASS
- `translation_state.json` — состояние сессии

## Constraints
- Перевод через deep-translator (НЕ googletrans — нестабилен)
- Rate limit: 80 запросов/минуту, экспоненциальные повторы при ошибках
- Инкрементальный перевод (продолжение с места остановки)
- Фильтрация: пропуск чисел, дат, спецсимволов; уникальные переводы

## External Dependencies
- ECLASS 16.0 данные (eclass_16*.xlsx)

## Переиспользование
- Модуль перевода, работа с openpyxl
