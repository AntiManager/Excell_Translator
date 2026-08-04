# Excell_Translator — Tech Context
> Updated: 2026-08-04

## Запуск
```bash
pip install -r requirements.txt
python main.py              # GUI
python run_translate.py     # консольный запуск
python check_progress.py    # статус
python check_cache.py       # кэш
```

## Стек
- pandas, openpyxl, deep-translator, requests, tqdm
- deep-translator вместо googletrans (нестабилен)

## Ключевые файлы
- `translator.py` — ядро перевода
- `batch_translate.py` — батч-обработка
- `translation_state.json` — состояние сессии (инкрементальный перевод)
- `eclass_16.xlsx`, `eclass_16_ru.xlsx` — данные ECLASS

## Лимиты
- RateLimit: 80 запросов/минуту
- Экспоненциальная задержка при ошибках сети
