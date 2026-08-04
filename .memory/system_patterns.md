# Excell_Translator — System Patterns
> Updated: 2026-08-04

## Architecture
```
Excel → translator.py (deep-translator) → Excel + _ru
Кэширование переводов, RateLimiter (80/мин), экспоненциальные повторы
translation_state.json — инкрементальный перевод (продолжение с остановки)
Фильтрация: числа/даты/спецсимволы пропускаются; уникальные переводы
```

## Conventions
- deep-translator (не googletrans)
- Батч-обработка + уникальные переводы для оптимизации
- Разделение сетевых и API-ошибок, автоматическое восстановление
