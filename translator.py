#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Улучшенный модуль для перевода Excel-файлов с поддержкой инкрементального перевода и управления состоянием.
Использует стабильный deep-translator вместо googletrans.
"""
import time
import pandas as pd
from deep_translator import GoogleTranslator
import logging
import json
import os
from pathlib import Path
from typing import Dict, List, Set, Any, Optional, Callable
from datetime import datetime
import random
from requests.exceptions import RequestException
import re


class RateLimiter:
    """Класс для контроля частоты запросов к API перевода"""
    
    def __init__(self, max_requests_per_minute: int = 100):
        self.max_requests = max_requests_per_minute
        self.requests = []
    
    def wait_if_needed(self):
        """Ожидает если превышен лимит запросов в минуту"""
        now = time.time()
        # Удаляем запросы старше 1 минуты
        self.requests = [req_time for req_time in self.requests if now - req_time < 60]
        
        if len(self.requests) >= self.max_requests:
            sleep_time = 60 - (now - self.requests[0])
            if sleep_time > 0:
                logging.info(f"Достигнут лимит запросов. Ожидание {sleep_time:.1f} секунд")
                time.sleep(sleep_time)
        
        self.requests.append(time.time())


class TranslationStateManager:
    """Менеджер состояния перевода для сохранения и восстановления прогресса"""
    
    def __init__(self, state_file: str = "translation_state.json"):
        self.state_file = state_file
        self.state = self._load_initial_state()
    
    def _load_initial_state(self) -> Dict:
        """Загружает начальное состояние из файла"""
        try:
            if os.path.exists(self.state_file):
                with open(self.state_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            logging.error(f"Ошибка загрузки состояния: {e}")
        
        return {
            'version': '2.0',
            'created_at': datetime.now().isoformat(),
            'last_updated': datetime.now().isoformat(),
            'file_path': '',
            'selected_sheets': {},
            'completed_sheets': {},
            'sheet_progress': {},
            'translation_cache': {},
            'failed_translations': {}
        }
    
    def save_state(self):
        """Сохраняет текущее состояние в файл"""
        try:
            self.state['last_updated'] = datetime.now().isoformat()
            with open(self.state_file, 'w', encoding='utf-8') as f:
                json.dump(self.state, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.error(f"Ошибка сохранения состояния: {e}")
    
    def load_state(self) -> Dict:
        """Загружает состояние из файла"""
        return self.state
    
    def update_state(self, updates: Dict):
        """Обновляет состояние"""
        self.state.update(updates)
        self.save_state()
    
    def clear_state(self):
        """Очищает состояние"""
        self.state = {
            'version': '2.0',
            'created_at': datetime.now().isoformat(),
            'last_updated': datetime.now().isoformat(),
            'file_path': '',
            'selected_sheets': {},
            'completed_sheets': {},
            'sheet_progress': {},
            'translation_cache': {},
            'failed_translations': {}
        }
        self.save_state()
    
    def mark_sheet_completed(self, sheet_name: str):
        """Отмечает лист как завершенный"""
        self.state['completed_sheets'][sheet_name] = datetime.now().isoformat()
        self.state['sheet_progress'][sheet_name] = 100.0
        self.save_state()
    
    def update_sheet_progress(self, sheet_name: str, progress: float):
        """Обновляет прогресс листа"""
        self.state['sheet_progress'][sheet_name] = progress
        self.save_state()
    
    def get_sheet_progress(self, sheet_name: str) -> str:
        """Возвращает прогресс листа в формате строки"""
        progress = self.state['sheet_progress'].get(sheet_name, 0)
        return f"{progress:.1f}%" if progress > 0 else ""
    
    def is_sheet_completed(self, sheet_name: str) -> bool:
        """Проверяет, завершен ли лист"""
        return sheet_name in self.state['completed_sheets']
    
    def get_completed_sheets(self) -> List[str]:
        """Возвращает список завершенных листов"""
        return list(self.state['completed_sheets'].keys())
    
    def add_to_cache(self, original: str, translated: str):
        """Добавляет перевод в кэш"""
        self.state['translation_cache'][original] = translated
        # Сохраняем только каждые 50 записей для производительности
        if len(self.state['translation_cache']) % 50 == 0:
            self.save_state()
    
    def get_from_cache(self, original: str) -> Optional[str]:
        """Получает перевод из кэша"""
        return self.state['translation_cache'].get(original)
    
    def mark_failed_translation(self, original: str, error: str):
        """Отмечает неудачный перевод"""
        self.state['failed_translations'][original] = {
            'error': error,
            'timestamp': datetime.now().isoformat()
        }
    
    def get_failed_count(self) -> int:
        """Возвращает количество неудачных переводов"""
        return len(self.state['failed_translations'])


class ExcelTranslator:
    def __init__(self, delay: float = 0.1, max_retries: int = 5, batch_size: int = 50):
        self.translator = GoogleTranslator(source='auto', target='ru')
        self.rate_limiter = RateLimiter(max_requests_per_minute=80)  # Безопасный лимит
        self.delay = delay
        self.max_retries = max_retries
        self.batch_size = batch_size
        self.state_manager = TranslationStateManager()
        self.session_start_time = time.time()
        self.total_requests = 0
        
        # Статистика
        self.stats = {
            'translated': 0,
            'cached': 0,
            'failed': 0,
            'retries': 0
        }
    
    def _should_translate(self, text: str) -> bool:
        """Проверяет, нужно ли переводить текст"""
        if not isinstance(text, str):
            return False
        
        text = text.strip()
        if not text:
            return False
        
        # Не переводим числа
        if text.replace('.', '').replace(',', '').isdigit():
            return False
        
        # Не переводим даты в формате YYYY-MM-DD
        if re.match(r'^\d{4}-\d{2}-\d{2}$', text):
            return False
        
        # Не переводим слишком короткие тексты (меньше 2 символов)
        if len(text) < 2:
            return False
        
        # Проверяем, не состоит ли текст в основном из специальных символов
        alpha_count = sum(1 for char in text if char.isalpha())
        if alpha_count / len(text) < 0.3:  # Меньше 30% букв
            return False
        
        return True
    
    def translate_text_with_retry(self, text: str) -> str:
        """Переводит текст с повторными попытками при ошибках."""
        if not self._should_translate(text):
            return text
        
        # Проверяем кэш состояния
        cached = self.state_manager.get_from_cache(text)
        if cached:
            self.stats['cached'] += 1
            return cached
        
        last_exception = None
        
        for attempt in range(self.max_retries):
            try:
                # Контроль частоты запросов
                self.rate_limiter.wait_if_needed()
                
                # Случайная задержка для избежания паттернов
                time.sleep(self.delay + random.uniform(0, 0.2))
                
                translated = self.translator.translate(text)
                self.total_requests += 1
                
                # Проверяем валидность перевода
                if translated and isinstance(translated, str) and translated.strip():
                    self.state_manager.add_to_cache(text, translated)
                    self.stats['translated'] += 1
                    
                    # Логируем каждые 50 запросов
                    if self.total_requests % 50 == 0:
                        elapsed = time.time() - self.session_start_time
                        rate = self.total_requests / elapsed if elapsed > 0 else 0
                        logging.info(f"Переведено запросов: {self.total_requests}, "
                                   f"скорость: {rate:.1f} запр/сек, "
                                   f"кэш: {self.stats['cached']}, "
                                   f"ошибки: {self.stats['failed']}")
                    
                    return translated
                else:
                    raise ValueError("Пустой или некорректный перевод")
                    
            except RequestException as e:
                last_exception = e
                wait_time = (2 ** attempt) + random.uniform(0, 1)  # Экспоненциальная задержка
                logging.warning(f"Сетевая ошибка (попытка {attempt + 1}/{self.max_retries}): {e}")
                time.sleep(wait_time)
                self.stats['retries'] += 1
                
            except Exception as e:
                last_exception = e
                if "429" in str(e):  # Too Many Requests
                    wait_time = 30 + random.uniform(0, 10)  # Длительная пауза при лимите
                    logging.warning(f"Превышен лимит запросов. Ожидание {wait_time:.1f} секунд")
                    time.sleep(wait_time)
                else:
                    wait_time = (2 ** attempt) * 0.5 + random.uniform(0, 0.5)
                    logging.warning(f"Ошибка перевода '{text[:30]}...' (попытка {attempt + 1}/{self.max_retries}): {e}")
                    time.sleep(wait_time)
                self.stats['retries'] += 1
        
        # Все попытки исчерпаны
        error_msg = f"Не удалось перевести после {self.max_retries} попыток: {last_exception}"
        logging.error(f"{error_msg}. Текст: '{text[:50]}...'")
        self.state_manager.mark_failed_translation(text, error_msg)
        self.stats['failed'] += 1
        return text
    
    def translate_batch(self, texts: List[str]) -> List[str]:
        """Переводит батч текстов, оптимизируя запросы."""
        if not texts:
            return []
        
        results = []
        
        for text in texts:
            if not text or not isinstance(text, str):
                results.append(text)
                continue
                
            cached = self.state_manager.get_from_cache(text)
            if cached:
                results.append(cached)
                self.stats['cached'] += 1
            else:
                translated = self.translate_text_with_retry(text)
                results.append(translated)
        
        return results
    
    def get_translation_stats(self) -> Dict[str, Any]:
        """Возвращает статистику перевода"""
        elapsed = time.time() - self.session_start_time
        return {
            **self.stats,
            'total_requests': self.total_requests,
            'elapsed_time': elapsed,
            'requests_per_second': self.total_requests / elapsed if elapsed > 0 else 0,
            'cache_size': len(self.state_manager.state['translation_cache']),
            'failed_count': self.state_manager.get_failed_count()
        }
    
    def get_sheet_info(self, file_path: str) -> Dict[str, List[str]]:
        """Возвращает информацию о листах и колонках файла."""
        try:
            excel_file = pd.ExcelFile(file_path)
            sheet_info = {}
            
            for sheet_name in excel_file.sheet_names:
                df = excel_file.parse(sheet_name, nrows=1)
                sheet_info[sheet_name] = df.columns.tolist()
            
            return sheet_info
        except Exception as e:
            logging.error(f"Ошибка при чтении файла {file_path}: {e}")
            return {}
    
    def get_sheet_preview(self, file_path: str, sheet_name: str, preview_rows: int = 10) -> List[Dict]:
        """Возвращает превью данных листа для отображения в интерфейсе."""
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=preview_rows)
            df = df.fillna('')
            return df.to_dict('records')
        except Exception as e:
            logging.error(f"Ошибка при чтении превью листа {sheet_name}: {e}")
            return []
    
    def estimate_sheet_volume(self, file_path: str, sheet_name: str, columns: List[str]) -> int:
        """Оценивает объем перевода для листа."""
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            total_cells = 0
            
            for col in columns:
                if col in df.columns and df[col].dtype == 'object':
                    non_empty = df[col].dropna()
                    # Считаем только тексты, которые нужно переводить
                    translatable = non_empty[non_empty.apply(self._should_translate)]
                    total_cells += len(translatable)
            
            return total_cells
        except Exception as e:
            logging.error(f"Ошибка оценки объема для листа {sheet_name}: {e}")
            return 0
    
    def process_sheet_incrementally(self, 
                                  file_path: str,
                                  output_path: str,
                                  sheet_name: str,
                                  columns: List[str],
                                  progress_callback: Optional[Callable] = None,
                                  stop_event: Optional[Callable] = None) -> bool:
        """Обрабатывает один лист с поддержкой инкрементальности."""
        try:
            if self.state_manager.is_sheet_completed(sheet_name):
                logging.info(f"Лист '{sheet_name}' уже переведен, пропускаем")
                return True
            
            logging.info(f"Начало обработки листа: {sheet_name}")
            
            # Читаем исходные данные
            df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str, keep_default_na=False)
            original_row_count = len(df)
            
            # Оцениваем общий объем
            total_cells = self.estimate_sheet_volume(file_path, sheet_name, columns)
            if total_cells == 0:
                logging.info(f"Нет данных для перевода в листе '{sheet_name}'")
                self.state_manager.mark_sheet_completed(sheet_name)
                return True
            
            processed_cells = 0
            
            if progress_callback:
                progress_callback(sheet_name, 0, f"Начало перевода листа {sheet_name} ({total_cells} ячеек)")
            
            # Обрабатываем каждую колонку
            for col_idx, col in enumerate(columns):
                if col not in df.columns:
                    logging.warning(f"Колонка '{col}' не найдена в листе '{sheet_name}'")
                    continue
                
                if stop_event and stop_event():
                    logging.info(f"Остановка перевода листа {sheet_name}")
                    return False
                
                logging.info(f"Перевод колонки: {col}")
                
                if progress_callback:
                    column_progress = (col_idx / len(columns)) * 50
                    progress_callback(sheet_name, column_progress, f"Перевод колонки {col}")
                
                # Получаем уникальные значения для оптимизации
                unique_values = df[col].unique().tolist()
                translation_dict = {}
                
                # Переводим только уникальные значения
                translatable_values = [val for val in unique_values if self._should_translate(val)]
                
                if translatable_values:
                    # Разбиваем на батчи для перевода
                    for i in range(0, len(translatable_values), self.batch_size):
                        if stop_event and stop_event():
                            return False
                        
                        batch = translatable_values[i:i + self.batch_size]
                        translated_batch = self.translate_batch(batch)
                        
                        # Создаем словарь переводов
                        for original, translated in zip(batch, translated_batch):
                            translation_dict[original] = translated
                        
                        processed_cells += len(batch)
                        
                        # Обновляем прогресс
                        if progress_callback and total_cells > 0:
                            translation_progress = 50 + (processed_cells / total_cells) * 50
                            current_stats = self.get_translation_stats()
                            status_msg = (f"Переведено {processed_cells}/{total_cells} ячеек | "
                                        f"Скорость: {current_stats['requests_per_second']:.1f} запр/сек | "
                                        f"Кэш: {current_stats['cache_size']}")
                            progress_callback(sheet_name, translation_progress, status_msg)
                
                # Применяем переводы ко всему столбцу
                df[col] = df[col].map(translation_dict).fillna(df[col])
            
            # Сохраняем результат
            mode = 'a' if os.path.exists(output_path) else 'w'
            if_sheet_exists = 'replace' if mode == 'a' else None
            
            with pd.ExcelWriter(output_path, engine='openpyxl', mode=mode, 
                              if_sheet_exists=if_sheet_exists) as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Отмечаем лист как завершенный
            self.state_manager.mark_sheet_completed(sheet_name)
            
            final_stats = self.get_translation_stats()
            completion_msg = (f"Лист '{sheet_name}' переведен успешно | "
                            f"Всего переведено: {final_stats['translated']} | "
                            f"Из кэша: {final_stats['cached']} | "
                            f"Ошибки: {final_stats['failed']}")
            
            if progress_callback:
                progress_callback(sheet_name, 100, completion_msg)
            
            logging.info(f"Лист '{sheet_name}' обработан успешно. Строк: {original_row_count}")
            return True
            
        except Exception as e:
            logging.error(f"Ошибка обработки листа '{sheet_name}': {e}")
            return False
    
    def process_excel_file(self, 
                          file_path: str, 
                          output_path: str,
                          selected_sheets: Dict[str, List[str]],
                          progress_callback: Optional[Callable] = None,
                          state_manager: Optional[TranslationStateManager] = None,
                          stop_event: Optional[Callable] = None) -> bool:
        """Основной метод для обработки Excel файла с поддержкой инкрементальности."""
        if state_manager:
            self.state_manager = state_manager
        
        try:
            logging.info(f"Начало обработки файла: {file_path}")
            
            # Сбрасываем статистику сессии
            self.session_start_time = time.time()
            self.total_requests = 0
            self.stats = {'translated': 0, 'cached': 0, 'failed': 0, 'retries': 0}
            
            total_sheets = len(selected_sheets)
            processed_sheets = 0
            
            # Обрабатываем каждый лист отдельно
            for sheet_name, columns in selected_sheets.items():
                if stop_event and stop_event():
                    logging.info("Перевод остановлен пользователем")
                    return False
                
                # Пропускаем уже завершенные листы
                if self.state_manager.is_sheet_completed(sheet_name):
                    logging.info(f"Лист '{sheet_name}' уже переведен, пропускаем")
                    processed_sheets += 1
                    continue
                
                # Обрабатываем лист
                success = self.process_sheet_incrementally(
                    file_path, output_path, sheet_name, columns, 
                    progress_callback, stop_event
                )
                
                if success:
                    processed_sheets += 1
                else:
                    if stop_event and stop_event():
                        return False
                    logging.error(f"Ошибка обработки листа '{sheet_name}', продолжаем со следующим")
                
                # Обновляем общий прогресс
                if progress_callback:
                    overall_progress = (processed_sheets / total_sheets) * 100
                    current_stats = self.get_translation_stats()
                    status_msg = (f"Обработано листов: {processed_sheets}/{total_sheets} | "
                                f"Всего переведено: {current_stats['translated']} | "
                                f"Ошибки: {current_stats['failed']}")
                    progress_callback(sheet_name, overall_progress, status_msg)
            
            final_stats = self.get_translation_stats()
            logging.info(f"Файл успешно обработан: {output_path}")
            logging.info(f"Итоговая статистика: {final_stats}")
            
            return True
            
        except Exception as e:
            logging.error(f"Критическая ошибка при обработке файла: {e}")
            return False