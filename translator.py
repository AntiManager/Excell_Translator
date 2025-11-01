# [file name]: translator.py
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Улучшенный модуль для перевода Excel-файлов с поддержкой инкрементального перевода и управления состоянием.
"""
import time
import pandas as pd
from googletrans import Translator
import logging
import json
import os
from pathlib import Path
from typing import Dict, List, Set, Any, Optional, Callable
from datetime import datetime


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
            'version': '1.0',
            'created_at': datetime.now().isoformat(),
            'last_updated': datetime.now().isoformat(),
            'file_path': '',
            'selected_sheets': {},
            'completed_sheets': {},
            'sheet_progress': {},
            'translation_cache': {}
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
            'version': '1.0',
            'created_at': datetime.now().isoformat(),
            'last_updated': datetime.now().isoformat(),
            'file_path': '',
            'selected_sheets': {},
            'completed_sheets': {},
            'sheet_progress': {},
            'translation_cache': {}
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
        # Сохраняем только каждые 100 записей для производительности
        if len(self.state['translation_cache']) % 100 == 0:
            self.save_state()
    
    def get_from_cache(self, original: str) -> Optional[str]:
        """Получает перевод из кэша"""
        return self.state['translation_cache'].get(original)


class ExcelTranslator:
    def __init__(self, delay: float = 0.1, max_retries: int = 3, batch_size: int = 100):
        self.translator = Translator()
        self.delay = delay  # Задержка между запросами
        self.max_retries = max_retries
        self.batch_size = batch_size
        self.state_manager = TranslationStateManager()
    
    def translate_text_with_retry(self, text: str) -> str:
        """Переводит текст с повторными попытками при ошибках."""
        if not isinstance(text, str) or not text.strip():
            return text
        
        # Проверяем кэш состояния
        cached = self.state_manager.get_from_cache(text)
        if cached:
            return cached
        
        for attempt in range(self.max_retries):
            try:
                time.sleep(self.delay)  # Задержка между запросами
                translated = self.translator.translate(text, dest='ru').text
                self.state_manager.add_to_cache(text, translated)
                return translated
            except Exception as e:
                logging.warning(f"Попытка {attempt + 1}/{self.max_retries} не удалась для текста '{text[:50]}...': {e}")
                if attempt == self.max_retries - 1:
                    logging.error(f"Не удалось перевести текст после {self.max_retries} попыток: '{text[:50]}...'")
                    return text
                time.sleep(1)  # Задержка перед повторной попыткой
    
    def translate_batch(self, texts: List[str]) -> List[str]:
        """Переводит батч текстов, оптимизируя запросы."""
        if not texts:
            return []
        
        # Фильтруем уже переведенные тексты
        unique_texts = []
        translation_map = {}
        results = []
        
        for text in texts:
            if not text or not isinstance(text, str):
                results.append(text)
                continue
                
            cached = self.state_manager.get_from_cache(text)
            if cached:
                results.append(cached)
            else:
                unique_texts.append(text)
                translation_map[text] = len(results)
                results.append(None)  # placeholder
        
        # Переводим только уникальные непереведенные тексты
        for i in range(0, len(unique_texts), self.batch_size):
            batch = unique_texts[i:i + self.batch_size]
            try:
                time.sleep(self.delay)
                translations = self.translator.translate(batch, dest='ru')
                
                for j, translation in enumerate(translations):
                    original_text = batch[j]
                    translated_text = translation.text
                    
                    # Обновляем результаты и кэш
                    result_index = translation_map[original_text]
                    results[result_index] = translated_text
                    self.state_manager.add_to_cache(original_text, translated_text)
                    
            except Exception as e:
                logging.error(f"Ошибка перевода батча: {e}")
                # При ошибке переводим по одному
                for text in batch:
                    translated = self.translate_text_with_retry(text)
                    result_index = translation_map[text]
                    results[result_index] = translated
        
        return results
    
    def get_sheet_info(self, file_path: str) -> Dict[str, List[str]]:
        """Возвращает информацию о листах и колонках файла."""
        try:
            excel_file = pd.ExcelFile(file_path)
            sheet_info = {}
            
            for sheet_name in excel_file.sheet_names:
                df = excel_file.parse(sheet_name, nrows=1)  # Читаем только первую строку для получения колонок
                sheet_info[sheet_name] = df.columns.tolist()
            
            return sheet_info
        except Exception as e:
            logging.error(f"Ошибка при чтении файла {file_path}: {e}")
            return {}
    
    def get_sheet_preview(self, file_path: str, sheet_name: str, preview_rows: int = 10) -> List[Dict]:
        """Возвращает превью данных листа для отображения в интерфейсе."""
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=preview_rows)
            # Заменяем NaN на пустые строки и преобразуем в словари
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
                    # Считаем только непустые текстовые ячейки
                    non_empty = df[col].dropna()
                    total_cells += len(non_empty[non_empty.str.strip() != ''])
            
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
            
            # Оцениваем общий объем
            total_cells = self.estimate_sheet_volume(file_path, sheet_name, columns)
            processed_cells = 0
            
            if progress_callback:
                progress_callback(sheet_name, 0, f"Начало перевода листа {sheet_name}")
            
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
                    progress_callback(sheet_name, (col_idx / len(columns)) * 50, 
                                    f"Перевод колонки {col}")
                
                # Получаем уникальные значения для батч-перевода
                unique_values = df[col].unique().tolist()
                translation_dict = {}
                
                # Разбиваем на батчи для перевода
                for i in range(0, len(unique_values), self.batch_size):
                    if stop_event and stop_event():
                        return False
                    
                    batch = unique_values[i:i + self.batch_size]
                    translated_batch = self.translate_batch(batch)
                    
                    # Создаем словарь переводов
                    for original, translated in zip(batch, translated_batch):
                        translation_dict[original] = translated
                    
                    processed_cells += len(batch)
                    
                    # Обновляем прогресс
                    if progress_callback and total_cells > 0:
                        progress = 50 + (processed_cells / total_cells) * 50
                        progress_callback(sheet_name, progress, 
                                        f"Переведено {processed_cells}/{total_cells} ячеек")
                
                # Применяем переводы ко всему столбцу
                df[col] = df[col].map(translation_dict)
            
            # Сохраняем результат
            mode = 'a' if os.path.exists(output_path) else 'w'
            with pd.ExcelWriter(output_path, engine='openpyxl', mode=mode) as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Отмечаем лист как завершенный
            self.state_manager.mark_sheet_completed(sheet_name)
            
            if progress_callback:
                progress_callback(sheet_name, 100, f"Лист '{sheet_name}' переведен успешно")
            
            logging.info(f"Лист '{sheet_name}' обработан успешно")
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
                    # Продолжаем со следующим листом при ошибке
                    logging.error(f"Ошибка обработки листа '{sheet_name}', продолжаем со следующим")
                
                # Обновляем общий прогресс
                if progress_callback:
                    overall_progress = (processed_sheets / total_sheets) * 100
                    progress_callback(sheet_name, overall_progress, 
                                    f"Обработано листов: {processed_sheets}/{total_sheets}")
            
            logging.info(f"Файл успешно обработан: {output_path}")
            return True
            
        except Exception as e:
            logging.error(f"Критическая ошибка при обработке файла: {e}")
            return False