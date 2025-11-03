#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI приложение для перевода Excell файлов.
Улучшенная версия со стабильным переводчиком.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
import threading
import logging
import sys
import json
import time
from typing import Dict, List, Optional, Tuple

# Добавляем путь для импорта translator
sys.path.append(Path(__file__).parent)

from translator import ExcelTranslator, TranslationStateManager


class ResizablePanedWindow(ttk.PanedWindow):
    """Переопределенный PanedWindow с улучшенным управлением размерами"""
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.bind("<Button-1>", self.on_click)
        self.bind("<B1-Motion>", self.on_drag)
        self.bind("<ButtonRelease-1>", self.on_release)
        self.dragging = False
    
    def on_click(self, event):
        self.dragging = True
    
    def on_drag(self, event):
        if self.dragging:
            self.place_configure(x=event.x)
    
    def on_release(self, event):
        self.dragging = False


class SheetPreviewDialog:
    """Диалог предпросмотра данных листа с выбором колонок"""
    
    def __init__(self, parent, sheet_name: str, columns: List[str], preview_data: List[Dict]):
        self.parent = parent
        self.sheet_name = sheet_name
        self.columns = columns
        self.preview_data = preview_data
        self.selected_columns = []
        self.dialog = None
        
    def show(self) -> List[str]:
        """Показывает диалог и возвращает выбранные колонки"""
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title(f"Выбор колонок - {self.sheet_name}")
        self.dialog.geometry("1000x600")
        self.dialog.minsize(800, 400)
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # Основной контейнер с возможностью изменения размеров
        main_paned = ttk.PanedWindow(self.dialog, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Левая панель - выбор колонок
        left_frame = ttk.LabelFrame(main_paned, text="Колонки для перевода", padding="10")
        main_paned.add(left_frame, weight=1)
        
        # Правая панель - предпросмотр данных
        right_frame = ttk.LabelFrame(main_paned, text="Предпросмотр данных (первые 10 строк)", padding="10")
        main_paned.add(right_frame, weight=2)
        
        self._setup_columns_frame(left_frame)
        self._setup_preview_frame(right_frame)
        
        # Кнопки управления
        button_frame = ttk.Frame(self.dialog)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(button_frame, text="Применить", 
                  command=self._apply_selection).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Отмена", 
                  command=self._cancel).pack(side=tk.RIGHT, padx=5)
        
        # Устанавливаем начальное положение разделителя
        main_paned.sashpos(0, 300)
        
        self.parent.wait_window(self.dialog)
        return self.selected_columns
    
    def _setup_columns_frame(self, parent):
        """Настраивает фрейм выбора колонок"""
        # Поле поиска
        search_frame = ttk.Frame(parent)
        search_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(search_frame, text="Поиск:").pack(side=tk.LEFT, padx=(0, 5))
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        search_entry.bind('<KeyRelease>', lambda e: self._filter_columns(search_var.get()))
        
        # Фрейм для чекбоксов с прокруткой
        columns_frame = ttk.Frame(parent)
        columns_frame.pack(fill=tk.BOTH, expand=True)
        
        # Canvas и scrollbar для чекбоксов
        canvas = tk.Canvas(columns_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(columns_frame, orient=tk.VERTICAL, command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Создаем чекбоксы для всех колонок
        self.column_vars = {}
        for i, column in enumerate(self.columns):
            var = tk.BooleanVar(value=True)
            self.column_vars[column] = var
            
            cb = ttk.Checkbutton(self.scrollable_frame, text=column, variable=var,
                               command=self._update_preview_highlighting)
            cb.grid(row=i, column=0, sticky=tk.W, pady=2, padx=5)
    
    def _setup_preview_frame(self, parent):
        """Настраивает фрейм предпросмотра данных"""
        # Создаем Treeview для отображения данных
        self.preview_tree = ttk.Treeview(parent, columns=self.columns, show='headings', height=15)
        
        # Настраиваем заголовки колонок
        for col in self.columns:
            self.preview_tree.heading(col, text=col)
            self.preview_tree.column(col, width=100, minwidth=50)
        
        # Заполняем данными
        for i, row in enumerate(self.preview_data):
            values = [row.get(col, '') for col in self.columns]
            self.preview_tree.insert('', tk.END, values=values, tags=(f"row_{i}",))
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.preview_tree.yview)
        h_scrollbar = ttk.Scrollbar(parent, orient=tk.HORIZONTAL, command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Размещаем элементы
        self.preview_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)
        
        # Настраиваем теги для выделения
        self.preview_tree.tag_configure('selected_col', background='#e6f3ff')
    
    def _filter_columns(self, search_text: str):
        """Фильтрует колонки по поисковому запросу"""
        search_text = search_text.lower()
        for widget in self.scrollable_frame.winfo_children():
            if isinstance(widget, ttk.Checkbutton):
                text = widget.cget('text').lower()
                if search_text in text:
                    widget.grid()
                else:
                    widget.grid_remove()
    
    def _update_preview_highlighting(self):
        """Обновляет выделение выбранных колонок в предпросмотре"""
        selected_cols = [col for col, var in self.column_vars.items() if var.get()]
        
        # Сбрасываем все теги
        for col in self.columns:
            self.preview_tree.tag_configure(col, background='')
        
        # Устанавливаем теги для выбранных колонок
        for col in selected_cols:
            if col in self.columns:
                col_index = self.columns.index(col)
                self.preview_tree.tag_configure(f"col_{col_index}", background='#e6f3ff')
    
    def _apply_selection(self):
        """Применяет выбранные колонки"""
        self.selected_columns = [col for col, var in self.column_vars.items() if var.get()]
        self.dialog.destroy()
    
    def _cancel(self):
        """Отменяет выбор"""
        self.selected_columns = []
        self.dialog.destroy()


class StatisticsDialog:
    """Диалог отображения статистики перевода"""
    
    def __init__(self, parent, stats: Dict):
        self.parent = parent
        self.stats = stats
        self.dialog = None
        
    def show(self):
        """Показывает диалог статистики"""
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("Статистика перевода")
        self.dialog.geometry("500x400")
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Статистика перевода", 
                 font=('Arial', 14, 'bold')).pack(pady=(0, 20))
        
        # Статистика
        stats_text = scrolledtext.ScrolledText(main_frame, height=15, wrap=tk.WORD)
        stats_text.pack(fill=tk.BOTH, expand=True)
        
        stats_str = self._format_stats()
        stats_text.insert(tk.END, stats_str)
        stats_text.config(state=tk.DISABLED)
        
        # Кнопка закрытия
        ttk.Button(main_frame, text="Закрыть", 
                  command=self.dialog.destroy).pack(pady=10)
        
        self.parent.wait_window(self.dialog)
    
    def _format_stats(self) -> str:
        """Форматирует статистику для отображения"""
        stats = self.stats
        return f"""
ОБЩАЯ СТАТИСТИКА:

Производительность:
• Всего переведено: {stats.get('translated', 0)}
• Использовано из кэша: {stats.get('cached', 0)}
• Всего запросов: {stats.get('total_requests', 0)}
• Скорость: {stats.get('requests_per_second', 0):.1f} запросов/сек
• Время работы: {stats.get('elapsed_time', 0):.1f} сек

Ошибки и повторы:
• Неудачных переводов: {stats.get('failed', 0)}
• Повторных попыток: {stats.get('retries', 0)}
• Размер кэша: {stats.get('cache_size', 0)} записей

ЭФФЕКТИВНОСТЬ:
• Эффективность кэша: {(stats.get('cached', 0) / max(1, stats.get('cached', 0) + stats.get('translated', 0)) * 100):.1f}%
• Успешных переводов: {(stats.get('translated', 0) / max(1, stats.get('total_requests', 0)) * 100):.1f}%
"""


class ExcelTranslatorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excell Translator v3.0 - Улучшенная версия")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 600)
        
        self.translator = ExcelTranslator()
        self.state_manager = TranslationStateManager()
        self.file_path: Optional[str] = None
        self.sheet_info: Dict[str, List[str]] = {}
        self.sheet_previews: Dict[str, List[Dict]] = {}
        self.selected_sheets: Dict[str, List[str]] = {}
        
        self.current_translation_thread: Optional[threading.Thread] = None
        self.stop_translation = False
        
        self.setup_ui()
        self.setup_logging()
        
        # Загружаем последнее состояние
        self._load_last_state()
    
    def setup_ui(self):
        """Создает улучшенный интерфейс пользователя с изменяемыми размерами."""
        # Главный контейнер с изменяемыми разделами
        self.main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.main_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Левая панель - управление файлами и листами
        self.left_frame = ttk.LabelFrame(self.main_paned, text="Управление переводом", padding="10")
        self.main_paned.add(self.left_frame, weight=1)
        
        # Правая панель - лог и прогресс
        self.right_frame = ttk.LabelFrame(self.main_paned, text="Прогресс и логирование", padding="10")
        self.main_paned.add(self.right_frame, weight=2)
        
        self._setup_left_panel()
        self._setup_right_panel()
        
        # Устанавливаем начальное положение разделителя
        self.main_paned.sashpos(0, 400)
    
    def _setup_left_panel(self):
        """Настраивает левую панель управления"""
        # Выбор файла
        file_frame = ttk.LabelFrame(self.left_frame, text="Файл Excell", padding="5")
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.file_var = tk.StringVar()
        file_entry_frame = ttk.Frame(file_frame)
        file_entry_frame.pack(fill=tk.X, pady=5)
        
        self.file_entry = ttk.Entry(file_entry_frame, textvariable=self.file_var, state='readonly')
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        ttk.Button(file_entry_frame, text="Выбрать файл", 
                  command=self.select_file).pack(side=tk.RIGHT)
        
        # Информация о файле
        self.file_info_var = tk.StringVar(value="Файл не выбран")
        ttk.Label(file_frame, textvariable=self.file_info_var).pack(anchor=tk.W)
        
        # Листы и колонки
        sheets_frame = ttk.LabelFrame(self.left_frame, text="Листы и колонки", padding="5")
        sheets_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Treeview с прокруткой
        tree_container = ttk.Frame(sheets_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        self.tree = ttk.Treeview(tree_container, columns=('Status', 'Columns', 'Progress'), 
                                show='tree headings', height=12)
        self.tree.heading('#0', text='Лист')
        self.tree.column('#0', width=150, minwidth=100)
        self.tree.heading('Status', text='Статус')
        self.tree.column('Status', width=80, minwidth=60)
        self.tree.heading('Columns', text='Колонки')
        self.tree.column('Columns', width=80, minwidth=60)
        self.tree.heading('Progress', text='Прогресс')
        self.tree.column('Progress', width=80, minwidth=60)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)
        
        # Кнопки управления листами
        btn_frame = ttk.Frame(sheets_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(btn_frame, text="Настроить колонки", 
                  command=self.configure_columns).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Выбрать все", 
                  command=self.select_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Снять выделение", 
                  command=self.deselect_all).pack(side=tk.LEFT, padx=2)
        
        # Управление переводом
        control_frame = ttk.LabelFrame(self.left_frame, text="Управление переводом", padding="5")
        control_frame.pack(fill=tk.X)
        
        # Оценка объема
        self.volume_var = tk.StringVar(value="Объем: не оценен")
        ttk.Label(control_frame, textvariable=self.volume_var).pack(anchor=tk.W, pady=2)
        
        # Кнопки перевода
        btn_frame2 = ttk.Frame(control_frame)
        btn_frame2.pack(fill=tk.X, pady=5)
        
        self.translate_btn = ttk.Button(btn_frame2, text="Начать перевод", 
                                       command=self.start_translation)
        self.translate_btn.pack(side=tk.LEFT, padx=2)
        
        self.stop_btn = ttk.Button(btn_frame2, text="Остановить", 
                                  command=self.stop_translation_process, state='disabled')
        self.stop_btn.pack(side=tk.LEFT, padx=2)
        
        self.resume_btn = ttk.Button(btn_frame2, text="Продолжить", 
                                    command=self.resume_translation)
        self.resume_btn.pack(side=tk.LEFT, padx=2)
        
        ttk.Button(btn_frame2, text="Статистика", 
                  command=self.show_statistics).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(btn_frame2, text="Сбросить прогресс", 
                  command=self.reset_progress).pack(side=tk.LEFT, padx=2)
        
        # Bind events
        self.tree.bind('<Double-1>', lambda e: self.configure_columns())
    
    def _setup_right_panel(self):
        """Настраивает правую панель с прогрессом и логами"""
        # Прогресс-бар
        progress_frame = ttk.LabelFrame(self.right_frame, text="Прогресс перевода", padding="5")
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.overall_progress_var = tk.DoubleVar()
        self.overall_progress = ttk.Progressbar(progress_frame, variable=self.overall_progress_var, maximum=100)
        self.overall_progress.pack(fill=tk.X, pady=5)
        
        self.status_var = tk.StringVar(value="Готов к работе")
        ttk.Label(progress_frame, textvariable=self.status_var).pack(anchor=tk.W)
        
        # Детальный прогресс по листам
        details_frame = ttk.LabelFrame(self.right_frame, text="Прогресс по листам", padding="5")
        details_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.details_text = tk.Text(details_frame, height=8, wrap=tk.WORD)
        details_scrollbar = ttk.Scrollbar(details_frame, orient=tk.VERTICAL, command=self.details_text.yview)
        self.details_text.configure(yscrollcommand=details_scrollbar.set)
        
        self.details_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        details_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Логи
        log_frame = ttk.LabelFrame(self.right_frame, text="Лог выполнения", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
    
    def setup_logging(self):
        """Настраивает логирование в GUI."""
        class TextHandler(logging.Handler):
            def __init__(self, text_widget):
                super().__init__()
                self.text_widget = text_widget
            
            def emit(self, record):
                msg = self.format(record)
                self.text_widget.config(state=tk.NORMAL)
                self.text_widget.insert(tk.END, msg + '\n')
                self.text_widget.see(tk.END)
                self.text_widget.config(state=tk.DISABLED)
        
        text_handler = TextHandler(self.log_text)
        text_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logging.getLogger().addHandler(text_handler)
        logging.getLogger().setLevel(logging.INFO)
    
    def _load_last_state(self):
        """Загружает последнее состояние перевода"""
        state = self.state_manager.load_state()
        if state and 'file_path' in state:
            self.file_path = state['file_path']
            self.file_var.set(self.file_path)
            self.load_file_info()
            
            if 'selected_sheets' in state:
                self.selected_sheets = state['selected_sheets']
                self._update_treeview_from_state()
    
    def _update_treeview_from_state(self):
        """Обновляет treeview из сохраненного состояния"""
        for item in self.tree.get_children():
            sheet_name = self.tree.item(item, 'text')
            if sheet_name in self.selected_sheets:
                selected_cols = self.selected_sheets[sheet_name]
                status = '✓ Выбран' if selected_cols else '✗ Не выбран'
                self.tree.set(item, 'Status', status)
                self.tree.set(item, 'Columns', f'{len(selected_cols)} колонок')
                self.tree.item(item, tags=('selected' if selected_cols else 'not_selected',))
    
    def select_file(self):
        """Выбор файла Excell."""
        file_path = filedialog.askopenfilename(
            title="Выберите файл Excell",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            self.file_path = file_path
            self.file_var.set(file_path)
            self.state_manager.update_state({'file_path': file_path})
            self.load_file_info()
    
    def load_file_info(self):
        """Загружает информацию о листах и колонках файла."""
        if not self.file_path:
            return
        
        self.status_var.set("Чтение информации о файле...")
        self.root.update()
        
        try:
            self.sheet_info = self.translator.get_sheet_info(self.file_path)
            
            self.sheet_previews = {}
            for sheet_name in self.sheet_info.keys():
                preview = self.translator.get_sheet_preview(self.file_path, sheet_name, preview_rows=10)
                self.sheet_previews[sheet_name] = preview
            
            self.populate_treeview()
            self._estimate_translation_volume()
            
            self.status_var.set("Файл загружен успешно")
            logging.info(f"Загружен файл: {self.file_path}")
            logging.info(f"Найдено листов: {len(self.sheet_info)}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось прочитать файл: {e}")
            self.status_var.set("Ошибка загрузки файла")
            logging.error(f"Ошибка загрузки файла: {e}")
    
    def _estimate_translation_volume(self):
        """Оценивает объем перевода"""
        total_cells = 0
        for sheet_name, columns in self.selected_sheets.items():
            if sheet_name in self.sheet_previews and columns:
                preview_data = self.sheet_previews[sheet_name]
                if preview_data:
                    total_cells += len(preview_data) * len(columns)
        
        if total_cells > 0:
            estimated_time = total_cells * 0.3  # Улучшенная оценка времени
            self.volume_var.set(f"Объем: ~{total_cells} ячеек, время: ~{estimated_time/60:.1f} мин")
        else:
            self.volume_var.set("Объем: не оценен (выберите колонки)")
    
    def populate_treeview(self):
        """Заполняет treeview информацией о листах и колонках."""
        self.tree.delete(*self.tree.get_children())
        
        for sheet_name, columns in self.sheet_info.items():
            selected_columns = self.selected_sheets.get(sheet_name, columns.copy())
            self.selected_sheets[sheet_name] = selected_columns
            
            status = '✓ Выбран' if selected_columns else '✗ Не выбран'
            progress = self.state_manager.get_sheet_progress(sheet_name)
            
            item = self.tree.insert('', 'end', text=sheet_name, 
                                  values=(status, f'{len(selected_columns)} колонок', progress))
            
            tags = ('selected',) if selected_columns else ('not_selected',)
            self.tree.item(item, tags=tags)
        
        # Настраиваем теги для цветового выделения
        self.tree.tag_configure('selected', background='#e6f3ff')
        self.tree.tag_configure('not_selected', background='#f5f5f5')
        self.tree.tag_configure('completed', background='#e8f5e8')
        self.tree.tag_configure('in_progress', background='#fff3cd')
    
    def select_all(self):
        """Выбирает все листы."""
        for item in self.tree.get_children():
            sheet_name = self.tree.item(item, 'text')
            self.selected_sheets[sheet_name] = self.sheet_info[sheet_name].copy()
            self.tree.set(item, 'Status', '✓ Выбран')
            self.tree.set(item, 'Columns', f'{len(self.sheet_info[sheet_name])} колонок')
            self.tree.item(item, tags=('selected',))
        
        self._estimate_translation_volume()
        self.state_manager.update_state({'selected_sheets': self.selected_sheets})
        logging.info("Выбраны все листы и колонки")
    
    def deselect_all(self):
        """Снимает выделение со всех листов."""
        for item in self.tree.get_children():
            sheet_name = self.tree.item(item, 'text')
            self.selected_sheets[sheet_name] = []
            self.tree.set(item, 'Status', '✗ Не выбран')
            self.tree.set(item, 'Columns', '0 колонок')
            self.tree.item(item, tags=('not_selected',))
        
        self._estimate_translation_volume()
        self.state_manager.update_state({'selected_sheets': self.selected_sheets})
        logging.info("Снято выделение со всех листов")
    
    def configure_columns(self):
        """Открывает диалог настройки колонок для выбранного листа с предпросмотром."""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите лист для настройки колонок")
            return
        
        item = selection[0]
        sheet_name = self.tree.item(item, 'text')
        
        if sheet_name not in self.sheet_previews:
            messagebox.showerror("Ошибка", f"Нет данных предпросмотра для листа '{sheet_name}'")
            return
        
        preview_data = self.sheet_previews[sheet_name]
        dialog = SheetPreviewDialog(self.root, sheet_name, self.sheet_info[sheet_name], preview_data)
        selected_columns = dialog.show()
        
        if selected_columns:
            self.selected_sheets[sheet_name] = selected_columns
            
            status = '✓ Выбран' if selected_columns else '✗ Не выбран'
            self.tree.set(item, 'Status', status)
            self.tree.set(item, 'Columns', f'{len(selected_columns)} колонок')
            self.tree.item(item, tags=('selected' if selected_columns else 'not_selected',))
            
            self._estimate_translation_volume()
            self.state_manager.update_state({'selected_sheets': self.selected_sheets})
            
            logging.info(f"Настроены колонки для листа '{sheet_name}': выбрано {len(selected_columns)} колонок")
    
    def show_statistics(self):
        """Показывает диалог статистики"""
        stats = self.translator.get_translation_stats()
        dialog = StatisticsDialog(self.root, stats)
        dialog.show()
    
    def update_progress(self, sheet_name: str, progress: float, message: str):
        """Обновляет прогресс-бар и статус."""
        self.overall_progress_var.set(progress)
        self.status_var.set(message)
        
        for item in self.tree.get_children():
            if self.tree.item(item, 'text') == sheet_name:
                progress_text = f"{progress:.1f}%" if progress > 0 else ""
                self.tree.set(item, 'Progress', progress_text)
                
                if progress >= 100:
                    self.tree.item(item, tags=('completed',))
                elif progress > 0:
                    self.tree.item(item, tags=('in_progress',))
                break
        
        self.details_text.config(state=tk.NORMAL)
        self.details_text.insert(tk.END, f"{message}\n")
        self.details_text.see(tk.END)
        self.details_text.config(state=tk.DISABLED)
        
        self.root.update_idletasks()
    
    def start_translation(self):
        """Запускает процесс перевода в отдельном потоке."""
        if not self.file_path:
            messagebox.showwarning("Предупреждение", "Сначала выберите файл Excell")
            return
        
        has_selected = any(columns for columns in self.selected_sheets.values())
        if not has_selected:
            messagebox.showwarning("Предупреждение", "Выберите хотя бы один лист с колонками для перевода")
            return
        
        input_path = Path(self.file_path)
        output_path = input_path.parent / f"{input_path.stem}_ru{input_path.suffix}"
        
        total_volume = sum(len(cols) for cols in self.selected_sheets.values())
        if total_volume > 100:
            response = messagebox.askyesno(
                "Подтверждение", 
                f"Будет переведено {total_volume} колонок. Это может занять значительное время. Продолжить?"
            )
            if not response:
                return
        
        self.stop_translation = False
        self.current_translation_thread = threading.Thread(
            target=self.run_translation,
            args=(str(self.file_path), str(output_path))
        )
        self.current_translation_thread.daemon = True
        self.current_translation_thread.start()
        
        self.translate_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.resume_btn.config(state='disabled')
    
    def stop_translation_process(self):
        """Останавливает процесс перевода."""
        self.stop_translation = True
        self.status_var.set("Остановка перевода...")
        self.stop_btn.config(state='disabled')
        self.resume_btn.config(state='normal')
    
    def resume_translation(self):
        """Продолжает остановленный перевод."""
        self.start_translation()
    
    def reset_progress(self):
        """Сбрасывает весь прогресс перевода."""
        if messagebox.askyesno("Подтверждение", "Сбросить весь прогресс перевода?"):
            self.state_manager.clear_state()
            self.selected_sheets = {}
            self.overall_progress_var.set(0)
            self.status_var.set("Прогресс сброшен")
            self.details_text.config(state=tk.NORMAL)
            self.details_text.delete(1.0, tk.END)
            self.details_text.config(state=tk.DISABLED)
            
            if self.file_path:
                self.load_file_info()
            
            logging.info("Прогресс перевода сброшен")
    
    def run_translation(self, input_path: str, output_path: str):
        """Запускает процесс перевода (выполняется в отдельном потоке)."""
        try:
            sheets_to_process = {
                sheet: columns for sheet, columns in self.selected_sheets.items() 
                if columns
            }
            
            logging.info(f"Начало перевода файла: {input_path}")
            logging.info(f"Будут обработаны листы: {list(sheets_to_process.keys())}")
            
            success = self.translator.process_excel_file(
                input_path,
                output_path,
                sheets_to_process,
                progress_callback=self.update_progress,
                state_manager=self.state_manager,
                stop_event=lambda: self.stop_translation
            )
            
            if success and not self.stop_translation:
                final_stats = self.translator.get_translation_stats()
                stats_msg = (f"Перевод завершен!\n\n"
                           f"Файл сохранен как:\n{output_path}\n\n"
                           f"Статистика:\n"
                           f"• Переведено: {final_stats['translated']}\n"
                           f"• Из кэша: {final_stats['cached']}\n"
                           f"• Ошибки: {final_stats['failed']}\n"
                           f"• Время: {final_stats['elapsed_time']:.1f} сек")
                
                messagebox.showinfo("Успех", stats_msg)
                self.status_var.set("Перевод завершен успешно")
                logging.info("Перевод завершен успешно")
            elif self.stop_translation:
                messagebox.showinfo("Информация", "Перевод остановлен пользователем")
                self.status_var.set("Перевод остановлен")
                logging.info("Перевод остановлен пользователем")
            else:
                messagebox.showerror("Ошибка", "Произошла ошибка при переводе. Проверьте лог для деталей.")
                self.status_var.set("Ошибка перевода")
                logging.error("Перевод завершен с ошибками")
                
        except Exception as e:
            messagebox.showerror("Ошибка", f"Критическая ошибка: {e}")
            logging.error(f"Критическая ошибка: {e}")
        finally:
            self.translate_btn.config(state='normal')
            self.stop_btn.config(state='disabled')
            self.resume_btn.config(state='normal')
            self.stop_translation = False


def main():
    root = tk.Tk()
    app = ExcelTranslatorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()