"""
GUI интерфейс на tkinter для работы с Excel файлами.
Предоставляет интерфейс для открытия/закрытия файлов, добавления/удаления строк,
валидации данных и управления стилями.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional, List, Any

from excel_handler import ExcelWorkbook, ExcelWorksheet, RowValidationError
from openpyxl.styles import Font, PatternFill, Alignment


class SpectrumExcelApp:
    """
    Приложение для работы с файлами спектров в Excel.
    """
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Обработка спектров - Excel")
        self.root.geometry("1000x700")
        
        self.workbook: Optional[ExcelWorkbook] = None
        self.current_sheet: Optional[ExcelWorksheet] = None
        
        self._create_widgets()
        self._create_menu()
    
    def _create_menu(self):
        """Создать меню приложения."""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # Файл меню
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Файл", menu=file_menu)
        file_menu.add_command(label="Открыть...", command=self.open_file, accelerator="Ctrl+O")
        file_menu.add_command(label="Создать новый", command=self.create_new_file, accelerator="Ctrl+N")
        file_menu.add_separator()
        file_menu.add_command(label="Сохранить", command=self.save_file, accelerator="Ctrl+S")
        file_menu.add_command(label="Сохранить как...", command=self.save_file_as)
        file_menu.add_separator()
        file_menu.add_command(label="Закрыть файл", command=self.close_file)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.root.quit, accelerator="Alt+F4")
        
        # Листы меню
        sheets_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Листы", menu=sheets_menu)
        sheets_menu.add_command(label="Создать лист...", command=self.create_sheet)
        sheets_menu.add_command(label="Переключить лист...", command=self.switch_sheet)
        
        # Правка меню
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Правка", menu=edit_menu)
        edit_menu.add_command(label="Добавить строку", command=self.add_row_dialog)
        edit_menu.add_command(label="Удалить строку", command=self.delete_row)
        edit_menu.add_separator()
        edit_menu.add_command(label="Сортировать по столбцу...", command=self.sort_column_dialog)
        
        # Стили меню
        style_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Стили", menu=style_menu)
        style_menu.add_command(label="Применить стиль заголовка", command=self.apply_header_style)
        style_menu.add_command(label="Очистить стили строки", command=self.clear_row_style)
        
        # Привязки клавиш
        self.root.bind('<Control-o>', lambda e: self.open_file())
        self.root.bind('<Control-n>', lambda e: self.create_new_file())
        self.root.bind('<Control-s>', lambda e: self.save_file())
    
    def _create_widgets(self):
        """Создать виджеты интерфейса."""
        # Верхняя панель
        top_frame = ttk.Frame(self.root, padding="5")
        top_frame.pack(side=tk.TOP, fill=tk.X)
        
        self.file_label = ttk.Label(top_frame, text="Файл не открыт", foreground="gray")
        self.file_label.pack(side=tk.LEFT)
        
        self.sheet_label = ttk.Label(top_frame, text="", foreground="blue")
        self.sheet_label.pack(side=tk.LEFT, padx=10)
        
        # Панель инструментов
        toolbar = ttk.Frame(self.root, padding="5")
        toolbar.pack(side=tk.TOP, fill=tk.X)
        
        ttk.Button(toolbar, text="Открыть", command=self.open_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Сохранить", command=self.save_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Добавить строку", command=self.add_row_dialog).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Удалить строку", command=self.delete_row).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Сортировать", command=self.sort_column_dialog).pack(side=tk.LEFT, padx=2)
        
        # Таблица данных
        table_frame = ttk.Frame(self.root, padding="5")
        table_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        # Создаем Treeview с полосой прокрутки
        columns = tuple(f"col{i}" for i in range(20))  # Максимум 20 столбцов
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)
        
        # Настройка заголовков
        for i, col in enumerate(columns):
            self.tree.heading(col, text=f"Столбец {i+1}")
            self.tree.column(col, width=80, anchor=tk.CENTER)
        
        # Полосы прокрутки
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        # Строка состояния
        self.status_var = tk.StringVar(value="Готов")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def update_status(self, message: str):
        """Обновить строку состояния."""
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def refresh_table(self):
        """Обновить отображение таблицы."""
        if not self.current_sheet:
            return
        
        # Очищаем таблицу
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Настраиваем количество видимых столбцов
        max_col = min(self.current_sheet.max_column, 20)
        visible_columns = tuple(f"col{i}" for i in range(max_col))
        self.tree['columns'] = visible_columns
        
        # Обновляем заголовки
        for i, col in enumerate(visible_columns):
            self.tree.heading(col, text=f"Столбец {i+1}")
            self.tree.column(col, width=100, anchor=tk.CENTER)
        
        # Загружаем данные
        for row_num in range(1, self.current_sheet.max_row + 1):
            row_data = []
            for col_num in range(1, max_col + 1):
                cell = self.current_sheet._worksheet.cell(row=row_num, column=col_num)
                row_data.append(cell.value if cell.value is not None else "")
            
            # Добавляем теги для стилизации
            tags = ()
            if row_num == 1:  # Заголовок
                tags = ('header',)
            
            self.tree.insert('', 'end', values=tuple(row_data), tags=tags)
        
        # Настраиваем стили
        self.tree.tag_configure('header', background='#4472C4', foreground='white', font=('Arial', 10, 'bold'))
        
        self.update_status(f"Строк: {self.current_sheet.max_row}, Столбцов: {max_col}")
    
    def open_file(self):
        """Открыть существующий файл."""
        file_path = filedialog.askopenfilename(
            title="Открыть файл Excel",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            if self.workbook:
                self.workbook.close()
            
            self.workbook = ExcelWorkbook(file_path)
            self.current_sheet = self.workbook.get_active_sheet()
            
            self.file_label.config(text=f"Файл: {file_path}", foreground="green")
            self.sheet_label.config(text=f"Лист: {self.current_sheet.name}")
            
            self.refresh_table()
            self.update_status(f"Файл открыт: {file_path}")
            
        except FileExistsError as e:
            messagebox.showerror("Ошибка", str(e))
            self.update_status("Ошибка: файл заблокирован")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{str(e)}")
            self.update_status("Ошибка открытия файла")
    
    def create_new_file(self):
        """Создать новый файл."""
        if self.workbook:
            if not messagebox.askyesno("Подтверждение", "Текущий файл будет закрыт. Продолжить?"):
                return
            self.workbook.close()
        
        self.workbook = ExcelWorkbook()
        self.current_sheet = self.workbook.get_active_sheet()
        
        self.file_label.config(text="Новый файл (не сохранен)", foreground="orange")
        self.sheet_label.config(text=f"Лист: {self.current_sheet.name}")
        
        self.refresh_table()
        self.update_status("Создан новый файл")
    
    def save_file(self):
        """Сохранить текущий файл."""
        if not self.workbook:
            messagebox.showwarning("Предупреждение", "Нет открытого файла")
            return
        
        if not self.workbook.file_path:
            self.save_file_as()
            return
        
        try:
            self.workbook.save()
            self.file_label.config(text=f"Файл: {self.workbook.file_path}", foreground="green")
            self.update_status("Файл сохранен")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{str(e)}")
            self.update_status("Ошибка сохранения")
    
    def save_file_as(self):
        """Сохранить файл как..."""
        if not self.workbook:
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Сохранить файл Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if not file_path:
            return
        
        try:
            self.workbook.save(file_path)
            self.file_label.config(text=f"Файл: {file_path}", foreground="green")
            self.update_status(f"Файл сохранен: {file_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{str(e)}")
    
    def close_file(self):
        """Закрыть текущий файл."""
        if not self.workbook:
            return
        
        if self.workbook.file_path and self.workbook._lock:
            if not messagebox.askyesno("Подтверждение", "Закрыть текущий файл?"):
                return
        
        self.workbook.close()
        self.workbook = None
        self.current_sheet = None
        
        self.file_label.config(text="Файл не открыт", foreground="gray")
        self.sheet_label.config(text="")
        
        # Очищаем таблицу
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        self.update_status("Файл закрыт")
    
    def create_sheet(self):
        """Создать новый лист."""
        if not self.workbook:
            messagebox.showwarning("Предупреждение", "Сначала откройте или создайте файл")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Создать лист")
        dialog.geometry("300x120")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Имя листа:").pack(pady=10)
        name_entry = ttk.Entry(dialog, width=40)
        name_entry.pack(pady=5)
        name_entry.focus()
        
        def on_create():
            name = name_entry.get().strip()
            if not name:
                messagebox.showwarning("Предупреждение", "Введите имя листа", parent=dialog)
                return
            
            try:
                self.workbook.create_sheet(name)
                self.current_sheet = self.workbook.set_active_sheet(name)
                self.sheet_label.config(text=f"Лист: {self.current_sheet.name}")
                self.refresh_table()
                self.update_status(f"Лист '{name}' создан")
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("Ошибка", str(e), parent=dialog)
        
        ttk.Button(dialog, text="Создать", command=on_create).pack(pady=10)
        dialog.bind('<Return>', lambda e: on_create())
    
    def switch_sheet(self):
        """Переключить активный лист."""
        if not self.workbook:
            messagebox.showwarning("Предупреждение", "Нет открытого файла")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Выбрать лист")
        dialog.geometry("300x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Доступные листы:").pack(pady=5)
        
        listbox = tk.Listbox(dialog, width=40, height=8)
        listbox.pack(pady=5)
        
        for name in self.workbook.sheetnames:
            listbox.insert(tk.END, name)
        
        def on_select():
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("Предупреждение", "Выберите лист", parent=dialog)
                return
            
            sheet_name = listbox.get(selection[0])
            try:
                self.current_sheet = self.workbook.set_active_sheet(sheet_name)
                self.sheet_label.config(text=f"Лист: {self.current_sheet.name}")
                self.refresh_table()
                self.update_status(f"Переключен на лист '{sheet_name}'")
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("Ошибка", str(e), parent=dialog)
        
        ttk.Button(dialog, text="Выбрать", command=on_select).pack(pady=5)
    
    def add_row_dialog(self):
        """Диалог добавления строки."""
        if not self.current_sheet:
            messagebox.showwarning("Предупреждение", "Нет активного листа")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Добавить строку")
        dialog.geometry("600x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Данные строки (разделитель: запятая или табуляция):").pack(pady=5)
        
        text_widget = tk.Text(dialog, width=70, height=10)
        text_widget.pack(pady=5, padx=10)
        text_widget.focus()
        
        # Пример из буфера обмена или предыдущих данных
        ttk.Label(dialog, text="Опции валидации:", foreground="blue").pack(pady=(10, 0))
        
        validate_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(dialog, text="Валидировать данные", variable=validate_var).pack()
        
        def on_add():
            data_text = text_widget.get("1.0", tk.END).strip()
            if not data_text:
                messagebox.showwarning("Предупреждение", "Введите данные", parent=dialog)
                return
            
            # Парсим данные
            if '\t' in data_text:
                row_data = [cell.strip() for cell in data_text.split('\t')]
            else:
                row_data = [cell.strip() for cell in data_text.split(',')]
            
            # Преобразуем типы
            parsed_data = []
            for cell in row_data:
                if not cell:
                    parsed_data.append(None)
                    continue
                
                # Пробуем int
                try:
                    parsed_data.append(int(cell))
                    continue
                except ValueError:
                    pass
                
                # Пробуем float
                try:
                    parsed_data.append(float(cell))
                    continue
                except ValueError:
                    pass
                
                # Оставляем как строку
                parsed_data.append(cell)
            
            try:
                self.current_sheet.add_row(
                    parsed_data,
                    validate=validate_var.get()
                )
                self.refresh_table()
                self.update_status(f"Строка добавлена ({len(parsed_data)} ячеек)")
                dialog.destroy()
            except RowValidationError as e:
                messagebox.showerror("Ошибка валидации", str(e), parent=dialog)
            except Exception as e:
                messagebox.showerror("Ошибка", str(e), parent=dialog)
        
        ttk.Button(dialog, text="Добавить", command=on_add).pack(pady=10)
        dialog.bind('<Control-Return>', lambda e: on_add())
    
    def delete_row(self):
        """Удалить выбранную строку."""
        if not self.current_sheet:
            messagebox.showwarning("Предупреждение", "Нет активного листа")
            return
        
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите строку для удаления")
            return
        
        # Получаем номер строки
        item = selection[0]
        row_index = self.tree.index(item) + 1  # +1 потому что индексация с 1
        
        if not messagebox.askyesno("Подтверждение", f"Удалить строку {row_index}?"):
            return
        
        try:
            self.current_sheet.delete_row(row_index)
            self.refresh_table()
            self.update_status(f"Строка {row_index} удалена")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
    
    def sort_column_dialog(self):
        """Диалог сортировки по столбцу."""
        if not self.current_sheet:
            messagebox.showwarning("Предупреждение", "Нет активного листа")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Сортировка")
        dialog.geometry("300x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Номер столбца:").pack(pady=5)
        col_entry = ttk.Entry(dialog, width=10)
        col_entry.pack(pady=5)
        col_entry.insert(0, "1")
        
        order_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(dialog, text="По возрастанию", variable=order_var).pack(pady=5)
        
        ttk.Label(dialog, text="Начиная со строки:").pack(pady=(10, 0))
        start_row_entry = ttk.Entry(dialog, width=10)
        start_row_entry.pack(pady=5)
        start_row_entry.insert(0, "2")  # Пропускаем заголовок
        
        def on_sort():
            try:
                col = int(col_entry.get())
                start_row = int(start_row_entry.get())
                
                if col < 1 or col > self.current_sheet.max_column:
                    raise ValueError("Неверный номер столбца")
                
                self.current_sheet.sort_by_column(
                    col,
                    ascending=order_var.get(),
                    min_row=start_row
                )
                self.refresh_table()
                self.update_status(f"Сортировка по столбцу {col}")
                dialog.destroy()
            except ValueError as e:
                messagebox.showerror("Ошибка", str(e), parent=dialog)
        
        ttk.Button(dialog, text="Сортировать", command=on_sort).pack(pady=10)
    
    def apply_header_style(self):
        """Применить стиль заголовка к первой строке."""
        if not self.current_sheet:
            return
        
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_alignment = Alignment(horizontal='center', vertical='center')
        
        self.current_sheet.set_row_style(
            1,
            font=header_font,
            fill=header_fill,
            alignment=header_alignment
        )
        
        self.refresh_table()
        self.update_status("Стиль заголовка применен")
    
    def clear_row_style(self):
        """Очистить стили выбранной строки."""
        if not self.current_sheet:
            return
        
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите строку")
            return
        
        row_index = self.tree.index(selection[0]) + 1
        
        default_font = Font(name='Calibri', size=11)
        default_alignment = Alignment(horizontal='left', vertical='center')
        
        self.current_sheet.set_row_style(
            row_index,
            font=default_font,
            alignment=default_alignment
        )
        
        self.refresh_table()
        self.update_status(f"Стили строки {row_index} сброшены")


def main():
    """Запуск приложения."""
    root = tk.Tk()
    app = SpectrumExcelApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
