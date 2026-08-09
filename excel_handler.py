"""
Модуль для работы с Excel файлами через openpyxl.
Предоставляет классы для управления рабочими книгами и листами с блокировкой файлов.
"""

import os
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, List, Optional, Tuple, Union

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


class CellValidationError(Exception):
    """Исключение при ошибке валидации ячейки."""
    pass


class RowValidationError(Exception):
    """Исключение при ошибке валидации строки."""
    pass


class ExcelFileLock:
    """
    Класс для блокировки Excel файла на уровне ОС.
    Создает .lock файл с информацией о процессе, который удерживает файл.
    Также пытается захватить эксклюзивную блокировку черезfcntl/msvcrt.
    """
    
    def __init__(self, file_path: str):
        self.file_path = Path(file_path).resolve()
        self.lock_file_path = self.file_path.with_suffix(self.file_path.suffix + '.lock')
        self.lock_info = {
            'pid': os.getpid(),
            'user': os.environ.get('USER', 'unknown'),
            'timestamp': datetime.now().isoformat()
        }
        self._locked = False
        self._lock_handle = None  # Дескриптор lock-файла для удержания блокировки
    
    def acquire(self) -> bool:
        """
        Попытка захватить блокировку файла.
        Создает .lock файл и удерживает его открытым.
        Возвращает True если блокировка успешна, False если файл уже заблокирован.
        """
        # Проверяем наличие lock-файла
        if self.lock_file_path.exists():
            # Проверяем, не stale ли lock файл
            try:
                with open(self.lock_file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                # Если процесс больше не существует, можно удалить lock
                try:
                    import psutil
                    locked_pid = int(content.split('\n')[0].split(':')[1].strip())
                    if not psutil.pid_exists(locked_pid):
                        self.lock_file_path.unlink()
                    else:
                        return False
                except ImportError:
                    # Если psutil не установлен, просто считаем файл заблокированным
                    return False
            except (FileNotFoundError, ValueError, IndexError):
                pass
        
        try:
            # Создаем/открываем lock-файл и удерживаем его открытым
            # Это предотвращает удаление файла пока мы его держим
            self._lock_handle = open(self.lock_file_path, 'w', encoding='utf-8')
            self._lock_handle.write(f"PID: {self.lock_info['pid']}\n")
            self._lock_handle.write(f"User: {self.lock_info['user']}\n")
            self._lock_handle.write(f"Timestamp: {self.lock_info['timestamp']}\n")
            self._lock_handle.flush()
            
            # Пытаемся захватить блокировку на lock-файле
            if os.name == 'nt':
                try:
                    import msvcrt
                    # Блокируем lock-файл
                    msvcrt.locking(self._lock_handle.fileno(), msvcrt.LK_NBLCK, 1)
                except (ImportError, OSError):
                    pass
            else:
                try:
                    import fcntl
                    fcntl.flock(self._lock_handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                except (ImportError, BlockingIOError):
                    pass
            
            self._locked = True
            return True
            
        except PermissionError:
            # Файл уже заблокирован другим процессом
            return False
        except Exception:
            if self._lock_handle:
                self._lock_handle.close()
                self._lock_handle = None
            return False
    
    def release(self):
        """Освободить блокировку файла."""
        if self._lock_handle:
            try:
                if os.name == 'nt':
                    try:
                        import msvcrt
                        msvcrt.locking(self._lock_handle.fileno(), msvcrt.LK_UNLCK, 1)
                    except:
                        pass
                else:
                    try:
                        import fcntl
                        fcntl.flock(self._lock_handle.fileno(), fcntl.LOCK_UN)
                    except:
                        pass
                self._lock_handle.close()
            except Exception:
                pass
            self._lock_handle = None
        
        # Удаляем lock-файл
        if self._locked and self.lock_file_path.exists():
            try:
                self.lock_file_path.unlink()
            except Exception:
                pass
        self._locked = False
    
    def __enter__(self):
        if not self.acquire():
            raise FileExistsError(f"Файл {self.file_path} уже заблокирован другим процессом")
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.release()


class ExcelWorksheet:
    """
    Класс для работы с отдельным листом Excel.
    Предоставляет методы для манипуляции строками и стилями.
    """
    
    def __init__(self, worksheet):
        self._worksheet = worksheet
        self._default_font = Font(name='Calibri', size=11)
        self._default_alignment = Alignment(horizontal='left', vertical='center')
        self._default_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
    
    @property
    def name(self) -> str:
        """Вернуть имя листа."""
        return self._worksheet.title
    
    @name.setter
    def name(self, value: str):
        """Установить имя листа."""
        self._worksheet.title = value
    
    @property
    def max_row(self) -> int:
        """Вернуть номер последней строки с данными."""
        return self._worksheet.max_row
    
    @property
    def max_column(self) -> int:
        """Вернуть номер последнего столбца с данными."""
        return self._worksheet.max_column
    
    def create_sheet(self, name: str, index: Optional[int] = None) -> 'ExcelWorksheet':
        """
        Создать новый лист в книге.
        
        Args:
            name: Имя нового листа
            index: Позиция листа (None - добавить в конец)
        
        Returns:
            ExcelWorksheet: Новый лист
        """
        wb = self._worksheet.parent
        new_ws = wb.create_sheet(title=name, index=index)
        return ExcelWorksheet(new_ws)
    
    def set_active_sheet(self, sheet_name: str) -> 'ExcelWorksheet':
        """
        Сделать указанный лист активным.
        
        Args:
            sheet_name: Имя листа
        
        Returns:
            ExcelWorksheet: Активный лист
        """
        wb = self._worksheet.parent
        if sheet_name in wb.sheetnames:
            wb.active = wb[sheet_name]
            return ExcelWorksheet(wb.active)
        raise ValueError(f"Лист '{sheet_name}' не найден")
    
    def get_active_sheet(self) -> 'ExcelWorksheet':
        """Вернуть активный лист."""
        wb = self._worksheet.parent
        return ExcelWorksheet(wb.active)
    
    def validate_cell(self, value: Any, expected_type: Optional[type] = None) -> bool:
        """
        Валидировать значение ячейки.
        
        Args:
            value: Значение для проверки
            expected_type: Ожидаемый тип данных (None - любой тип)
        
        Returns:
            bool: True если валидация успешна
        """
        if value is None:
            return True
        
        if expected_type is not None:
            if not isinstance(value, expected_type):
                # Попытка неявного приведения типов
                try:
                    if expected_type == int:
                        if isinstance(value, (int, float)) and not isinstance(value, bool):
                            return True
                        float(value)
                        return True
                    elif expected_type == float:
                        float(value)
                        return True
                    elif expected_type == str:
                        return True
                    elif expected_type == bool:
                        return isinstance(value, bool) or value in ('True', 'False', 'true', 'false', 0, 1)
                except (ValueError, TypeError):
                    return False
                return True
            return True
        
        return True
    
    def validate_row(self, row_data: List[Any], min_cells: Optional[int] = None, 
                     max_cells: Optional[int] = None, 
                     cell_types: Optional[List[type]] = None) -> Tuple[bool, str]:
        """
        Валидировать строку данных.
        
        Args:
            row_data: Список значений ячеек
            min_cells: Минимальное количество ячеек
            max_cells: Максимальное количество ячеек
            cell_types: Список ожидаемых типов для каждой ячейки
        
        Returns:
            Tuple[bool, str]: (успех, сообщение об ошибке)
        """
        if not row_data:
            return False, "Строка пуста"
        
        if min_cells is not None and len(row_data) < min_cells:
            return False, f"Слишком мало ячеек: {len(row_data)} < {min_cells}"
        
        if max_cells is not None and len(row_data) > max_cells:
            return False, f"Слишком много ячеек: {len(row_data)} > {max_cells}"
        
        if cell_types is not None:
            for i, (value, expected_type) in enumerate(zip(row_data, cell_types)):
                if not self.validate_cell(value, expected_type):
                    return False, f"Ячейка {i+1}: неверный тип данных"
        
        return True, ""
    
    def add_row(self, row_data: List[Any], start_column: int = 1, 
                validate: bool = True, cell_types: Optional[List[type]] = None) -> int:
        """
        Добавить строку переменной длины.
        
        Args:
            row_data: Данные для строки
            start_column: Начальный столбец
            validate: Выполнять валидацию
            cell_types: Типы данных для валидации
        
        Returns:
            int: Номер добавленной строки
        
        Raises:
            RowValidationError: Если валидация не пройдена
        """
        if validate:
            is_valid, error_msg = self.validate_row(row_data, cell_types=cell_types)
            if not is_valid:
                raise RowValidationError(error_msg)
        
        row_num = self._worksheet.max_row + 1
        for col_idx, value in enumerate(row_data, start=start_column):
            cell = self._worksheet.cell(row=row_num, column=col_idx, value=value)
            cell.font = self._default_font
            cell.alignment = self._default_alignment
            cell.border = self._default_border
        
        return row_num
    
    def delete_row(self, row_num: int):
        """
        Удалить строку по номеру.
        
        Args:
            row_num: Номер строки для удаления
        """
        self._worksheet.delete_rows(row_num)
    
    def set_row_style(self, row_num: int, font: Optional[Font] = None,
                      fill: Optional[PatternFill] = None,
                      alignment: Optional[Alignment] = None,
                      border: Optional[Border] = None):
        """
        Изменить стиль всей строки.
        
        Args:
            row_num: Номер строки
            font: Шрифт
            fill: Заполнение
            alignment: Выравнивание
            border: Границы
        """
        for col in range(1, self._worksheet.max_column + 1):
            cell = self._worksheet.cell(row=row_num, column=col)
            if font:
                cell.font = font
            if fill:
                cell.fill = fill
            if alignment:
                cell.alignment = alignment
            if border:
                cell.border = border
    
    def sort_by_column(self, column: Union[int, str], ascending: bool = True,
                       min_row: int = 2, max_row: Optional[int] = None):
        """
        Сортировать строки по значению в заданном столбце.
        
        Args:
            column: Номер или буква столбца
            ascending: По возрастанию (True) или убыванию (False)
            min_row: Первая строка диапазона сортировки (обычно 2, если есть заголовок)
            max_row: Последняя строка диапазона (None - до конца)
        """
        if isinstance(column, str):
            from openpyxl.utils import column_index_from_string
            col_idx = column_index_from_string(column.upper())
        else:
            col_idx = column
        
        if max_row is None:
            max_row = self._worksheet.max_row
        
        # Собираем данные для сортировки
        rows_data = []
        for row in range(min_row, max_row + 1):
            row_values = []
            for col in range(1, self._worksheet.max_column + 1):
                cell = self._worksheet.cell(row=row, column=col)
                # Сохраняем только значения, стили восстановим из существующих ячеек
                row_values.append(cell.value)
            rows_data.append(row_values)
        
        # Сортируем
        def get_sort_key(row):
            value = row[col_idx - 1]
            # Обрабатываем None и разные типы
            if value is None:
                return (2, '', 0)  # None значения в конец
            if isinstance(value, bool):
                return (1, '', int(value))
            if isinstance(value, (int, float)):
                return (0, '', value)
            try:
                # Пытаемся привести к числу
                num_value = float(value)
                return (0, '', num_value)
            except (ValueError, TypeError):
                # Строковое значение
                return (1, str(value), 0)
        
        rows_data.sort(key=get_sort_key, reverse=not ascending)
        
        # Записываем обратно только значения
        for row_idx, row_values in enumerate(rows_data, start=min_row):
            for col_idx, value in enumerate(row_values, start=1):
                cell = self._worksheet.cell(row=row_idx, column=col_idx, value=value)


class ExcelWorkbook:
    """
    Основной класс для работы с Excel файлом.
    Открывает файл с блокировкой, предоставляет доступ к листам.
    """
    
    def __init__(self, file_path: Optional[str] = None, create_new: bool = False):
        """
        Конструктор рабочей книги.
        
        Args:
            file_path: Путь к файлу (None для создания новой книги)
            create_new: Создать новый файл даже если file_path указан
        """
        self._file_path: Optional[Path] = None
        self._workbook: Optional[Workbook] = None
        self._lock: Optional[ExcelFileLock] = None
        self._active_sheet: Optional[ExcelWorksheet] = None
        
        if file_path:
            self.open(file_path, create_new=create_new)
        else:
            self._workbook = Workbook()
            self._active_sheet = ExcelWorksheet(self._workbook.active)
    
    def open(self, file_path: str, create_new: bool = False) -> None:
        """
        Открыть существующий файл или создать новый.
        
        Args:
            file_path: Путь к файлу
            create_new: Создать новый файл
        
        Raises:
            FileNotFoundError: Файл не найден и create_new=False
            FileExistsError: Файл заблокирован другим процессом
            ValueError: Некорректный файл Excel
        """
        self._file_path = Path(file_path).resolve()
        
        # Если файл не существует и не создаем новый - ошибка
        if not self._file_path.exists():
            if not create_new:
                raise FileNotFoundError(f"Файл {self._file_path} не найден")
            # Создаем новую книгу и сохраняем
            self._workbook = Workbook()
            self._workbook.save(str(self._file_path))
            self._active_sheet = ExcelWorksheet(self._workbook.active)
            return
        
        # Блокируем файл через lock-файл
        self._lock = ExcelFileLock(str(self._file_path))
        if not self._lock.acquire():
            raise FileExistsError(
                f"Файл {self._file_path} заблокирован другим процессом"
            )
        
        # Загружаем книгу
        try:
            self._workbook = load_workbook(str(self._file_path))
            self._active_sheet = ExcelWorksheet(self._workbook.active)
        except Exception as e:
            # Освобождаем блокировку при ошибке загрузки
            if self._lock:
                self._lock.release()
                self._lock = None
            raise ValueError(f"Ошибка загрузки файла: {str(e)}")
    
    def save(self, file_path: Optional[str] = None) -> None:
        """
        Сохранить файл.
        
        Args:
            file_path: Путь для сохранения (None - сохранить в текущий файл)
        """
        if file_path:
            save_path = Path(file_path).resolve()
        else:
            if not self._file_path:
                raise ValueError("Не указан путь для сохранения")
            save_path = self._file_path
        
        # Освобождаем старую блокировку если путь изменился
        if self._lock and save_path != self._file_path:
            self._lock.release()
            self._lock = ExcelFileLock(str(save_path))
            if not self._lock.acquire():
                raise FileExistsError(
                    f"Файл {save_path} заблокирован другим процессом"
                )
            self._file_path = save_path
        
        self._workbook.save(str(save_path))
    
    def close(self) -> None:
        """Закрыть файл и освободить блокировку."""
        if self._lock:
            self._lock.release()
            self._lock = None
        self._workbook = None
        self._active_sheet = None
        self._file_path = None
    
    def create_sheet(self, name: str, index: Optional[int] = None) -> ExcelWorksheet:
        """
        Создать новый лист.
        
        Args:
            name: Имя листа
            index: Позиция (None - в конец)
        
        Returns:
            ExcelWorksheet: Созданный лист
        """
        if not self._workbook:
            raise RuntimeError("Книга не открыта")
        
        new_ws = self._workbook.create_sheet(title=name, index=index)
        return ExcelWorksheet(new_ws)
    
    def set_active_sheet(self, sheet_name: str) -> ExcelWorksheet:
        """
        Сделать лист активным.
        
        Args:
            sheet_name: Имя листа
        
        Returns:
            ExcelWorksheet: Активный лист
        """
        if not self._workbook:
            raise RuntimeError("Книга не открыта")
        
        if sheet_name not in self._workbook.sheetnames:
            raise ValueError(f"Лист '{sheet_name}' не найден")
        
        self._workbook.active = self._workbook[sheet_name]
        self._active_sheet = ExcelWorksheet(self._workbook.active)
        return self._active_sheet
    
    def get_active_sheet(self) -> ExcelWorksheet:
        """Вернуть активный лист."""
        if not self._active_sheet:
            raise RuntimeError("Книга не открыта")
        return self._active_sheet
    
    def get_sheet(self, name_or_index: Union[str, int]) -> ExcelWorksheet:
        """
        Получить лист по имени или индексу.
        
        Args:
            name_or_index: Имя или индекс листа
        
        Returns:
            ExcelWorksheet: Лист
        """
        if not self._workbook:
            raise RuntimeError("Книга не открыта")
        
        if isinstance(name_or_index, str):
            ws = self._workbook[name_or_index]
        else:
            ws = self._workbook.worksheets[name_or_index]
        
        return ExcelWorksheet(ws)
    
    @property
    def sheetnames(self) -> List[str]:
        """Вернуть список имен листов."""
        if not self._workbook:
            return []
        return self._workbook.sheetnames
    
    @property
    def active_sheet(self) -> ExcelWorksheet:
        """Вернуть активный лист."""
        if not self._active_sheet:
            raise ValueError("Нет активного листа. Откройте файл или создайте новый.")
        return self._active_sheet
    
    @property
    def file_path(self) -> Optional[str]:
        """Вернуть путь к файлу."""
        return str(self._file_path) if self._file_path else None
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()


# Пример использования
if __name__ == "__main__":
    # Создаем тестовый файл
    test_file = "test_spectrum_data.xlsx"
    
    with ExcelWorkbook(test_file, create_new=True) as wb:
        # Создаем лист для данных
        data_sheet = wb.create_sheet("Spectrum Data")
        wb.set_active_sheet("Spectrum Data")
        
        # Добавляем заголовок
        header = ["ID", "Wavelength", "Intensity", "Date", "Notes"]
        data_sheet.add_row(header)
        
        # Добавляем данные
        data_sheet.add_row([1, 450.5, 1234.56, "2024-01-01", "Sample 1"])
        data_sheet.add_row([2, 460.2, 2345.67, "2024-01-02", "Sample 2"])
        data_sheet.add_row([3, 470.8, 3456.78, "2024-01-03", "Sample 3"])
        
        # Сортируем по интенсивности
        data_sheet.sort_by_column(3, ascending=False)
        
        # Применяем стиль к заголовку
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        data_sheet.set_row_style(1, font=header_font, fill=header_fill)
        
        wb.save()
    
    print(f"Тестовый файл создан: {test_file}")
    print("Открываем снова для чтения...")
    
    with ExcelWorkbook(test_file) as wb:
        sheet = wb.get_active_sheet()
        print(f"Лист: {sheet.name}")
        print(f"Всего строк: {sheet.max_row}")
        print(f"Всего столбцов: {sheet.max_column}")
