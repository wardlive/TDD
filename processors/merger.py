"""
Процессор объединения листов PL INV - полная реализация на основе рабочего кода
ИСПРАВЛЕНО: Точное соответствие логике оригинального merge.py
"""

import re
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from core.base_processor import BaseProcessor, ProcessingResult

class SheetMerger(BaseProcessor):
    """Объединение листов PL INV в сводный лист 'Общий'"""
    
    def __init__(self, file_manager, settings):
        super().__init__(file_manager, settings)
        
        # Константы из оригинального кода
        self.PATTERN_PL = re.compile(r'^PL\s+INV\s+', re.I)
        self.TARGET_SHEET = 'Общий'
        self.UNWANTED = (
            'Group:', 'Delivery:', 'Customer', '[', '350062', 'YARKAYA',
            'Total pieces', 'Boxes:', 'Gross weight', 'Net weight', 'Volum'
        )
        self.RE_UNWANTED = re.compile(
            r'^(?:' + '|'.join(map(re.escape, self.UNWANTED)) + r')', re.I
        )
        self.GREEN_FILL = PatternFill(start_color='C8E6C9', end_color='C8E6C9', fill_type='solid')
        self.BOLD_FONT = Font(bold=True)
    
    def get_name(self) -> str:
        return "Объединение листов PL INV"
    
    def get_description(self) -> str:
        return "Объединяет все листы 'PL INV ...' в сводный лист 'Общий' с форматированием"
    
    def process_file(self, input_path: Path, **kwargs) -> ProcessingResult:
        """Объединение листов PL INV в файле"""
        try:
            if not self.validate_input_file(input_path):
                return ProcessingResult(
                    success=False, 
                    input_file=input_path,
                    message="Ошибка валидации входного файла"
                )
            
            # Загружаем книгу Excel
            book = load_workbook(input_path, read_only=False)
            
            # Находим листы PL INV
            pl_sheets = [s for s in book.sheetnames if self._is_pl_inv(s)]
            
            if not pl_sheets:
                return ProcessingResult(
                    success=False,
                    input_file=input_path,
                    message="Не найдено листов с названием 'PL INV'"
                )
            
            print(f"📋 Файл: {input_path.name}")
            print(f"📋 Найдено {len(pl_sheets)} листов PL INV:")
            for sheet in pl_sheets:
                print(f"   • {sheet}")
            
            # Читаем и объединяем все листы PL INV
            merged_data = []
            total_rows = 0
            
            for sheet_name in pl_sheets:
                df = self._read_sheet(input_path, sheet_name)
                merged_data.append(df)
                total_rows += len(df)
                print(f"   • {sheet_name}: {len(df)} строк после очистки")
            
            # Объединяем все данные
            merged_df = pd.concat(merged_data, ignore_index=True)
            print(f"📊 Сводная таблица: {len(merged_df)} строк")
            
            # Создаем/обновляем лист "Общий"
            if self.TARGET_SHEET in book.sheetnames:
                book.remove(book[self.TARGET_SHEET])
            
            ws = book.create_sheet(self.TARGET_SHEET, 0)
            
            # Записываем данные
            ws.append([])  # пустая строка 1
            ws.append(list(merged_df.columns))  # строка 2 – заголовки
            
            for row in merged_df.itertuples(index=False):
                ws.append(list(row))
            
            # Применяем форматирование
            last_row = len(merged_df) + 2
            self._format_worksheet(ws, last_row)
            
            # Сохраняем файл
            output_path = self.file_manager.generate_output_path(input_path, "_merged")
            book.save(output_path)
            
            # Архивирование исходного файла
            self.file_manager.archive_file(input_path)
            
            result_message = f"Объединено {len(pl_sheets)} листов → {len(merged_df)} строк в лист '{self.TARGET_SHEET}'"
            
            return ProcessingResult(
                success=True,
                input_file=input_path,
                output_file=output_path,
                message=result_message,
                details={
                    "merged_sheets": len(pl_sheets),
                    "total_rows": len(merged_df),
                    "sheet_names": pl_sheets
                }
            )
            
        except Exception as e:
            error_msg = f"Ошибка объединения: {str(e)}"
            self.logger.error(f"Ошибка обработки {input_path}: {e}")
            return ProcessingResult(success=False, input_file=input_path, message=error_msg)
    
    def _is_pl_inv(self, sheet_name: str) -> bool:
        """Проверка, является ли лист листом PL INV"""
        return bool(self.PATTERN_PL.match(sheet_name.strip()))
    
    def _clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """Очистка DataFrame от ненужных строк"""
        # Удаляем полностью пустые строки
        df = df.dropna(how='all')
        
        # Создаем маску для удаления строк с нежелательными данными
        mask = df.apply(
            lambda row: row.fillna('').astype(str).str.strip()
                        .str.match(self.RE_UNWANTED).any(),
            axis=1
        )
        
        return df.loc[~mask].reset_index(drop=True)
    
    def _read_sheet(self, file_path: Path, sheet_name: str) -> pd.DataFrame:
        """Чтение и обработка одного листа PL INV"""
        # Читаем лист как строки для сохранения форматирования
        df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str, engine=self.settings.excel_engine)
        
        # Очищаем от ненужных строк
        df = self._clean_dataframe(df)
        
        # Добавляем столбец с номером инвойса в начало
        invoice_number = sheet_name.split('PL INV', 1)[1].strip() if 'PL INV' in sheet_name else sheet_name
        df.insert(0, 'Invoice', invoice_number)
        
        return df
    
    def _format_worksheet(self, ws, last_row: int) -> None:
        """Применение форматирования к рабочему листу"""
        try:
            # Форматирование числового столбца R (если существует)
            if 'R' <= get_column_letter(ws.max_column):
                self._format_numeric_column(ws, 'R', 3, last_row)
                
                # Добавляем формулу SUBTOTAL в R1
                ws['R1'] = f"=SUBTOTAL(9,R3:R{last_row})"
            
            # Применяем зеленую заливку к столбцам D и R
            for col_letter in ('D', 'R'):
                if col_letter <= get_column_letter(ws.max_column):
                    for row in range(1, last_row + 1):
                        cell = ws[f'{col_letter}{row}']
                        cell.fill = self.GREEN_FILL
            
            # Автоматическая настройка ширины столбцов и форматирование заголовков
            self._autosize_columns(ws)
            
            print("✓ Применено форматирование: зеленая заливка, жирные заголовки, автоширина")
            
        except Exception as e:
            self.logger.warning(f"Ошибка применения форматирования: {e}")
            print(f"⚠️ Частичная ошибка форматирования: {e}")
    
    def _format_numeric_column(self, ws, col_letter: str, start_row: int, end_row: int) -> None:
        """Форматирование числового столбца"""
        for row in range(start_row, end_row + 1):
            cell = ws[f'{col_letter}{row}']
            cell.number_format = '0'
            
            # Преобразование строковых значений в числа
            val = cell.value
            if isinstance(val, str) and val.strip():
                try:
                    cell.value = int(float(val.strip()))
                except (ValueError, TypeError):
                    pass  # Оставляем как есть, если не удается преобразовать
    
    def _autosize_columns(self, ws) -> None:
        """Автоматическая настройка ширины столбцов и форматирование заголовков"""
        for idx, col in enumerate(ws.columns, start=1):
            # Вычисляем максимальную длину содержимого в столбце
            max_length = 0
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            
            # Устанавливаем ширину столбца (минимум 8, максимум 50)
            width = min(max(max_length + 2, 8), 50)
            ws.column_dimensions[get_column_letter(idx)].width = width
            
            # Форматируем заголовок (строка 2)
            header_cell = ws.cell(row=2, column=idx)
            header_cell.font = self.BOLD_FONT
            header_cell.fill = self.GREEN_FILL
