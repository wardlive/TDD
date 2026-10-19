"""
Процессор сжатия спецификаций - объединение v2.2 и v2.3
"""

import pandas as pd
from pathlib import Path
from core.base_processor import BaseProcessor, ProcessingResult

class SpecificationCompressor(BaseProcessor):
    """Сжатие спецификаций с опциональным расчетом долей"""
    
    def get_name(self) -> str:
        return "Сжатие спецификаций"
    
    def get_description(self) -> str:
        return "Группировка товаров по артикулам с опциональным расчетом долей в коробках"
    
    def process_file(self, input_path: Path, **kwargs) -> ProcessingResult:
        """Обработка файла спецификации"""
        calculate_percentages = kwargs.get('calculate_percentages', False)
        
        try:
            if not self.validate_input_file(input_path):
                return ProcessingResult(
                    success=False, 
                    input_file=input_path,
                    message="Ошибка валидации входного файла"
                )
            
            # Чтение файла
            df = pd.read_excel(input_path, header=1, engine="openpyxl")
            if len(df.columns) < 14:
                raise ValueError("Для сжатия нужны столбцы A–N, заголовки в строке 2")
            if df.iloc[:, 0].isna().any():
                raise ValueError("В A есть строки без артикула: сжатие могло бы потерять эти данные")
            
            original_rows = len(df)
            
            # Опциональный расчет долей в коробках
            if calculate_percentages:
                df = self._calculate_box_percentages(df)
                
                # Промежуточное сохранение
                self.file_manager.save_intermediate(
                    df, input_path, "_with_percentages"
                )
            
            # Основное сжатие
            compressed_df = self._compress_specifications(df, calculate_percentages)
            
            compressed_rows = len(compressed_df)
            compression_ratio = original_rows / compressed_rows if compressed_rows > 0 else 1
            
            # Сохранение результата
            suffix = "_compressed_with_percentages" if calculate_percentages else "_compressed"
            output_path = self.file_manager.generate_output_path(input_path, suffix)
            with self.file_manager.atomic_output(output_path) as temporary_path:
                compressed_df.to_excel(temporary_path, index=False, engine=self.settings.excel_engine)
            
            # Архивирование исходного файла
            self.file_manager.archive_file(input_path)
            
            result_message = f"{original_rows} → {compressed_rows} строк (сжатие {compression_ratio:.1f}x)"
            
            return ProcessingResult(
                success=True,
                input_file=input_path,
                output_file=output_path,
                message=result_message,
                details={
                    "original_rows": original_rows,
                    "compressed_rows": compressed_rows,
                    "compression_ratio": compression_ratio,
                    "with_percentages": calculate_percentages
                }
            )
            
        except Exception as e:
            error_msg = f"Ошибка сжатия: {str(e)}"
            self.logger.error(f"Ошибка обработки {input_path}: {e}")
            return ProcessingResult(success=False, input_file=input_path, message=error_msg)
    
    def _calculate_box_percentages(self, df: pd.DataFrame) -> pd.DataFrame:
        """Расчет долей товаров в коробках"""
        if len(df.columns) < 22:
            raise ValueError("Для расчёта долей нужны вес нетто в N и номер коробки в V")
        df = df.copy()
        weights = pd.to_numeric(df.iloc[:, 13], errors="raise")
        boxes = df.iloc[:, 21]
        if weights.isna().any() or (weights < 0).any() or boxes.isna().any():
            raise ValueError("Для каждой строки нужны неотрицательный вес в N и номер коробки в V")
        totals = weights.groupby(boxes).transform("sum")
        if (totals <= 0).any():
            raise ValueError("Суммарный вес каждой коробки должен быть больше нуля")
        shares = weights / totals
        df.rename(columns={df.columns[14]: "Доля в % в коробке", df.columns[15]: "Доля округленная"}, inplace=True)
        df[df.columns[14]] = shares
        df[df.columns[15]] = shares.round(2)
        if not shares.groupby(boxes).sum().sub(1).abs().le(0.01).all():
            raise ValueError("Сумма долей в коробке отличается от единицы")
        return df
    
    def _compress_specifications(self, df: pd.DataFrame, with_percentages: bool = False) -> pd.DataFrame:
        """Основное сжатие спецификаций"""
        if df.empty:
            return df.copy()
        grouped_data = []
        
        for article, group in df.groupby(df.iloc[:, 0]):  # Группировка по артикулу
            # Берем первую строку как основу
            base_row = group.iloc[0].copy()
            
            # Суммируем ключевые показатели
            base_row.iloc[10] = group.iloc[:, 10].sum()  # Количество
            base_row.iloc[13] = group.iloc[:, 13].sum()  # Вес нетто
            
            if with_percentages and len(df.columns) > 15:
                # Суммируем доли (новые столбцы O и P)
                base_row.iloc[14] = group.iloc[:, 14].sum()  # Доли в коробках
                base_row.iloc[15] = round(group.iloc[:, 14].sum(), 2)  # Округленные доли
                
            brutto_col, cost_col, box_col = 16, 19, 21
            
            # Суммируем остальные показатели
            if len(df.columns) > brutto_col:
                base_row.iloc[brutto_col] = group.iloc[:, brutto_col].sum()  # Вес брутто
            if len(df.columns) > cost_col:
                base_row.iloc[cost_col] = group.iloc[:, cost_col].sum()  # Стоимость
            
            # Создаем диапазон размеров
            sizes_data = group.iloc[:, 11].tolist()  # Столбец размеров
            base_row.iloc[11] = self._create_size_range(sizes_data)
            
            # Объединяем номера коробок
            if len(df.columns) > box_col:
                boxes = group.iloc[:, box_col].dropna().astype(str).unique()
                base_row.iloc[box_col] = ','.join(sorted(boxes))
            
            grouped_data.append(base_row)
        
        return pd.DataFrame(grouped_data)
    
    def _create_size_range(self, sizes_list) -> str:
        """Создание диапазона размеров"""
        try:
            # Убираем пустые значения
            sizes = [str(s).strip().upper() for s in sizes_list if pd.notna(s) and str(s).strip()]
            
            if not sizes:
                return ""
            
            # Разделяем на числовые и буквенные
            numeric_sizes = []
            letter_sizes = []
            
            for size in sizes:
                try:
                    numeric_sizes.append(float(size.replace(',', '.')))
                except ValueError:
                    letter_sizes.append(size)
            
            # Создаем диапазоны
            parts = []
            
            if numeric_sizes:
                numeric_sizes = sorted(set(numeric_sizes))
                if len(numeric_sizes) > 1:
                    parts.append(f"{numeric_sizes[0]:g}-{numeric_sizes[-1]:g}")
                else:
                    parts.append(f"{numeric_sizes[0]:g}")
            
            if letter_sizes:
                letter_order = ["XS", "S", "M", "L", "XL", "XXL", "XXXL"]
                letter_sizes = sorted(set(letter_sizes), 
                                    key=lambda x: (letter_order.index(x) if x in letter_order else 999, x))
                if len(letter_sizes) > 1:
                    parts.append(f"{letter_sizes[0]}-{letter_sizes[-1]}")
                else:
                    parts.append(letter_sizes[0])
            
            return "; ".join(parts)
            
        except Exception as e:
            self.logger.error(f"Ошибка создания диапазона размеров: {e}")
            return ", ".join(str(s) for s in sizes_list if pd.notna(s))
