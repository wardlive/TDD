"""
Процессор сжатия спецификаций - объединение v2.2 и v2.3
"""

import pandas as pd
from pathlib import Path
from typing import Dict, Any
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
            df = pd.read_excel(input_path, header=1, engine=self.settings.excel_engine)
            
            original_rows = len(df)
            unique_articles = df.iloc[:, 0].nunique()
            
            print(f"📊 Строк: {original_rows}, Столбцов: {len(df.columns)}")
            print(f"🏷  Уникальных артикулов: {unique_articles}")
            
            # Опциональный расчет долей в коробках
            if calculate_percentages:
                print("📦 Обработка долей в коробках...")
                df = self._calculate_box_percentages(df)
                
                # Промежуточное сохранение
                intermediate_path = self.file_manager.save_intermediate(
                    df, input_path, "_with_percentages"
                )
                print(f"💾 Промежуточный файл: {intermediate_path.name}")
            
            # Основное сжатие
            print("🔄 Сжатие данных...")
            compressed_df = self._compress_specifications(df, calculate_percentages)
            
            compressed_rows = len(compressed_df)
            compression_ratio = original_rows / compressed_rows if compressed_rows > 0 else 1
            
            # Сохранение результата
            suffix = "_compressed_with_percentages" if calculate_percentages else "_compressed"
            output_path = self.file_manager.generate_output_path(input_path, suffix)
            compressed_df.to_excel(output_path, index=False, engine=self.settings.excel_engine)
            
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
        try:
            # Проверяем и переименовываем столбец O (индекс 14)
            if len(df.columns) > 14:
                if df.columns[14] != 'Доля в % в коробке':
                    df.rename(columns={df.columns[14]: 'Доля в % в коробке'}, inplace=True)
            else:
                df.insert(14, 'Доля в % в коробке', 0.0)
            
            # Добавляем столбец P (индекс 15) если его нет
            if len(df.columns) <= 15:
                df.insert(15, 'Доля округленная', 0.0)
            
            # Инициализируем столбцы
            df.iloc[:, 14] = 0.0  # Столбец O
            df.iloc[:, 15] = 0.0  # Столбец P
            
            # Группируем по номеру коробки (столбец 21 или последний столбец с коробками)
            box_column = min(21, len(df.columns) - 1)
            netto_column = 13  # Столбец N (вес нетто)
            
            box_groups = df.groupby(df.iloc[:, box_column])
            validation_passed = True
            
            for box_num, group in box_groups:
                if pd.isna(box_num):
                    continue
                
                # Рассчитываем общий вес коробки
                total_weight = group.iloc[:, netto_column].sum()
                
                if total_weight > 0:
                    # Рассчитываем доли для каждого товара в коробке
                    for idx in group.index:
                        item_weight = df.iloc[idx, netto_column]
                        percentage = item_weight / total_weight
                        
                        # Записываем в столбцы O и P
                        df.iloc[idx, 14] = percentage
                        df.iloc[idx, 15] = round(percentage, 2)
                    
                    # Валидация суммы долей
                    sum_percentages = group.iloc[:, 14].sum()
                    if abs(sum_percentages - 1.0) > 0.01:
                        validation_passed = False
                        self.logger.warning(f"Коробка {box_num}: сумма долей = {sum_percentages:.3f}")
            
            if validation_passed:
                print("✓ Валидация долей успешна")
            else:
                print("⚠️ Валидация долей: есть предупреждения")
            
            return df
            
        except Exception as e:
            self.logger.error(f"Ошибка расчета долей: {e}")
            return df
    
    def _compress_specifications(self, df: pd.DataFrame, with_percentages: bool = False) -> pd.DataFrame:
        """Основное сжатие спецификаций"""
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
                
                # Сдвигаем индексы из-за новых столбцов
                brutto_col = 16 if len(df.columns) > 16 else 16
                cost_col = 19 if len(df.columns) > 19 else 19
                box_col = 21 if len(df.columns) > 21 else 21
            else:
                brutto_col = 16
                cost_col = 19
                box_col = 21
            
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
                    numeric_sizes.append(int(size))
                except ValueError:
                    letter_sizes.append(size)
            
            # Создаем диапазоны
            parts = []
            
            if numeric_sizes:
                numeric_sizes = sorted(set(numeric_sizes))
                if len(numeric_sizes) > 1:
                    parts.append(f"{numeric_sizes[0]}-{numeric_sizes[-1]}")
                else:
                    parts.append(str(numeric_sizes[0]))
            
            if letter_sizes:
                letter_order = ["XS", "S", "M", "L", "XL", "XXL", "XXXL"]
                letter_sizes = sorted(set(letter_sizes), 
                                    key=lambda x: letter_order.index(x) if x in letter_order else 999)
                if len(letter_sizes) > 1:
                    parts.append(f"{letter_sizes[0]}-{letter_sizes[-1]}")
                else:
                    parts.append(letter_sizes[0])
            
            return "; ".join(parts)
            
        except Exception as e:
            self.logger.error(f"Ошибка создания диапазона размеров: {e}")
            return ", ".join(str(s) for s in sizes_list if pd.notna(s))
