"""
Процессор очистки файлов - удаление апострофов и замена D на точку
ИСПРАВЛЕНО: Точный поиск столбцов и замена D→. в нужных столбцах
"""

import pandas as pd
import numpy as np
from pathlib import Path
from core.base_processor import BaseProcessor, ProcessingResult

class FileCleaner(BaseProcessor):
    """Очистка файлов от апострофов и некорректных символов"""
    
    def get_name(self) -> str:
        return "Очистка файлов"
    
    def get_description(self) -> str:
        return "Удаление апострофов и замена 'D' на '.' в столбцах Price (AB), Net Weight (AJ), Gross Weight (AK)"
    
    def process_file(self, input_path: Path, **kwargs) -> ProcessingResult:
        """Очистка файла"""
        try:
            if not self.validate_input_file(input_path):
                return ProcessingResult(
                    success=False, 
                    input_file=input_path,
                    message="Ошибка валидации входного файла"
                )
            
            # Чтение файла
            df = pd.read_excel(input_path, engine=self.settings.excel_engine)
            
            print(f"📊 Размер файла: {len(df)} строк, {len(df.columns)} столбцов")
            
            # Статистика изменений
            changes = {
                "apostrophes_removed": 0,
                "d_to_dot_replacements": 0,
                "processed_cells": 0,
                "target_columns_found": []
            }
            
            # 1. Удаление апострофов из всех ячеек
            print("🧹 Удаление апострофов из всех ячеек...")
            for col_idx, col in enumerate(df.columns):
                for row_idx in df.index:
                    cell_value = df.iloc[row_idx, col_idx]
                    if isinstance(cell_value, str) and cell_value.startswith("'"):
                        df.iloc[row_idx, col_idx] = cell_value[1:]  # Удаляем первый символ (апостроф)
                        changes["apostrophes_removed"] += 1
                        changes["processed_cells"] += 1
            
            print(f"   ✓ Удалено апострофов: {changes['apostrophes_removed']}")
            
            # 2. Поиск и обработка целевых столбцов для замены D→.
            print("🔄 Поиск столбцов Price, Net Weight, Gross Weight...")
            
            target_columns = self._find_target_columns(df)
            changes["target_columns_found"] = target_columns
            
            if target_columns:
                print(f"   ✓ Найдено целевых столбцов: {len(target_columns)}")
                for col_info in target_columns:
                    print(f"      • {col_info['name']} (столбец {col_info['letter']}, индекс {col_info['index']})")
            else:
                print("   ⚠️  Целевые столбцы не найдены, используем стандартные индексы")
                # Используем стандартные индексы столбцов AB, AJ, AK
                target_columns = self._get_default_columns(df)
            
            # 3. Замена D→. в найденных столбцах
            print("🔄 Замена 'D' на '.' в числовых столбцах...")
            
            for col_info in target_columns:
                col_idx = col_info['index']
                col_name = col_info['name']
                
                if col_idx < len(df.columns):
                    print(f"   Обрабатываем: {col_name} (индекс {col_idx})")
                    
                    for row_idx in df.index:
                        cell_value = df.iloc[row_idx, col_idx]
                        
                        # Проверяем, что это строка и содержит 'D'
                        if isinstance(cell_value, str) and 'D' in cell_value:
                            old_value = cell_value
                            new_value = cell_value.replace('D', '.')
                            df.iloc[row_idx, col_idx] = new_value
                            changes["d_to_dot_replacements"] += 1
                            changes["processed_cells"] += 1
                            
                            # Логируем первые несколько замен для проверки
                            if changes["d_to_dot_replacements"] <= 5:
                                print(f"      Замена: '{old_value}' → '{new_value}'")
                
                else:
                    print(f"   ⚠️  Столбец {col_name} (индекс {col_idx}) не существует в файле")
            
            print(f"   ✓ Выполнено замен D→.: {changes['d_to_dot_replacements']}")
            
            # Сохранение результата
            output_path = self.file_manager.generate_output_path(input_path, "_cleaned")
            df.to_excel(output_path, index=False, engine=self.settings.excel_engine)
            
            # Архивирование исходного файла
            self.file_manager.archive_file(input_path)
            
            result_message = (f"Апострофы: {changes['apostrophes_removed']}, "
                            f"D→.: {changes['d_to_dot_replacements']} "
                            f"(столбцы: {len(changes['target_columns_found'])})")
            
            return ProcessingResult(
                success=True,
                input_file=input_path,
                output_file=output_path,
                message=result_message,
                details=changes
            )
            
        except Exception as e:
            error_msg = f"Ошибка очистки: {str(e)}"
            self.logger.error(f"Ошибка обработки {input_path}: {e}")
            return ProcessingResult(success=False, input_file=input_path, message=error_msg)
    
    def _find_target_columns(self, df: pd.DataFrame) -> list:
        """Поиск целевых столбцов Price, Net Weight, Gross Weight"""
        target_columns = []
        
        # Список ключевых слов для поиска столбцов
        search_patterns = {
            'price': ['price', 'стоимость', 'цена'],
            'net_weight': ['net weight', 'nett weight', 'вес нетто', 'нетто'],
            'gross_weight': ['gross weight', 'вес брутто', 'брутто']
        }
        
        # Поиск по названиям столбцов
        for col_idx, col_name in enumerate(df.columns):
            col_name_lower = str(col_name).lower().strip()
            
            for pattern_type, patterns in search_patterns.items():
                for pattern in patterns:
                    if pattern in col_name_lower:
                        target_columns.append({
                            'name': str(col_name),
                            'index': col_idx,
                            'letter': self._get_excel_column_letter(col_idx),
                            'type': pattern_type
                        })
                        break
                if any(pattern in col_name_lower for pattern in patterns):
                    break
        
        # Удаляем дубликаты по индексу
        unique_columns = []
        seen_indices = set()
        for col in target_columns:
            if col['index'] not in seen_indices:
                unique_columns.append(col)
                seen_indices.add(col['index'])
        
        return unique_columns
    
    def _get_default_columns(self, df: pd.DataFrame) -> list:
        """Получить столбцы по стандартным индексам AB(27), AJ(35), AK(36)"""
        default_indices = [
            {'index': 27, 'letter': 'AB', 'name': 'Price (AB)', 'type': 'price'},           # AB
            {'index': 35, 'letter': 'AJ', 'name': 'Net Weight (AJ)', 'type': 'net_weight'}, # AJ  
            {'index': 36, 'letter': 'AK', 'name': 'Gross Weight (AK)', 'type': 'gross_weight'} # AK
        ]
        
        # Проверяем, существуют ли эти столбцы в файле
        available_columns = []
        for col_info in default_indices:
            if col_info['index'] < len(df.columns):
                available_columns.append(col_info)
        
        return available_columns
    
    def _get_excel_column_letter(self, col_idx: int) -> str:
        """Преобразование индекса столбца в букву Excel (0->A, 1->B, 27->AB)"""
        result = ""
        col_idx += 1  # Excel считает с 1
        
        while col_idx > 0:
            col_idx -= 1
            result = chr(65 + col_idx % 26) + result
            col_idx //= 26
        
        return result
    
    def _debug_columns_info(self, df: pd.DataFrame) -> None:
        """Отладочная информация о столбцах (для разработки)"""
        print("\n📋 ОТЛАДКА - Информация о столбцах:")
        
        # Показываем столбцы в районе AB, AJ, AK
        target_indices = [25, 26, 27, 28, 33, 34, 35, 36, 37]
        
        for idx in target_indices:
            if idx < len(df.columns):
                col_name = df.columns[idx]
                letter = self._get_excel_column_letter(idx)
                print(f"   {letter} (индекс {idx}): '{col_name}'")
        
        print("=" * 50)
