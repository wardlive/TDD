#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Процессор заполнения артикулов - ИСПРАВЛЕННАЯ ВЕРСИЯ
Заголовки ВСЕГДА во второй строке (header=1)
"""

import pandas as pd
from pathlib import Path
from typing import Dict, Optional, Tuple
from core.base_processor import BaseProcessor, ProcessingResult

class ArticleFiller(BaseProcessor):
    """Заполнение столбцов артикулов и инвойсов на основе данных из эталонного файла"""
    
    def get_name(self) -> str:
        return "Заполнение артикулов"
    
    def get_description(self) -> str:
        return "Заполняет столбцы 'Инвойс' и 'Артикул' на основании сопоставления номеров коробок с эталонным файлом"
    
    def process_file(self, input_path: Path, **kwargs) -> ProcessingResult:
        """Заполнение артикулов в файле"""
        try:
            if not self.validate_input_file(input_path):
                return ProcessingResult(
                    success=False, 
                    input_file=input_path,
                    message="Ошибка валидации входного файла"
                )
            
            print(f"🔍 Обрабатываем файл: {input_path.name}")
            
            # Поиск эталонного файла
            reference_file = self._find_reference_file(input_path, kwargs)
            if not reference_file:
                return ProcessingResult(
                    success=False,
                    input_file=input_path,
                    message="Не найден эталонный файл для заполнения артикулов"
                )
            
            print(f"📋 Эталонный файл: {reference_file.name}")
            
            # Определяем, какой файл какой, по характерным столбцам
            weights_file, spec_file = self._identify_files(input_path, reference_file)
            
            # Чтение файлов с заголовками во второй строке
            weights_df = self._read_file_with_fixed_header(weights_file)
            spec_df = self._read_file_with_fixed_header(spec_file)
            
            print(f"📊 Файл весов: {len(weights_df)} строк, столбцы: {list(weights_df.columns)}")
            print(f"📊 Файл спецификации: {len(spec_df)} строк, столбцы: {list(spec_df.columns)}")
            
            # Создаем маппинги
            box_to_article_mapping, box_to_invoice_mapping = self._create_mappings(spec_df)
            
            if not box_to_article_mapping:
                return ProcessingResult(
                    success=False,
                    input_file=input_path,
                    message="Не удалось создать сопоставление коробок и артикулов"
                )
            
            print(f"✅ Найдено {len(box_to_article_mapping)} сопоставлений коробка → артикул")
            print(f"✅ Найдено {len(box_to_invoice_mapping)} сопоставлений коробка → инвойс")
            
            # Заполняем данные
            filled_df, stats = self._fill_data(weights_df, box_to_article_mapping, box_to_invoice_mapping)
            
            print(f"\n🔄 Результаты заполнения:")
            print(f"📊 Заполнено артикулов: {stats['filled_articles']}")
            print(f"📊 Заполнено инвойсов: {stats['filled_invoices']}")
            print(f"⚠️  Не найдено соответствий: {stats['not_found']}")
            
            # Сохранение результата
            output_path = self.file_manager.generate_output_path(input_path, "_filled")
            filled_df.to_excel(output_path, index=False, engine=self.settings.excel_engine)
            
            # Показываем примеры
            self._show_examples(filled_df)
            
            # Архивирование исходного файла
            self.file_manager.archive_file(input_path)
            
            result_message = f"Заполнено артикулов: {stats['filled_articles']}, инвойсов: {stats['filled_invoices']}"
            
            return ProcessingResult(
                success=True,
                input_file=input_path,
                output_file=output_path,
                message=result_message,
                details={
                    "filled_articles": stats['filled_articles'],
                    "filled_invoices": stats['filled_invoices'],
                    "not_found": stats['not_found'],
                    "reference_file": str(reference_file)
                }
            )
            
        except Exception as e:
            error_msg = f"Ошибка заполнения: {str(e)}"
            self.logger.error(f"Ошибка обработки {input_path}: {e}")
            return ProcessingResult(success=False, input_file=input_path, message=error_msg)
    
    def _find_reference_file(self, input_path: Path, kwargs: dict) -> Optional[Path]:
        """Поиск эталонного файла"""
        # Проверяем, передан ли эталонный файл в параметрах
        reference_file = kwargs.get('reference_file')
        if reference_file and Path(reference_file).exists():
            return Path(reference_file)
        
        # Ищем в папке input файлы, которые могут быть эталонными
        all_files = self.file_manager.get_input_files()
        
        # Исключаем текущий обрабатываемый файл
        reference_candidates = [f for f in all_files if f != input_path]
        
        # Если есть только один другой файл - берем его
        if len(reference_candidates) == 1:
            return reference_candidates[0]
        
        # Ищем файлы с характерными названиями
        priority_patterns = ['replay', 'spec', 'specification', 'invoice', 'inv', 'эталон']
        
        for pattern in priority_patterns:
            for file in reference_candidates:
                if pattern.lower() in file.name.lower():
                    return file
        
        # Если не найдено по паттернам, берем первый доступный файл
        if reference_candidates:
            return reference_candidates[0]
        
        return None
    
    def _identify_files(self, file1: Path, file2: Path) -> Tuple[Path, Path]:
        """Определяем, какой файл - файл весов, а какой - спецификация"""
        
        # Простой способ: файл с "box" или "weight" в названии = файл весов
        if any(keyword in file1.name.lower() for keyword in ['box', 'weight', 'вес']):
            return file1, file2  # file1 - weights, file2 - spec
        elif any(keyword in file2.name.lower() for keyword in ['box', 'weight', 'вес']):
            return file2, file1  # file2 - weights, file1 - spec
        
        # Если не определили по названию, читаем заголовки
        try:
            df1 = pd.read_excel(file1, header=1, nrows=0)
            df2 = pd.read_excel(file2, header=1, nrows=0)
            
            # Файл весов содержит 'Handling Unit'
            if 'Handling Unit' in df1.columns:
                return file1, file2
            elif 'Handling Unit' in df2.columns:
                return file2, file1
            
            # Файл спецификации содержит больше столбцов
            if len(df1.columns) > len(df2.columns):
                return df2, df1  # df2 - weights, df1 - spec
            else:
                return df1, df2  # df1 - weights, df2 - spec
                
        except Exception as e:
            self.logger.warning(f"Не удалось автоматически определить типы файлов: {e}")
            # По умолчанию считаем первый файл - файлом весов
            return file1, file2
    
    def _read_file_with_fixed_header(self, file_path: Path) -> pd.DataFrame:
        """Чтение файла с заголовками ВСЕГДА во второй строке"""
        print(f"🔍 Читаем файл: {file_path.name}")
        
        # ВСЕГДА читаем с header=1 (вторая строка)
        df = pd.read_excel(file_path, header=1, engine=self.settings.excel_engine)
        
        # Очищаем имена столбцов
        df.columns = [str(c).strip() for c in df.columns]
        
        # Удаляем полностью пустые строки
        df = df.dropna(how='all').reset_index(drop=True)
        
        return df
    
    def _create_mappings(self, spec_df: pd.DataFrame) -> Tuple[Dict[str, str], Dict[str, str]]:
        """Создание словарей сопоставления из файла спецификации"""
        print("\n🗂️  Создаем словари сопоставления...")
        
        # Ищем нужные столбцы в спецификации
        box_col = None
        article_col = None
        invoice_col = None
        
        print(f"   Доступные столбцы в спецификации: {list(spec_df.columns)}")
        
        for col in spec_df.columns:
            col_lower = str(col).lower().strip()
            
            # Ищем столбец с номерами коробок
            if any(pattern in col_lower for pattern in ['№ кор', 'номер кор', 'коробк', 'box', '№кор']):
                box_col = col
                print(f"   ✅ Найден столбец коробок: '{col}'")
            
            # Ищем столбец с артикулами
            elif any(pattern in col_lower for pattern in ['артикул', 'article']):
                article_col = col
                print(f"   ✅ Найден столбец артикулов: '{col}'")
            
            # Ищем столбец с инвойсами
            elif any(pattern in col_lower for pattern in ['№ инвойса', 'номер инвойса', 'invoice', '№инвойса']):
                invoice_col = col
                print(f"   ✅ Найден столбец инвойсов: '{col}'")
        
        if not box_col or not article_col:
            self.logger.error(f"Не найдены ключевые столбцы: коробки='{box_col}', артикулы='{article_col}'")
            return {}, {}
        
        # Создаем маппинги
        box_to_article = {}
        box_to_invoice = {}
        
        try:
            for _, row in spec_df.iterrows():
                box_number = str(row[box_col]).strip()
                article = str(row[article_col]).strip()
                
                if box_number and box_number not in ['nan', ''] and article and article not in ['nan', '']:
                    box_to_article[box_number] = article
                
                if invoice_col:
                    invoice = str(row[invoice_col]).strip()
                    if box_number and box_number not in ['nan', ''] and invoice and invoice not in ['nan', '']:
                        box_to_invoice[box_number] = invoice
            
            print(f"✅ Создано {len(box_to_article)} сопоставлений коробка → артикул")
            print(f"✅ Создано {len(box_to_invoice)} сопоставлений коробка → инвойс")
            
            # Показываем примеры
            print("📋 Примеры сопоставлений:")
            for i, (box, article) in enumerate(list(box_to_article.items())[:5]):
                invoice = box_to_invoice.get(box, "N/A")
                print(f"   {box} → Артикул: {article}, Инвойс: {invoice}")
            
            return box_to_article, box_to_invoice
            
        except Exception as e:
            self.logger.error(f"Ошибка создания маппингов: {e}")
            return {}, {}
    
    def _fill_data(self, weights_df: pd.DataFrame, 
                   box_to_article: Dict[str, str], 
                   box_to_invoice: Dict[str, str]) -> Tuple[pd.DataFrame, Dict[str, int]]:
        """Заполнение данных в файле весов"""
        print("\n🔄 Заполняем пустые данные...")
        
        df = weights_df.copy()
        
        # Ищем нужные столбцы в файле весов
        handling_unit_col = None
        invoice_col = None
        article_col = None
        
        print(f"   Доступные столбцы в файле весов: {list(df.columns)}")
        
        for col in df.columns:
            col_lower = str(col).lower().strip()
            if 'handling unit' in col_lower:
                handling_unit_col = col
                print(f"   ✅ Найден столбец номеров коробок: '{col}'")
            elif col_lower == 'инвойс' or 'invoice' in col_lower:
                invoice_col = col
                print(f"   ✅ Найден столбец инвойсов: '{col}'")
            elif col_lower == 'артикул' or 'article' in col_lower:
                article_col = col
                print(f"   ✅ Найден столбец артикулов: '{col}'")
        
        if not handling_unit_col:
            raise ValueError("Не найден столбец 'Handling Unit' в файле весов")
        
        # Статистика
        filled_articles = 0
        filled_invoices = 0
        not_found = 0
        
        # Заполняем данные
        for idx, row in df.iterrows():
            handling_unit = str(row[handling_unit_col]).strip()
            
            if handling_unit and handling_unit not in ['nan', '']:
                # Заполняем артикул
                if article_col and handling_unit in box_to_article:
                    current_article = str(row[article_col]).strip() if pd.notna(row[article_col]) else ''
                    if not current_article or current_article == 'nan':
                        df.at[idx, article_col] = box_to_article[handling_unit]
                        filled_articles += 1
                
                # Заполняем инвойс
                if invoice_col and handling_unit in box_to_invoice:
                    current_invoice = str(row[invoice_col]).strip() if pd.notna(row[invoice_col]) else ''
                    if not current_invoice or current_invoice == 'nan':
                        df.at[idx, invoice_col] = box_to_invoice[handling_unit]
                        filled_invoices += 1
                
                # Считаем не найденные
                if handling_unit not in box_to_article:
                    not_found += 1
        
        stats = {
            'filled_articles': filled_articles,
            'filled_invoices': filled_invoices,
            'not_found': not_found
        }
        
        return df, stats
    
    def _show_examples(self, filled_df: pd.DataFrame) -> None:
        """Показ примеров заполненных данных"""
        try:
            print("\n📋 Примеры заполненных данных:")
            
            # Ищем столбцы
            handling_unit_col = None
            invoice_col = None
            article_col = None
            
            for col in filled_df.columns:
                col_lower = str(col).lower().strip()
                if 'handling unit' in col_lower:
                    handling_unit_col = col
                elif col_lower == 'инвойс':
                    invoice_col = col
                elif col_lower == 'артикул':
                    article_col = col
            
            if handling_unit_col:
                # Показываем первые 10 заполненных строк
                cols_to_show = [handling_unit_col]
                if invoice_col:
                    cols_to_show.append(invoice_col)
                if article_col:
                    cols_to_show.append(article_col)
                
                sample_data = filled_df[cols_to_show].head(10)
                
                for _, row in sample_data.iterrows():
                    handling_unit = row[handling_unit_col]
                    invoice = row[invoice_col] if invoice_col else 'N/A'
                    article = row[article_col] if article_col else 'N/A'
                    print(f"   {handling_unit} → Инвойс: {invoice}, Артикул: {article}")
            
        except Exception as e:
            self.logger.warning(f"Ошибка показа примеров: {e}")
