"""
Интерактивное меню TDD-Suite
"""

import sys
from pathlib import Path
from typing import List

from processors.compressor import SpecificationCompressor
from processors.cleaner import FileCleaner
from processors.merger import SheetMerger
from processors.article_filler import ArticleFiller

class MainMenu:
    """Главное интерактивное меню TDD-Suite"""
    
    def __init__(self, file_manager, settings):
        self.file_manager = file_manager
        self.settings = settings
        self.processors = self._initialize_processors()
    
    def _initialize_processors(self):
        """Инициализация процессоров"""
        return {
            "compressor": SpecificationCompressor(self.file_manager, self.settings),
            "cleaner": FileCleaner(self.file_manager, self.settings),
            "merger": SheetMerger(self.file_manager, self.settings),
            "article_filler": ArticleFiller(self.file_manager, self.settings)
        }
    
    def run(self):
        """Запуск интерактивного меню"""
        while True:
            self._show_main_menu()
            choice = input("\nВаш выбор: ").strip()
            
            try:
                if choice == "1":
                    self._compression_menu()
                elif choice == "2":
                    self._run_cleaner()
                elif choice == "3":
                    self._run_merger()
                elif choice == "4":
                    self._run_article_filler()
                elif choice == "5":
                    self._batch_processing_menu()
                elif choice == "6":
                    self._settings_menu()
                elif choice == "7":
                    self._show_statistics()
                elif choice == "8":
                    print("\n👋 До свидания!")
                    break
                else:
                    print("✘ Неверный выбор. Попробуйте еще раз.")
                    
            except KeyboardInterrupt:
                print("\n\n👋 Возврат в главное меню")
                continue
            except Exception as e:
                print(f"\n✘ Ошибка: {e}")
                input("Нажмите Enter для продолжения...")
    
    def _show_main_menu(self):
        """Отображение главного меню"""
        print("\n" + "=" * 70)
        print("🚀 TDD-Suite v3.0 - Комплексная обработка Excel файлов")
        print("=" * 70)
        print("1. 📊 Сжатие спецификаций")
        print("   • Простое сжатие")
        print("   • Сжатие с расчетом долей")
        print("2. 🧹 Очистка файлов (апострофы, D→.)")
        print("3. 🔗 Объединение листов PL INV")
        print("4. 📝 Заполнение артикулов")
        print("5. 📦 Пакетная обработка")
        print("6. ⚙️  Настройки")
        print("7. 📊 Статистика")
        print("8. ✘ Выход")
    
    def _compression_menu(self):
        """Подменю сжатия спецификаций"""
        while True:
            print("\n" + "="*50)
            print("📊 СЖАТИЕ СПЕЦИФИКАЦИЙ")
            print("="*50)
            print("1. 🔄 Простое сжатие")
            print("2. 📦 Сжатие с расчетом долей в коробках")
            print("3. 📄 Обработать выбранный файл")
            print("4. 📚 Обработать все файлы")
            print("5. 🔙 Назад")
            
            choice = input("\nВыберите режим: ").strip()
            
            if choice == "1":
                self._run_compression(calculate_percentages=False)
            elif choice == "2":
                self._run_compression(calculate_percentages=True)
            elif choice == "3":
                self._run_single_file_compression()
            elif choice == "4":
                self._run_all_files_compression()
            elif choice == "5":
                break
            else:
                print("✘ Неверный выбор")
    
    def _run_compression(self, calculate_percentages: bool = False):
        """Запуск сжатия с выбором режима"""
        files = self.file_manager.get_input_files()
        if not files:
            print("\n📭 Нет файлов для обработки в папке input/")
            input("Нажмите Enter для продолжения...")
            return
        
        mode_name = "с расчетом долей" if calculate_percentages else "простое"
        print(f"\n🔄 Запуск сжатия ({mode_name}) для {len(files)} файлов...")
        
        processor = self.processors["compressor"]
        processor.process_batch(files, calculate_percentages=calculate_percentages)
        
        input("\nНажмите Enter для продолжения...")
    
    def _run_single_file_compression(self):
        """Сжатие одного выбранного файла"""
        files = self.file_manager.get_input_files()
        if not files:
            print("\n📭 Нет файлов для обработки")
            return
        
        print("\n📋 Доступные файлы:")
        for i, file in enumerate(files, 1):
            print(f"{i:2d}. {file.name}")
        
        try:
            choice = int(input(f"\nВыберите файл (1-{len(files)}): "))
            if 1 <= choice <= len(files):
                selected_file = files[choice - 1]
                
                # Выбор режима сжатия
                print("\nВыберите режим сжатия:")
                print("1. Простое сжатие")
                print("2. С расчетом долей в коробках")
                
                mode_choice = input("Режим (1-2): ").strip()
                calculate_percentages = (mode_choice == "2")
                
                processor = self.processors["compressor"]
                result = processor.process_file(selected_file, calculate_percentages=calculate_percentages)
                
                if result.success:
                    print(f"\n✔︎ Файл успешно обработан: {result.message}")
                else:
                    print(f"\n✘ Ошибка обработки: {result.message}")
            else:
                print("✘ Неверный номер файла")
        except ValueError:
            print("✘ Введите число")
        except Exception as e:
            print(f"✘ Ошибка: {e}")
        
        input("\nНажмите Enter для продолжения...")
    
    def _run_all_files_compression(self):
        """Сжатие всех файлов"""
        files = self.file_manager.get_input_files()
        if not files:
            print("\n📭 Нет файлов для обработки")
            return
        
        print("\nВыберите режим сжатия для всех файлов:")
        print("1. Простое сжатие")
        print("2. С расчетом долей в коробках")
        
        mode_choice = input("Режим (1-2): ").strip()
        calculate_percentages = (mode_choice == "2")
        
        mode_name = "с расчетом долей" if calculate_percentages else "простое"
        print(f"\n🔄 Начинаем {mode_name} сжатие {len(files)} файлов...")
        
        processor = self.processors["compressor"]
        processor.process_batch(files, calculate_percentages=calculate_percentages)
        
        input("\nНажмите Enter для продолжения...")
    
    def _run_cleaner(self):
        """Запуск очистки файлов"""
        files = self.file_manager.get_input_files()
        if not files:
            print("\n📭 Нет файлов для обработки в папке input/")
            input("Нажмите Enter для продолжения...")
            return
        
        print(f"\n🧹 Запуск очистки файлов ({len(files)} файлов)...")
        
        processor = self.processors["cleaner"]
        processor.process_batch(files)
        
        input("\nНажмите Enter для продолжения...")
    
    def _show_statistics(self):
        """Показать статистику по папкам"""
        print("\n" + "="*50)
        print("📊 СТАТИСТИКА TDD-Suite")
        print("="*50)
        
        stats = self.file_manager.get_statistics()
        
        for folder_name, data in stats.items():
            folder_emoji = {"input": "📥", "output": "📤", "intermediate": "🔄", "archive": "📦"}.get(folder_name, "📁")
            print(f"\n{folder_emoji} {folder_name.upper()}: {data['count']} файлов")
            
            if data['files']:
                print("   Файлы:")
                for file in data['files']:
                    print(f"   • {file}")
                if data['count'] > len(data['files']):
                    print(f"   • ... и еще {data['count'] - len(data['files'])} файлов")
        
        input("\nНажмите Enter для продолжения...")
    
    def _settings_menu(self):
        """Меню настроек"""
        while True:
            print("\n" + "="*50)
            print("⚙️ НАСТРОЙКИ TDD-Suite")
            print("="*50)
            
            settings = self.settings.get_settings_dict()
            
            print("Текущие настройки:")
            print(f"1. Режим сжатия: {settings['compression_mode']}")
            print(f"2. Автоархивирование: {settings['auto_archive']}")
            print(f"3. Создание бэкапов: {settings['create_backups']}")
            print(f"4. Движок Excel: {settings['excel_engine']}")
            print("5. 🔙 Назад")
            
            choice = input("\nИзменить настройку: ").strip()
            
            if choice == "1":
                self._change_compression_mode()
            elif choice == "2":
                self.settings.auto_archive = not self.settings.auto_archive
                print(f"✓ Автоархивирование: {self.settings.auto_archive}")
            elif choice == "3":
                self.settings.create_backups = not self.settings.create_backups
                print(f"✓ Создание бэкапов: {self.settings.create_backups}")
            elif choice == "4":
                self._change_excel_engine()
            elif choice == "5":
                break
            else:
                print("✘ Неверный выбор")
            
            if choice in ["1", "2", "3", "4"]:
                input("Нажмите Enter для продолжения...")
    
    def _change_compression_mode(self):
        """Изменение режима сжатия по умолчанию"""
        print("\nВыберите режим сжатия по умолчанию:")
        print("1. simple - простое сжатие")
        print("2. with_percentages - с расчетом долей")
        
        choice = input("Режим: ").strip()
        
        if choice == "1":
            self.settings.compression_mode = "simple"
            print("✓ Установлен простой режим сжатия")
        elif choice == "2":
            self.settings.compression_mode = "with_percentages"
            print("✓ Установлен режим с расчетом долей")
        else:
            print("✘ Неверный выбор")

    def _run_merger(self):
        """Запуск объединения листов"""
        files = self.file_manager.get_input_files()
        if not files:
            print("\n📭 Нет файлов для обработки в папке input/")
            input("Нажмите Enter для продолжения...")
            return
        
        print(f"\n🔗 Запуск объединения листов ({len(files)} файлов)...")
        
        processor = self.processors["merger"]
        processor.process_batch(files)
        
        input("\nНажмите Enter для продолжения...")

    def _run_article_filler(self):
        """Запуск заполнения артикулов"""
        files = self.file_manager.get_input_files()
        if not files:
            print("\n📭 Нет файлов для обработки в папке input/")
            input("Нажмите Enter для продолжения...")
            return
        
        print(f"\n📝 Запуск заполнения артикулов ({len(files)} файлов)...")
        
        processor = self.processors["article_filler"]
        processor.process_batch(files)
        
        input("\nНажмите Enter для продолжения...")

    def _batch_processing_menu(self):
        """Меню пакетной обработки"""
        print("\n" + "="*50)
        print("📦 ПАКЕТНАЯ ОБРАБОТКА")
        print("="*50)
        print("1. Сжатие всех файлов")
        print("2. Очистка всех файлов")
        print("3. Объединение листов во всех файлах")
        print("4. Заполнение артикулов во всех файлах")
        print("5. 🔙 Назад")
        
        choice = input("\nВыберите операцию: ").strip()
        
        if choice == "1":
            self._run_all_files_compression()
        elif choice == "2":
            self._run_cleaner()
        elif choice == "3":
            self._run_merger()
        elif choice == "4":
            self._run_article_filler()
        elif choice == "5":
            return
        else:
            print("✘ Неверный выбор")

    def _change_excel_engine(self):
        """Изменение движка Excel"""
        print("\nВыберите движок Excel:")
        print("1. openpyxl (рекомендуется)")
        print("2. xlsxwriter")
        
        choice = input("Движок: ").strip()
        
        if choice == "1":
            self.settings.excel_engine = "openpyxl"
            print("✓ Установлен движок openpyxl")
        elif choice == "2":
            self.settings.excel_engine = "xlsxwriter"  
            print("✓ Установлен движок xlsxwriter")
        else:
            print("✘ Неверный выбор")
