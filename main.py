#!/usr/bin/env python3
"""
🚀 TDD-Suite v3.0 - таможня дает добро!
Единая система для всех операций с документами
"""

import sys
import logging
from pathlib import Path

# Добавляем текущую папку в путь для импорта модулей
sys.path.append(str(Path(__file__).parent))

from cli.menu import MainMenu
from core.file_manager import FileManager
from config.settings import Settings

def setup_logging():
    """Настройка системы логирования"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('data/tdd_suite.log', encoding='utf-8'),
            logging.StreamHandler()
        ]
    )

def main():
    """Точка входа в TDD-Suite"""
    # Настройка логирования
    setup_logging()
    logger = logging.getLogger(__name__)
    
    try:
        # Приветственное сообщение
        print("\n" + "=" * 70)
        print("TDD-Suite v3.0 - таможня дает добро!")
        print("✨ Единая система для всех операций с документами")
        print("=" * 70)
        
        # Инициализация системы
        logger.info("Инициализация TDD-Suite v3.0")
        settings = Settings()
        file_manager = FileManager(settings)
        
        # Создание рабочих папок
        file_manager.ensure_directories()
        
        # Запуск главного меню
        menu = MainMenu(file_manager, settings)
        menu.run()
        
    except KeyboardInterrupt:
        print("\n\n👋 Работа прервана пользователем")
    except Exception as e:
        print(f"\n✘ Критическая ошибка: {e}")
        logger.error(f"Критическая ошибка: {e}")
    finally:
        logger.info("Завершение работы TDD-Suite")

if __name__ == "__main__":
    main()
