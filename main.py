#!/usr/bin/env python3
"""
🚀 TDD-Suite v3.0 - таможня дает добро!
Единая система для всех операций с документами
"""

import sys
import logging

from cli.menu import MainMenu
from core.file_manager import FileManager
from config.settings import Settings

def setup_logging(data_dir):
    """Настройка системы логирования"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(data_dir / 'tdd_suite.log', encoding='utf-8')
        ]
    )

def main():
    """Точка входа в TDD-Suite"""
    settings = Settings()
    file_manager = FileManager(settings)
    file_manager.ensure_directories()
    setup_logging(settings.data_dir)
    logger = logging.getLogger(__name__)
    
    try:
        # Инициализация системы
        logger.info("Инициализация TDD-Suite v3.0")
        # Запуск главного меню
        menu = MainMenu(file_manager, settings)
        menu.run()
        
    except KeyboardInterrupt:
        print("\n\n👋 Работа прервана пользователем")
        return 130
    except Exception as e:
        print(f"\n✘ Критическая ошибка: {e}")
        logger.exception("Критическая ошибка")
        return 1
    finally:
        logger.info("Завершение работы TDD-Suite")

if __name__ == "__main__":
    sys.exit(main())
