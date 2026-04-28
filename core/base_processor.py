"""
Базовые классы для всех процессоров TDD-Suite
"""

from abc import ABC, abstractmethod
from pathlib import Path
from typing import List, Dict, Any, Optional
import logging

class ProcessingResult:
    """Результат обработки файла"""
    
    def __init__(self, success: bool, input_file: Path, output_file: Path = None, 
                 message: str = "", details: Dict[str, Any] = None):
        self.success = success
        self.input_file = input_file
        self.output_file = output_file
        self.message = message
        self.details = details or {}

class BaseProcessor(ABC):
    """Базовый класс для всех процессоров"""
    
    def __init__(self, file_manager, settings):
        self.file_manager = file_manager
        self.settings = settings
        self.logger = logging.getLogger(self.__class__.__name__)
    
    @abstractmethod
    def process_file(self, input_path: Path, **kwargs) -> ProcessingResult:
        """Обработать один файл"""
        pass
    
    @abstractmethod
    def get_name(self) -> str:
        """Название процессора"""
        pass
    
    @abstractmethod
    def get_description(self) -> str:
        """Описание процессора"""
        pass
    
    def validate_input_file(self, file_path: Path) -> bool:
        """Валидация входного файла"""
        if not file_path.exists():
            self.logger.error(f"Файл не существует: {file_path}")
            return False
        
        if not file_path.suffix.lower() in ['.xlsx', '.xls']:
            self.logger.error(f"Неподдерживаемый формат файла: {file_path}")
            return False
        
        return True
    
    def process_batch(self, input_files: List[Path], **kwargs) -> Dict[str, Any]:
        """Пакетная обработка файлов"""
        results = {
            "total": len(input_files),
            "success": 0,
            "failed": 0,
            "results": [],
            "errors": []
        }
        
        self.logger.info(f"Начало пакетной обработки {len(input_files)} файлов")
        
        for i, file_path in enumerate(input_files, 1):
            try:
                print(f"\n📄 Обрабатываем ({i}/{len(input_files)}): {file_path.name}")
                
                result = self.process_file(file_path, **kwargs)
                results["results"].append(result)
                
                if result.success:
                    results["success"] += 1
                    print(f"✔︎ Успешно: {result.message}")
                else:
                    results["failed"] += 1
                    print(f"✘ Ошибка: {result.message}")
                    results["errors"].append(f"{file_path.name}: {result.message}")
                    
            except Exception as e:
                results["failed"] += 1
                error_msg = f"Неожиданная ошибка: {str(e)}"
                results["errors"].append(f"{file_path.name}: {error_msg}")
                print(f"✘ {error_msg}")
                self.logger.error(f"Ошибка обработки {file_path}: {e}")
        
        # Итоговый отчет
        print(f"\n🎉 ПАКЕТНАЯ ОБРАБОТКА ЗАВЕРШЕНА:")
        print(f"   ✔︎ Успешно: {results['success']}/{results['total']}")
        print(f"   ✘ Ошибок: {results['failed']}/{results['total']}")
        
        if results['errors']:
            print(f"\n⚠️ Ошибки:")
            for error in results['errors'][:5]:  # Показываем первые 5 ошибок
                print(f"   • {error}")
            if len(results['errors']) > 5:
                print(f"   • ... и еще {len(results['errors']) - 5} ошибок")
        
        self.logger.info(f"Пакетная обработка завершена: {results['success']}/{results['total']} успешно")
        return results
