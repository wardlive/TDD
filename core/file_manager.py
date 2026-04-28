"""
Управление файлами и папками TDD-Suite
"""

import shutil
from datetime import datetime
from pathlib import Path
from typing import List, Optional
import logging

class FileManager:
    """Менеджер файлов и папок"""
    
    def __init__(self, settings):
        self.settings = settings
        self.logger = logging.getLogger(__name__)
    
    def ensure_directories(self) -> None:
        """Создание всех необходимых папок"""
        directories = [
            self.settings.data_dir,
            self.settings.input_dir,
            self.settings.output_dir,
            self.settings.intermediate_dir,
            self.settings.archive_dir
        ]
        
        for directory in directories:
            directory.mkdir(parents=True, exist_ok=True)
            
        self.logger.info("Рабочие папки созданы/проверены")
    
    def get_input_files(self, pattern: str = "*.xlsx") -> List[Path]:
        """Получение списка входных файлов"""
        files = list(self.settings.input_dir.glob(pattern))
        files.extend(list(self.settings.input_dir.glob("*.xls")))
        return sorted(files)
    
    def generate_output_path(self, input_path: Path, suffix: str = "_processed") -> Path:
        """Генерация пути для выходного файла"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        stem = input_path.stem
        ext = input_path.suffix
        
        filename = f"{stem}{suffix}_{timestamp}{ext}"
        return self.settings.output_dir / filename
    
    def save_intermediate(self, df, input_path: Path, suffix: str = "_intermediate") -> Path:
        """Сохранение промежуточного файла"""
        output_path = self.generate_intermediate_path(input_path, suffix)
        df.to_excel(output_path, index=False, engine=self.settings.excel_engine)
        self.logger.info(f"Промежуточный файл сохранен: {output_path.name}")
        return output_path
    
    def generate_intermediate_path(self, input_path: Path, suffix: str = "_intermediate") -> Path:
        """Генерация пути для промежуточного файла"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        stem = input_path.stem
        ext = input_path.suffix
        
        filename = f"{stem}{suffix}_{timestamp}{ext}"
        return self.settings.intermediate_dir / filename
    
    def archive_file(self, file_path: Path) -> Optional[Path]:
        """Архивирование файла"""
        if not self.settings.auto_archive:
            return None
        
        try:
            archive_path = self.settings.archive_dir / file_path.name
            
            # Если файл уже есть в архиве, добавляем timestamp
            if archive_path.exists():
                timestamp = datetime.now().strftime("_%Y%m%d_%H%M%S")
                stem = file_path.stem
                ext = file_path.suffix
                archive_path = self.settings.archive_dir / f"{stem}{timestamp}{ext}"
            
            shutil.copy2(file_path, archive_path)
            self.logger.info(f"Файл заархивирован: {archive_path.name}")
            return archive_path
            
        except Exception as e:
            self.logger.error(f"Ошибка архивирования {file_path}: {e}")
            return None
    
    def get_statistics(self) -> dict:
        """Статистика по папкам"""
        stats = {}
        
        directories = {
            "input": self.settings.input_dir,
            "output": self.settings.output_dir,
            "intermediate": self.settings.intermediate_dir,
            "archive": self.settings.archive_dir
        }
        
        for name, directory in directories.items():
            if directory.exists():
                files = list(directory.glob("*.xlsx")) + list(directory.glob("*.xls"))
                stats[name] = {
                    "count": len(files),
                    "files": [f.name for f in files[-5:]]  # Последние 5 файлов
                }
            else:
                stats[name] = {"count": 0, "files": []}
        
        return stats
