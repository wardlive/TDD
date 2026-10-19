"""
Управление файлами и папками TDD-Suite
"""

import shutil
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from tempfile import NamedTemporaryFile
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
        return self.get_excel_files(self.settings.input_dir, pattern)

    @staticmethod
    def get_excel_files(directory: Path, pattern: str = "*.xlsx") -> List[Path]:
        return sorted(
            (path for path in directory.glob("*" if pattern == "*.xlsx" else pattern)
             if path.is_file() and not path.name.startswith("~$")
             and path.suffix.lower() == ".xlsx"),
            key=lambda path: path.name.casefold(),
        )

    @contextmanager
    def atomic_output(self, output_path: Path):
        """Не оставлять недописанный результат после ошибки или прерывания."""
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with NamedTemporaryFile(dir=output_path.parent, prefix=".tdd_",
                                suffix=output_path.suffix, delete=False) as handle:
            temporary_path = Path(handle.name)
        try:
            yield temporary_path
            temporary_path.replace(output_path)
        finally:
            temporary_path.unlink(missing_ok=True)
    
    def generate_output_path(self, input_path: Path, suffix: str = "_processed") -> Path:
        """Генерация пути для выходного файла"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        stem = input_path.stem
        ext = input_path.suffix
        
        filename = f"{stem}{suffix}_{timestamp}{ext}"
        return self.settings.output_dir / filename
    
    def save_intermediate(self, df, input_path: Path, suffix: str = "_intermediate") -> Path:
        """Сохранение промежуточного файла"""
        output_path = self.generate_intermediate_path(input_path, suffix)
        with self.atomic_output(output_path) as temporary_path:
            df.to_excel(temporary_path, index=False, engine=self.settings.excel_engine)
        self.logger.info(f"Промежуточный файл сохранен: {output_path.name}")
        return output_path
    
    def generate_intermediate_path(self, input_path: Path, suffix: str = "_intermediate") -> Path:
        """Генерация пути для промежуточного файла"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
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
                timestamp = datetime.now().strftime("_%Y%m%d_%H%M%S_%f")
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
                files = sorted(self.get_excel_files(directory), key=lambda path: path.stat().st_mtime)
                stats[name] = {
                    "count": len(files),
                    "files": [f.name for f in files[-5:]]  # Последние 5 файлов
                }
            else:
                stats[name] = {"count": 0, "files": []}
        
        return stats
