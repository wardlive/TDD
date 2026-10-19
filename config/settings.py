"""
Настройки и конфигурация TDD-Suite
"""

from pathlib import Path
from dataclasses import dataclass
from typing import Dict, Any

@dataclass
class Settings:
    """Конфигурация TDD-Suite"""
    
    # Базовые пути
    base_dir: Path = Path(__file__).parent.parent
    data_dir: Path = None
    input_dir: Path = None
    output_dir: Path = None
    intermediate_dir: Path = None
    archive_dir: Path = None
    
    # Настройки обработки
    auto_archive: bool = True
    
    # Настройки Excel
    excel_engine: str = "openpyxl"
    
    # Настройки сжатия
    compression_mode: str = "with_percentages"  # "simple" или "with_percentages"
    
    def __post_init__(self):
        """Инициализация путей после создания объекта"""
        if self.data_dir is None:
            self.data_dir = self.base_dir / "data"
        
        if self.input_dir is None:
            self.input_dir = self.data_dir / "input"
        
        if self.output_dir is None:
            self.output_dir = self.data_dir / "output"
        
        if self.intermediate_dir is None:
            self.intermediate_dir = self.data_dir / "intermediate"
        
        if self.archive_dir is None:
            self.archive_dir = self.data_dir / "archive"
    
    def update_setting(self, key: str, value: Any) -> None:
        """Обновление настройки"""
        if hasattr(self, key):
            setattr(self, key, value)
        else:
            raise ValueError(f"Неизвестная настройка: {key}")
    
    def get_settings_dict(self) -> Dict[str, Any]:
        """Получение всех настроек в виде словаря"""
        return {
            "compression_mode": self.compression_mode,
            "auto_archive": self.auto_archive,
            "excel_engine": self.excel_engine,
        }
