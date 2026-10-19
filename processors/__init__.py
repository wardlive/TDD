"""
Процессоры TDD-Suite v3.0
"""

from .compressor import SpecificationCompressor
from .cleaner import FileCleaner
from .merger import SheetMerger
from .diesel import DieselProcessor

__all__ = [
    'SpecificationCompressor',
    'FileCleaner',
    'SheetMerger',
    'DieselProcessor'
]
