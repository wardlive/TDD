"""
Процессоры TDD-Suite v3.0
"""

from .compressor import SpecificationCompressor
from .cleaner import FileCleaner
from .merger import SheetMerger
from .article_filler import ArticleFiller

__all__ = [
    'SpecificationCompressor',
    'FileCleaner',
    'SheetMerger',
    'ArticleFiller'
]
