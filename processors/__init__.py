# Processors package
from .base import DocumentProcessor, ProcessResult
from .pdf_processor import PDFProcessor
from .pptx_processor import PPTXProcessor
from .word_processor import WordProcessor
from .excel_processor import ExcelProcessor

__all__ = [
    'DocumentProcessor',
    'ProcessResult', 
    'PDFProcessor',
    'PPTXProcessor',
    'WordProcessor',
    'ExcelProcessor'
]
