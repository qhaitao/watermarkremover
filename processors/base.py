#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
处理器基类 - 定义统一接口
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, List, Set

# ============================================================================
#                           处理结果
# ============================================================================

@dataclass
class ProcessResult:
    """处理结果"""
    success: bool
    output_path: Optional[str] = None
    message: str = ""
    removed_count: int = 0
    page_count: int = 0
    
    def __repr__(self):
        return f"ProcessResult(success={self.success}, removed={self.removed_count}, msg='{self.message}')"

# ============================================================================
#                           处理器基类
# ============================================================================

class DocumentProcessor(ABC):
    """文档处理器抽象基类"""
    
    # 子类需定义支持的扩展名
    SUPPORTED_EXTENSIONS: Set[str] = set()
    
    def __init__(self, preview: bool = False, keywords: List[str] = None):
        """
        初始化处理器
        
        Args:
            preview: 预览模式（只扫描不修改）
            keywords: 过滤关键词列表
        """
        self.preview = preview
        self.keywords = keywords or []
        self.stats = {'pages': 0, 'removed': 0, 'scanned': 0}
    
    @classmethod
    def supports(cls, file_path: str) -> bool:
        """检查是否支持该文件类型"""
        ext = Path(file_path).suffix.lower()
        return ext in cls.SUPPORTED_EXTENSIONS
    
    @staticmethod
    def get_default_output(input_path: str, suffix: str = "_处理后") -> str:
        """生成默认输出路径"""
        p = Path(input_path)
        return str(p.parent / f"{p.stem}{suffix}{p.suffix}")
    
    @abstractmethod
    def process(self, input_path: str, output_path: str = None,
                progress_callback=None) -> ProcessResult:
        """
        处理文件
        
        Args:
            input_path: 输入文件路径
            output_path: 输出文件路径（可选）
            progress_callback: 进度回调函数 callback(current, total, message)
            
        Returns:
            ProcessResult: 处理结果
        """
        pass
    
    @abstractmethod
    def get_description(self) -> str:
        """获取处理器描述"""
        pass
