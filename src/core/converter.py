"""Converter"""
from abc import ABC, abstractmethod

class Converter(ABC):
    """转换器抽象基类"""
    
    @abstractmethod
    def convert(self, input_path: str, output_path: str) -> bool:
        """执行文件转换
        
        Args:
            input_path: 输入文件路径
            output_path: 输出文件路径
            
        Returns:
            bool: 转换是否成功
        """
        pass

class ConverterService:
    """转换服务"""
    
    def __init__(self):
        self._converters = {}
        
    def register_converter(self, name: str, converter: Converter):
        """注册转换器
        
        Args:
            name: 转换器名称
            converter: 转换器实例
        """
        self._converters[name] = converter
        
    def get_converter(self, name: str) -> Converter:
        """获取转换器
        
        Args:
            name: 转换器名称
            
        Returns:
            Converter: 转换器实例
        """
        return self._converters.get(name)