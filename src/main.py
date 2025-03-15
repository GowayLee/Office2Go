import os
import sys

# 将src目录及其子目录添加到Python路径
base_dir = os.path.dirname(os.path.dirname(__file__))
sys.path.extend([
    base_dir,
    os.path.join(base_dir, 'core'),
    os.path.join(base_dir, 'services'),
    os.path.join(base_dir, 'presentation'),
    os.path.join(base_dir, 'infrastructure')
])

from core import ConverterService
from infrastructure.office_impl import PPT2PDFConverter, PPT2IMGConverter
from presentation.gui import MainApplication

def initialize_converters():
    """初始化所有转换器"""
    converter_service = ConverterService()
    converter_service.register_converter("PPT转PDF", PPT2PDFConverter())
    converter_service.register_converter("PPT转长图", PPT2IMGConverter())
    return converter_service

def main():
    # 初始化转换器服务
    converter_service = initialize_converters()
    
    # 初始化应用程序
    app = MainApplication(converter_service)
    app.run()

if __name__ == "__main__":
    main()