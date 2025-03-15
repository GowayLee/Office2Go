import office
from core.converter import Converter

class PPT2PDFConverter(Converter):
    """PPT转PDF转换器"""
    
    def convert(self, input_path: str, output_path: str) -> bool:
        try:
            office.ppt.ppt2pdf(input_path, output_path)
            return True
        except Exception as e:
            print(f"PPT转PDF失败: {e}")
            return False

class PPT2IMGConverter(Converter):
    """PPT转长图转换器"""
    
    def convert(self, input_path: str, output_path: str) -> bool:
        try:
            office.ppt.ppt2img(input_path, output_path)
            return True
        except Exception as e:
            print(f"PPT转长图失败: {e}")
            return False