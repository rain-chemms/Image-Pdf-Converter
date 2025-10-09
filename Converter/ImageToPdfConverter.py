import img2pdf
import os
from pathlib import Path

class ImageToPdfConverter:
    """
    图片转PDF工具类
    
    属性：
        image_dir (str): 图片文件夹路径（默认当前目录）
        pdf_path (str): 生成的PDF文件保存路径（默认同图片目录下的combined_images.pdf）
    
    方法：
        convert(): 执行图片转PDF操作
    """
    
    def __init__(self, image_dir=None, pdf_path=None):
        # 设置默认图片目录为当前工作目录
        self.image_dir = Path.cwd() if image_dir is None else Path(image_dir)
        # 设置默认PDF路径为图片目录下的combined_images.pdf
        self.pdf_path = self.image_dir / "combined_images.pdf" if pdf_path is None else Path(pdf_path)
        
        # 自动创建PDF保存目录（如果不存在）
        self.pdf_path.parent.mkdir(parents=True, exist_ok=True)

    """
        执行图片转PDF操作    
        返回：
            bool: 操作成功返回True，失败返回False
    """    
    def Convert(self):
        try:
            # 获取图片文件列表
            image_files = []
            for ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff']:
                image_files.extend(self.image_dir.glob(f"**/*{ext}"))
            
            if not image_files:
                print("错误：未找到任何支持的图片文件")
                return False
            
            # 按文件名排序
            image_files.sort(key=lambda x: x.stem)
            
            # 转换为PDF
            with open(self.pdf_path, "wb") as f:
                f.write(img2pdf.convert([str(f) for f in image_files]))
            
            print(f"成功生成PDF文件：{self.pdf_path}")
            print(f"共包含 {len(image_files)} 张图片")
            return True
        
        except Exception as e:
            print(f"转换过程中发生错误：{str(e)}")
            return False
    
    """设置图片文件夹路径"""
    def SetImageDir(self, image_dir):
        self.image_dir = Path(image_dir)
        
    """设置PDF保存路径"""      
    def SetPdfPath(self, pdf_path):
        self.pdf_path = Path(pdf_path)


"""
# 使用示例
#if __name__ == "__main__":
    # 方式1：使用默认参数
#    converter = ImageToPdfConverter()
#    converter.convert()
    
    # 方式2：自定义图片目录和PDF路径
#    custom_converter = ImageToPdfConverter(
#        image_dir="D:\\PROJECT\\Python\\Convert\\load",
#        pdf_path="D:\\OUTPUT\\my_photos.pdf"
#    )
#    custom_converter.convert()
"""