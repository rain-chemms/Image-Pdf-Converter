import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from Converter.ImageToPdfConverter import ImageToPdfConverter
from Converter.PdfToImageConverter import PdfToImageConverter

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF/Image Converter")
        self.root.geometry("500x200")
        
        # 初始化菜单栏
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)
        
        # 文件菜单
        self.setting_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="设置", menu=self.setting_menu)
        self.setting_menu.add_command(label="退出", command=self.root.quit)
        self.setting_menu.add_command(label="作者信息", command=self.show_author_info)
        
        
        # 转换菜单
        self.convert_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="转换器选择", menu=self.convert_menu)
        self.convert_menu.add_command(label="PDF 转 Images", command=lambda: self.show_frame("pdf_to_img"))
        self.convert_menu.add_command(label="Images 转 PDF", command=lambda: self.show_frame("img_to_pdf"))
        
        # 初始化框架
        self.current_frame = None
        
        # 创建PDF转图片界面
        self.pdf_to_img_frame = ttk.Frame(self.root)
        self._create_pdf_to_img_widgets()
        
        # 创建图片转PDF界面
        self.img_to_pdf_frame = ttk.Frame(self.root)
        self._create_img_to_pdf_widgets()
        
        # 默认显示第一个界面
        self.show_frame("pdf_to_img")
    
    def _create_pdf_to_img_widgets(self):
        # PDF路径输入
        ttk.Label(self.pdf_to_img_frame, text="PDF文件路径:").grid(column=0, row=0, padx=10, pady=10)
        self.pdf_path_var = tk.StringVar()
        self.pdf_entry = ttk.Entry(self.pdf_to_img_frame, textvariable=self.pdf_path_var, width=40)
        self.pdf_entry.grid(column=1, row=0, padx=10, pady=10)
        ttk.Button(self.pdf_to_img_frame, text="浏览...", command=lambda: self.browse_file(self.pdf_path_var)).grid(column=2, row=0)
        
        # 输出路径输入
        ttk.Label(self.pdf_to_img_frame, text="输出目录:").grid(column=0, row=1, padx=10, pady=10)
        self.output_dir_var = tk.StringVar()
        self.output_entry = ttk.Entry(self.pdf_to_img_frame, textvariable=self.output_dir_var, width=40)
        self.output_entry.grid(column=1, row=1, padx=10, pady=10)
        ttk.Button(self.pdf_to_img_frame, text="浏览...", command=lambda: self.browse_directory(self.output_dir_var)).grid(column=2, row=1)
        
        # 转换按钮
        self.convert_button = ttk.Button(self.pdf_to_img_frame, text="开始转换", command=self.convert_pdf_to_images)
        self.convert_button.grid(column=1, row=2, pady=20)
    
    def _create_img_to_pdf_widgets(self):
        # 图片文件夹输入
        ttk.Label(self.img_to_pdf_frame, text="图片文件夹:").grid(column=0, row=0, padx=10, pady=10)
        self.img_dir_var = tk.StringVar()
        self.img_entry = ttk.Entry(self.img_to_pdf_frame, textvariable=self.img_dir_var, width=40)
        self.img_entry.grid(column=1, row=0, padx=10, pady=10)
        ttk.Button(self.img_to_pdf_frame, text="浏览...", command=lambda: self.browse_directory(self.img_dir_var)).grid(column=2, row=0)
        
        # 输出PDF路径
        ttk.Label(self.img_to_pdf_frame, text="输出PDF路径:").grid(column=0, row=1, padx=10, pady=10)
        self.pdf_output_var = tk.StringVar()
        self.pdf_output_entry = ttk.Entry(self.img_to_pdf_frame, textvariable=self.pdf_output_var, width=40)
        self.pdf_output_entry.grid(column=1, row=1, padx=10, pady=10)
        ttk.Button(self.img_to_pdf_frame, text="浏览...", command=lambda: self.browse_file(self.pdf_output_var)).grid(column=2, row=1)
        
        # 转换按钮
        self.convert_button = ttk.Button(self.img_to_pdf_frame, text="开始转换", command=self.convert_images_to_pdf)
        self.convert_button.grid(column=1, row=2, pady=20)
    
    def show_frame(self, frame_name):
        """切换显示的框架"""
        if self.current_frame:
            self.current_frame.pack_forget()
        if frame_name == "pdf_to_img":
            self.current_frame = self.pdf_to_img_frame
        elif frame_name == "img_to_pdf":
            self.current_frame = self.img_to_pdf_frame
        self.current_frame.pack(fill=tk.BOTH, expand=True)
    
    def browse_file(self, var):
        """文件选择对话框"""
        filename = filedialog.askopenfilename(
            title="选择文件",
            filetypes=(("PDF Files", "*.pdf"), ("All Files", "*.*"))
        )
        if filename:
            var.set(filename)
    
    def browse_directory(self, var):
        """目录选择对话框"""
        directory = filedialog.askdirectory(title="选择目录")
        if directory:
            var.set(directory)
    
    def convert_pdf_to_images(self):
        """PDF转图片处理"""
        pdf_path = self.pdf_path_var.get()
        output_dir = self.output_dir_var.get() or Path(pdf_path).parent
        
        try:
            converter = PdfToImageConverter(pdf_path, output_dir=output_dir)
            converter.convert()
            messagebox.showinfo("完成", f"成功生成Word文档：{converter.docx_path}")
        except Exception as e:
            messagebox.showerror("错误", str(e))
    
    def convert_images_to_pdf(self):
        """图片转PDF处理"""
        img_dir = self.img_dir_var.get()
        pdf_output = self.pdf_output_var.get() or Path(img_dir).parent / "combined_images.pdf"
        
        try:
            converter = ImageToPdfConverter(image_dir=img_dir, pdf_path=pdf_output)
            converter.Convert()
            messagebox.showinfo("完成", f"成功生成PDF文件：{pdf_output}")
        except Exception as e:
            messagebox.showerror("错误", str(e))
            
    def show_author_info(self):
        about_window = tk.Toplevel(self.root)
        about_window.title("关于本程序:")
        about_window.geometry("300x300+400+200")  # 设置初始位置
        
        # 添加软件信息标签
        info_label = ttk.Label(
            about_window,
            text=
            """
                作者: John Rain\n
                邮箱: com.rainjohn.ch@gmail.com\n
                版本: 1.0\n
                支持格式: .PDF/.PNG/.JPG/.BMP
            """,
            justify=tk.LEFT,
            padding=(10, 10)
        )
        info_label.pack(padx=20, pady=20)
        
        # 添加关闭按钮
        close_button = ttk.Button(about_window, text="关闭", command=about_window.destroy)
        close_button.pack(pady=10)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()