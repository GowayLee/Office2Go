"""GUI"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from services import TaskService

class MainApplication:
    """主应用程序"""
    def __init__(self, converter_service):
        self.root = tk.Tk()
        self.converter_service = converter_service
        self.task_service = TaskService()
        self._setup_ui()

    def _setup_ui(self):
        self.root.title("批量Office文件操作脚本")
        self.root.geometry("700x400")
        self._create_widgets()

    def _create_widgets(self):
        # 创建主界面组件
        self._create_input_frame()
        self._create_output_frame()
        self._create_options_frame()
        self._create_status_bar()
        self._create_start_button()

    def _create_input_frame(self):
        frame = ttk.Frame(self.root)
        frame.grid(row=0, column=0, padx=10, pady=10, sticky='ew')

        ttk.Label(frame, text="选择操作文件夹:").grid(row=0, column=0, padx=5, pady=5)
        self.input_folder_entry = ttk.Entry(frame, width=40)
        self.input_folder_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(frame, text="浏览", command=self._select_input_folder).grid(row=0, column=2, padx=5, pady=5)

    def _select_input_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.input_folder_entry.delete(0, tk.END)
            self.input_folder_entry.insert(0, folder_path)
            self._validate_inputs()

    def _create_output_frame(self):
        frame = ttk.Frame(self.root)
        frame.grid(row=1, column=0, padx=10, pady=10, sticky='ew')

        ttk.Label(frame, text="选择输出文件夹:").grid(row=0, column=0, padx=5, pady=5)
        self.output_folder_entry = ttk.Entry(frame, width=40)
        self.output_folder_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(frame, text="浏览", command=self._select_output_folder).grid(row=0, column=2, padx=5, pady=5)

    def _select_output_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.output_folder_entry.delete(0, tk.END)
            self.output_folder_entry.insert(0, folder_path)
            self._validate_inputs()

    def _create_options_frame(self):
        frame = ttk.Frame(self.root)
        frame.grid(row=2, column=0, padx=10, pady=10, sticky='ew')

        ttk.Label(frame, text="转换选项:").grid(row=0, column=0, padx=5, pady=5)
        self.conversion_option = tk.StringVar(self.root)
        self.conversion_option.set("选择操作")
        
        # 获取转换器选项
        options = ["选择操作"] + self.converter_service.get_available_converters()
        
        # 创建选项菜单
        self.option_menu = ttk.OptionMenu(frame, self.conversion_option, *options, command=lambda _: self._validate_inputs())
        self.option_menu.grid(row=0, column=1, padx=5, pady=5)

    def _create_status_bar(self):
        self.status_label = ttk.Label(self.root, text="状态: 等待操作", relief=tk.SUNKEN, anchor='w')
        self.status_label.grid(row=6, column=0, columnspan=3, sticky='ew')

    def _create_start_button(self):
        self.start_button = ttk.Button(self.root, text="开始转换", command=self._start_conversion, state=tk.DISABLED)
        self.start_button.grid(row=5, column=1, padx=5, pady=20)

    def _validate_inputs(self):
        if self.input_folder_entry.get() and self.output_folder_entry.get() and self.conversion_option.get() != "选择操作":
            self.start_button.config(state=tk.NORMAL)
        else:
            self.start_button.config(state=tk.DISABLED)

    def _start_conversion(self):
        input_folder = self.input_folder_entry.get()
        output_folder = self.output_folder_entry.get()
        conversion_type = self.conversion_option.get()

        self.start_button.config(state=tk.DISABLED)
        self.status_label.config(text="状态: 转换中...")

        self.task_service.submit(self._run_conversion, input_folder, output_folder, conversion_type)

    def _run_conversion(self, input_folder, output_folder, conversion_type):
        converter = self.converter_service.get_converter(conversion_type)
        if not converter:
            messagebox.showerror("错误", "无效的转换类型")
            return

        # 执行转换逻辑
        try:
            for file_name in os.listdir(input_folder):
                file_path = os.path.join(input_folder, file_name)
                if file_name.endswith('.ppt') or file_name.endswith('.pptx'):
                    converter.convert(file_path, output_folder)
        except Exception as e:
            messagebox.showerror("错误", f"转换过程中发生错误: {e}")
        finally:
            self.root.after(0, lambda: self.status_label.config(text="状态: 转换完成"))
            self.root.after(0, lambda: self.start_button.config(state=tk.NORMAL))

    def run(self):
        self.root.mainloop()