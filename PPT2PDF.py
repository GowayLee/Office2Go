import office
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# 创建主窗口
root = tk.Tk()
root.title("批量Office文件操作脚本")
root.geometry("700x400")
# 创建状态标签
status_label = tk.Label(root, text="状态: 等待操作", relief=tk.SUNKEN, anchor='w')
status_label.grid(row=6, column=0, columnspan=3, sticky='ew')

# 定义选择文件夹的函数
def select_folder(entry_widget):
    folder_path = filedialog.askdirectory()
    if folder_path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, folder_path)
        validate_inputs()

# 输入验证
def validate_inputs():
    if input_folder_entry.get() and output_folder_entry.get() and conversion_option.get() != "选择操作":
        start_button.config(state=tk.NORMAL)
    else:
        start_button.config(state=tk.DISABLED)

# 定义转换处理函数
conversion_handlers = {
    "PPT转PDF": (office.ppt.ppt2pdf, ['.ppt', '.pptx']),
    "PPT转长图": (office.ppt.ppt2img, ['.ppt', '.pptx']),
}

def convert():
    input_folder_path = input_folder_entry.get()
    output_folder_path = output_folder_entry.get()
    case = conversion_option.get()
    
    if not input_folder_path or not output_folder_path:
        messagebox.showwarning("Warning", "请选择输入和输出文件夹")
        return

    if case == "选择操作":
        messagebox.showwarning("Warning", "请选择操作选项")
        return

    # 禁用开始按钮
    start_button.config(state=tk.DISABLED)
    status_label.config(text="状态: 转换中...")

    # 启动线程
    thread = threading.Thread(target=run_conversion, args=(input_folder_path, output_folder_path, case))
    thread.start()

def run_conversion(input_folder_path, output_folder_path, case):
    errors = []
    handler, extensions = conversion_handlers[case]
    
    for file_name in os.listdir(input_folder_path):
        if any(file_name.endswith(ext) for ext in extensions):
            file_path = os.path.join(input_folder_path, file_name)
            try:
                print(f"正在转换: {file_name}")
                handler(file_path, output_folder_path)
            except Exception as e:
                errors.append(f"转换 {file_name} 时发生错误: {e}")

    if errors:
        messagebox.showerror("Error", "\n".join(errors))
    
    # 更新状态
    root.after(0, lambda: status_label.config(text="状态: 转换完成"))
    root.after(0, lambda: start_button.config(state=tk.NORMAL))

# ======================================= 窗口布局 =======================================================
frame_input = ttk.Frame(root)
frame_input.grid(row=0, column=0, padx=10, pady=10, sticky='ew')

tk.Label(frame_input, text="选择操作文件夹:").grid(row=0, column=0, padx=5, pady=5)
input_folder_entry = ttk.Entry(frame_input, width=40)
input_folder_entry.grid(row=0, column=1, padx=5, pady=5)
ttk.Button(frame_input, text="浏览", command=lambda: select_folder(input_folder_entry)).grid(row=0, column=2, padx=5, pady=5)

frame_output = ttk.Frame(root)
frame_output.grid(row=1, column=0, padx=10, pady=10, sticky='ew')

tk.Label(frame_output, text="选择输出文件夹:").grid(row=0, column=0, padx=5, pady=5)
output_folder_entry = ttk.Entry(frame_output, width=40)
output_folder_entry.grid(row=0, column=1, padx=5, pady=5)
ttk.Button(frame_output, text="浏览", command=lambda: select_folder(output_folder_entry)).grid(row=0, column=2, padx=5, pady=5)

frame_options = ttk.Frame(root)
frame_options.grid(row=2, column=0, padx=10, pady=10, sticky='ew')

tk.Label(frame_options, text="转换选项:").grid(row=0, column=0, padx=5, pady=5)
options = ["选择操作", "PPT转PDF", "PPT转长图", "PDF转PPT"]
conversion_option = tk.StringVar(root)
conversion_option.set(options[0])
option_menu = ttk.OptionMenu(frame_options, conversion_option, *options, command=lambda _: validate_inputs())
option_menu.grid(row=0, column=1, padx=5, pady=5)

start_button = ttk.Button(root, text="开始转换", command=convert, state=tk.DISABLED)
start_button.grid(row=5, column=1, padx=5, pady=20)

# 绑定输入框变化事件
input_folder_entry.bind("<KeyRelease>", lambda _: validate_inputs())
output_folder_entry.bind("<KeyRelease>", lambda _: validate_inputs())

# ==========================================================================================================

root.mainloop()