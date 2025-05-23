import os
import zipfile
from tkinter import Tk, Button, Label, filedialog, messagebox, Entry, StringVar
from pathlib import Path

# 固定输出路径
FIXED_OUTPUT_FOLDER = r"文件夹路径"

# 提取图片并按顺序重命名函数
def extract_images_from_docx(docx_path, prefix):
    # 创建输出文件夹
    if not os.path.exists(FIXED_OUTPUT_FOLDER):
        os.makedirs(FIXED_OUTPUT_FOLDER)

    image_counter = 1

    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        for file_info in docx_zip.infolist():
            if file_info.filename.startswith('word/media/'):
                image_ext = os.path.splitext(file_info.filename)[1]
                new_image_filename = f"{prefix}{image_counter}{image_ext}"
                image_path = os.path.join(FIXED_OUTPUT_FOLDER, new_image_filename)

                with open(image_path, 'wb') as image_file:
                    image_file.write(docx_zip.read(file_info.filename))

                print(f"提取并重命名图片: {new_image_filename} 到 {FIXED_OUTPUT_FOLDER}")
                image_counter += 1

    messagebox.showinfo("完成", f"图片已保存到: {FIXED_OUTPUT_FOLDER}")

# 选择Word文件
def select_docx_file():
    file_path = filedialog.askopenfilename(
        title="选择Word文档",
        filetypes=[("Word 文件", "*.docx")]
    )
    if file_path:
        docx_path.set(file_path)

# 执行提取操作
def run_extraction():
    if not docx_path.get():
        messagebox.showwarning("警告", "请选择Word文档！")
        return

    prefix = filename_prefix.get()
    if not prefix:
        prefix = ""

    extract_images_from_docx(docx_path.get(), prefix)

# 创建UI窗口
root = Tk()
root.title("Word 图片提取工具")
root.geometry("400x330")

docx_path = StringVar()
filename_prefix = StringVar()

Label(root, text="选择Word文档:").pack(pady=5)
Button(root, text="选择文档", command=select_docx_file).pack(pady=5)
Label(root, textvariable=docx_path, wraplength=380).pack(pady=5)

Label(root, text="设置文件名前缀（可选）:").pack(pady=5)
Entry(root, textvariable=filename_prefix).pack(pady=5)

Label(root, text=f"图片将保存至固定路径:\n{FIXED_OUTPUT_FOLDER}", wraplength=380, fg="gray").pack(pady=10)

Button(root, text="提取图片", command=run_extraction).pack(pady=20)

root.mainloop()
