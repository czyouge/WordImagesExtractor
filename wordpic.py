import os
import zipfile
from tkinter import Tk, Button, Label, filedialog, messagebox
from tkinter import StringVar
from shutil import copyfile
from pathlib import Path

# 提取图片并按顺序重命名函数
def extract_images_from_docx(docx_path, output_folder):
    # 创建输出文件夹
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 图片计数器
    image_counter = 1

    # 将docx文件视为zip文件解压
    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        # 列出压缩包中的文件
        for file_info in docx_zip.infolist():
            # 提取图像文件
            if file_info.filename.startswith('word/media/'):
                image_ext = os.path.splitext(file_info.filename)[1]  # 获取图片扩展名
                new_image_filename = f"{image_counter}{image_ext}"   # 新文件名：1.jpg, 2.png等
                image_path = os.path.join(output_folder, new_image_filename)

                # 将图像文件保存到指定的输出文件夹
                with open(image_path, 'wb') as image_file:
                    image_file.write(docx_zip.read(file_info.filename))

                print(f"提取并重命名图片: {new_image_filename} 到 {output_folder}")
                image_counter += 1  # 增加计数器

    messagebox.showinfo("完成", f"图片已保存到: {output_folder}")

# 选择Word文件
def select_docx_file():
    file_path = filedialog.askopenfilename(
        title="选择Word文档",
        filetypes=[("Word 文件", "*.docx")]
    )
    if file_path:
        docx_path.set(file_path)

# 选择保存文件夹
def select_output_folder():
    folder_path = filedialog.askdirectory(title="选择保存图片的文件夹")
    if folder_path:
        output_folder.set(folder_path)

# 执行提取操作
def run_extraction():
    if not docx_path.get():
        messagebox.showwarning("警告", "请选择Word文档！")
        return
    if not output_folder.get():
        messagebox.showwarning("警告", "请选择保存图片的文件夹！")
        return

    extract_images_from_docx(docx_path.get(), output_folder.get())

# 创建UI窗口
root = Tk()
root.title("Word 图片提取工具")
root.geometry("400x500")

# 文档路径
docx_path = StringVar()
output_folder = StringVar()

# 标签和按钮
Label(root, text="选择Word文档:").pack(pady=5)
Button(root, text="选择文档", command=select_docx_file).pack(pady=5)
Label(root, textvariable=docx_path).pack(pady=5)

Label(root, text="选择保存文件夹:").pack(pady=5)
Button(root, text="选择文件夹", command=select_output_folder).pack(pady=5)
Label(root, textvariable=output_folder).pack(pady=5)

Button(root, text="提取图片", command=run_extraction).pack(pady=20)

# 运行主循环
root.mainloop()
