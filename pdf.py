# -*- coding: utf-8 -*-
import fitz  # PyMuPDF
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox


# ==================== 字体处理核心模块 ====================
def resource_path(relative_path):
    """ 获取资源的绝对路径（兼容打包环境） """
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def load_korean_font(page):
    """ 加载韩文字体 """
    try:
        # 检查字体是否已存在
        if any(font[3] == "korean" for font in page.get_fonts()):
            return True
        font_path = resource_path("NanumGothic.ttf")
        page.insert_font(fontname="korean", fontfile=font_path)
        return True
    except Exception as e:
        messagebox.showerror("字体错误", str(e))
        return False


# ==================== PDF处理核心逻辑 ====================
def calculate_text_rect(page, text, fontname, fontsize, point):
    """ 计算文本的矩形区域 """
    # 使用文本长度和字体大小估算文本宽度
    # 注意：这是一个近似值，实际宽度取决于字符本身
    avg_char_width = fontsize * 0.6  # 平均字符宽度估算
    text_width = len(text) * avg_char_width
    text_height = fontsize * 1.2  # 文本高度估算

    # 返回矩形区域 (x0, y0, x1, y1)
    return fitz.Rect(
        point[0], point[1] - fontsize,  # 左上角 (x0, y0)
                  point[0] + text_width, point[1] + text_height / 2  # 右下角 (x1, y1)
    )


def add_text_to_pdf(input_path, output_path, text1, text2):
    """ 在固定位置添加文本 """
    try:
        doc = fitz.open(input_path)
        page = doc[0]  # 假设所有操作都在第一页

        # 字体设置
        fontname = "korean"
        fontsize = 13
        text_color = (1, 0, 0)  # 红色

        # 添加第一处文本（坐标452,120）
        if text1:
            if not load_korean_font(page):
                return False
            point1 = (452, 125)
            # 计算文本覆盖区域
            cover_rect1 = calculate_text_rect(page, text1, fontname, fontsize, point1)
            # 覆盖原有内容
            page.draw_rect(cover_rect1, color=(1, 1, 1), fill=(1, 1, 1), overlay=True)
            # 插入新文本
            page.insert_text(
                point=point1,
                text=text1,
                fontname=fontname,
                fontsize=fontsize,
                color=text_color,
                overlay=True
            )

        # 添加第二处文本（坐标340,419）
        if text2:
            if not load_korean_font(page):
                return False
            point2 = (340, 419)
            # 计算文本覆盖区域
            cover_rect2 = calculate_text_rect(page, text2, fontname, fontsize, point2)
            # 覆盖原有内容
            page.draw_rect(cover_rect2, color=(1, 1, 1), fill=(1, 1, 1), overlay=True)
            page.insert_text(
                point=point2,
                text=text2,
                fontname=fontname,
                fontsize=fontsize,
                color=text_color,
                overlay=True
            )

        doc.save(output_path)
        return True
    except Exception as e:
        messagebox.showerror("PDF处理错误", str(e))
        return False


# ==================== 图形用户界面 ====================
class PDFEditorApp:
    def __init__(self, master):
        self.master = master
        master.title("项目编号编辑")
        master.geometry("450x270")

        # 初始化变量
        self.input_pdf = ""
        self.text1 = tk.StringVar()  # 第一处文本
        self.text2 = tk.StringVar()  # 第二处文本

        # 创建界面组件
        self.create_widgets()
        # 状态提示
        self.lbl_status = tk.Label(self.master, text="By Chase Wang", fg="gray")
        self.lbl_status.grid(row=5, column=0, columnspan=2, pady=10)

        # 添加右下角署名
        self.signature = tk.Label(
            self.master,
            text="(*´∀`)~♥",
            fg="gray60",
            font=("Arial", 13)
        )
        self.signature.grid(row=20, column=3, padx=10, pady=5, sticky="se")

    def create_widgets(self):
        # 文件选择区域
        tk.Button(self.master, text="1. 选择PDF文件", command=self.select_pdf).grid(row=0, column=0, padx=10, pady=10)
        self.lbl_file = tk.Label(self.master, text="未选择文件", fg="gray")
        self.lbl_file.grid(row=0, column=1, sticky="w")

        # 第一处文本输入
        tk.Label(self.master, text="第一处内容：").grid(row=1, column=0, sticky="e")
        tk.Entry(self.master, textvariable=self.text1, width=30).grid(row=1, column=1, padx=5)

        # 第二处文本输入
        tk.Label(self.master, text="第二处内容：").grid(row=2, column=0, sticky="e")
        tk.Entry(self.master, textvariable=self.text2, width=30).grid(row=2, column=1, padx=5)

        # 生成按钮
        tk.Button(self.master, text="生成PDF", command=self.generate_pdf,
                  bg="#4CAF50", fg="white", width=15).grid(row=4, column=0, columnspan=2, pady=20)

    def select_pdf(self):
        """ 选择PDF文件 """
        file_path = filedialog.askopenfilename(filetypes=[("PDF文件", "*.pdf")])
        if file_path:
            self.input_pdf = file_path
            self.lbl_file.config(text=os.path.basename(file_path), fg="green")

    def generate_pdf(self):
        """ 生成新PDF """
        if not self.input_pdf:
            messagebox.showwarning("警告", "请先选择PDF文件！")
            return

        text1 = self.text1.get().strip()
        text2 = self.text2.get().strip()

        if not text1 and not text2:
            messagebox.showwarning("警告", "至少需要输入一处文本内容！")
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF文件", "*.pdf")]
        )
        if not output_path:
            return

        success = add_text_to_pdf(self.input_pdf, output_path, text1, text2)
        if success:
            messagebox.showinfo("成功", f"文件已生成：\n{output_path}")
            os.startfile(os.path.dirname(output_path))  # 打开所在文件夹


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFEditorApp(root)
    root.mainloop()
