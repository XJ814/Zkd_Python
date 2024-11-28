import os
import pdfplumber
import sys
import io
from PyPDF2 import PdfReader
import pandas as pd
import re
import traceback

# 设置标准输出为 UTF-8 编码
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
# 获取文件夹内所有的PDF文件
# \2021-D30
folder_path = r'G:\2023\2023小红头'
pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]

# 用于提取PDF第一页标题并重命名文件的函数
def extract_title_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        first_page = pdf.pages[0]  # 获取第一页
        words = first_page.extract_words()  # 提取所有文本元素及其坐标
        
        # 存储标题文本
        possible_titles = []
        
        for word in words:
            text = word['text']
            # print(f"text: {text}")
            top = word['top']
            # print(f"top: {top}")
            font_size = word['bottom'] - word['top']  # 字体大小通常可以通过top差值来估计
            # print(f"font_size:{font_size}")
            # 识别标题区域
            if 350 > top > 190:
                possible_titles.append((text, top))
        
        # 按照 top 值（Y坐标）排序，通常标题在页面的顶部
        possible_titles = sorted(possible_titles, key=lambda x: x[1])
        if possible_titles:
            # 拼接标题
            title = "".join([text.replace(" ", "") for text, _ in possible_titles]).strip()
            return title
        else:
            return None, None

# 处理每个PDF文件
for pdf_file in pdf_files:
  try:
    pdf_path = os.path.join(folder_path, pdf_file)
    
    # 提取标题和文号
    title = extract_title_from_pdf(pdf_path)
    if title:
        # 生成新文件名并重命名文件
        title = re.sub(r'[：“”、<>?()]', '', title)
        new_pdf_file = f"{title}.pdf"
        new_pdf_path = os.path.join(folder_path, new_pdf_file)
        
        # 检查新文件名是否已存在，如果已存在则避免重命名
        if not os.path.exists(new_pdf_path):
            # print(f"文件旧名字 {pdf_file} 文件新名字为 {new_pdf_file}")
            os.rename(pdf_path, new_pdf_path)
            print(f"文件 {pdf_file} 已重命名为 {new_pdf_file}")
        else:
            print(f"文件 {new_pdf_file} 已存在，跳过重命名")
    else:
        print(f"无法提取 {pdf_file} 的标题，跳过重命名")
  except Exception as e:
    print("捕获到异常:")
    traceback.print_exc()

