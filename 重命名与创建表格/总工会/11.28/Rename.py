import os
import pdfplumber
import sys
import io
from PyPDF2 import PdfReader
import pandas as pd
import re

# 设置标准输出为 UTF-8 编码
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
# 获取文件夹内所有的PDF文件
# \2021-D30
folder_path = r'G:\总工会\18-20\2020\2020-Y'
pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]

# 用于提取PDF第一页标题并重命名文件的函数
def extract_title_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        first_page = pdf.pages[0]  # 获取第一页
        words = first_page.extract_words()  # 提取所有文本元素及其坐标
        
        # 用于保存文件序号,用于排序
        file_number = ""

        # 用于保存文号
        document_number = ""
        
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
            if 800 > top > 530 and font_size > 33:
                possible_titles.append((text, top))
            
            # 在指定区域内识别文号
            if 400 < top < 530:
                # 提取文号的不同部分
                # print(f"text: {text}")
                document_number += re.sub(r'[〔〕「」j了汇;\[\]]', '', text)
                # print(f"document_number: {document_number}")        

        # 重新拼接文号为: 宛XX[YYYY]文号X号
        if document_number:
          # 提取年份和编号
          year_and_number = re.findall(r'\d+', document_number[3:]) 
          # print(f"year_and_number:{year_and_number}")
          year = ''.join( year_and_number)[0:4]
          # print(f"year:{year}")
          number = ''.join( year_and_number)[4:]
          # print(f"number:{number}")
          file_number = number # 保存文件序号
          # 格式化后的字符串
          document_number = f"{document_number[:3]}〔{year}〕{number}号"
          # print(f"document_number: {document_number}")

        
        # 按照 top 值（Y坐标）排序，通常标题在页面的顶部
        # possible_titles = sorted(possible_titles, key=lambda x: x[1])
        if possible_titles:
            # 拼接标题
            title = "".join([text.replace(" ", "") for text, _ in possible_titles]).strip()
            return title, document_number, file_number
        else:
            return None, None
# 处理每个PDF文件
for pdf_file in pdf_files:
  try:
    pdf_path = os.path.join(folder_path, pdf_file)
    
    # 提取标题和文号
    title, document_number, file_number = extract_title_from_pdf(pdf_path)
    if title:
        # 生成新文件名并重命名文件
        title = title.replace(":",'').replace("“",'').replace("”",'').replace("、",'').replace('<','').replace('>','').replace('?','').replace('/','').replace("|",'').replace("(",'').replace(")",'')
        new_pdf_file = f"{file_number}-{document_number}-{title}.pdf"
        new_pdf_path = os.path.join(folder_path, new_pdf_file)
        
        # 检查新文件名是否已存在，如果已存在则避免重命名
        if not os.path.exists(new_pdf_path):
            print(f"文件旧名字 {pdf_file} 文件新名字为 {new_pdf_file}")
            os.rename(pdf_path, new_pdf_path)
            print(f"文件 {pdf_file} 已重命名为 {new_pdf_file}")
        else:
            print(f"文件 {new_pdf_file} 已存在，跳过重命名")
    else:
        print(f"无法提取 {pdf_file} 的标题，跳过重命名")
  except:
    print('An exception occurred')
