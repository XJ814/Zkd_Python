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
folder_path = r'G:\2022\2022-D10'
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
            if 740 > top > 530 and font_size > 33:
                possible_titles.append((text, top))
            
            # 在指定区域内识别文号
            if top < 530 and font_size < 33:
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
        possible_titles = sorted(possible_titles, key=lambda x: x[1])
        if possible_titles:
            # 拼接标题
            title = "".join([text.replace(" ", "") for text, _ in possible_titles]).strip()
            return title, document_number, file_number
        else:
            return None, None

# 处理每个PDF文件
for pdf_file in pdf_files:
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

# 获取目录下所有的PDF文件
pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]

# 创建一个空列表，用来存储文件名和页数
data = []

# 获取每个PDF文件的页数
for pdf_file in pdf_files:
  pdf_path = os.path.join(folder_path, pdf_file)
  
  # 获取文件的最后修改时间
  file_modified_time = os.path.getmtime(pdf_path)
  
  # 去掉文件名中的扩展名
  file_name = os.path.splitext(pdf_file)[0]

  # 分割文件名为文件号、文号和文件名三部分
  try:
      parts = file_name.split('-', 2)  # 最多分割成三个部分
      if len(parts) == 3:
          file_number, document_number, file_name_part = parts
      else:
          # 如果分割后部分数少于3，进行适当处理
          file_number = parts[0]
          document_number = parts[1] if len(parts) > 1 else ''
          file_name_part = parts[2] if len(parts) > 2 else ''
  except Exception as e:
      # 处理异常情况
      file_number = file_name
      document_number = ''
      file_name_part = ''
  
  # 尝试读取PDF文件并获取页数
  try:
    with open(pdf_path, 'rb') as f:
      reader = PdfReader(f)
      num_pages = len(reader.pages)  # 获取页面数量
      
    # 将文号、文件名和页数添加到列表中
    # data.append([document_number, file_name_part, num_pages]) # 序号从 1 开始
    data.append([int(file_number),document_number, file_name_part, num_pages]) # 序号从 1 开始

  except Exception as e:
    print(f"无法读取文件 {pdf_file}: {e}")

# 按照文件号升序排序
data.sort(key=lambda x: x[0])  # 默认升序排序

# 添加序号列
for i, entry in enumerate(data, start=1):
    entry.insert(0, i)  # 在每一行的开头插入序号

# 将数据转换为DataFrame
df = pd.DataFrame(data, columns=['序号', '文件号', '文号', '文件名', '页数'])

# 将DataFrame导出到Excel
excel_path = r'G:\2022\2022-D10\目录.xlsx'  # 您可以修改路径
df.to_excel(excel_path, index=False, engine='openpyxl')

print(f'目录已生成并保存为 {excel_path}')