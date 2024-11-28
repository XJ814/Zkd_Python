import os
import pdfplumber
import sys
import io
from PyPDF2 import PdfReader
import pandas as pd
import re

# 设置标准输出为 UTF-8 编码
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 获取目录下所有的PDF文件
folder_path = r'G:\2022\2022-D10'
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