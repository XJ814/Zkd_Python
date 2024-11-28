import os
from PyPDF2 import PdfReader
import pandas as pd
import sys
import io

# 设置标准输出为 UTF-8 编码
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
# 获取目录下所有的PDF文件
folder_path = r'E:\Python'
pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]

# 创建一个空列表，用来存储文件名和页数
data = []

# 获取每个PDF文件的页数
for pdf_file in pdf_files:
    pdf_path = os.path.join(folder_path, pdf_file)
    
    # 去掉文件名中的扩展名
    file_name = os.path.splitext(pdf_file)[0]
    # print(f'正在读取文件 {file_name}')
    
    # 尝试读取PDF文件并获取页数
    try:
        with open(pdf_path, 'rb') as f:
            reader = PdfReader(f)
            num_pages = len(reader.pages)  # 获取页面数量
            
        # 将文件名和页数添加到列表中
        data.append([len(data) + 1, file_name, num_pages]) # 序号从 1 开始
    except Exception as e:
        print(f"无法读取文件 {pdf_file}: {e}")

# 将数据转换为DataFrame
df = pd.DataFrame(data, columns=['序号', '文件名', '页数'])

# 将DataFrame导出到Excel
excel_path = r'E:\Python\pdf目录.xlsx'  # 您可以修改路径
df.to_excel(excel_path, index=False, engine='openpyxl')

print(f'目录已生成并保存为 {excel_path}')