import os
import pdfplumber
import sys
import io
from PyPDF2 import PdfReader
import pandas as pd

# 设置标准输出为 UTF-8 编码
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
# 获取文件夹内所有的PDF文件
folder_path = r'E:\Python'
pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]

# 用于提取PDF第一页标题并重命名文件的函数
def extract_title_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        first_page = pdf.pages[0]  # 获取第一页
        words = first_page.extract_words()  # 提取所有文本元素及其坐标
        
        # 选择距离页面顶部较近且字体较大的文本作为标题
        possible_titles = []
        for word in words:
            # word包含了文本信息及其坐标 (x0, top, x1, bottom)
            text = word['text']
            # print(f'text:{text}')
            top = word['top']
            # print(f'top:{top}')
            font_size = word['bottom'] - word['top']  # 字体大小通常可以通过top差值来估计
            # print(f'font_size:{font_size}')
            # 调整条件
            if 500 > top > 400 and font_size > 19:  # 增加top的阈值，减小font_size的要求
                possible_titles.append((text, top))
        
        # 调试输出
        # print(f"提取的文本及其坐标: {possible_titles}")
        
        # 按照 top 值（Y坐标）排序，通常标题在页面的顶部
        possible_titles = sorted(possible_titles, key=lambda x: x[1])
        if possible_titles:
            # 拼接文本作为标题
            title = "".join([text.replace(" ", "") for text, _ in possible_titles]).strip()
            # print(f"文件 {pdf_file} 的标题为 {title}")
            
            # 使用正则表达式匹配 "关于" 开头，"批复" 结尾的部分
            match = re.search(r"关于.*批复", title)
            if match:
                title = match.group(0)  # 提取匹配的部分作为文件名
                return title
            else:
                return None
        else:
            return None

# 重命名处理每个PDF文件
for pdf_file in pdf_files:
    pdf_path = os.path.join(folder_path, pdf_file)
    
    # 提取标题
    title = extract_title_from_pdf(pdf_path)
    # print(f"文件 {pdf_file} 的标题为 {title}")
    if title:
        # 生成新文件名并重命名文件
        new_pdf_file = f"{title}.pdf"
        new_pdf_path = os.path.join(folder_path, new_pdf_file)
        
        # 检查新文件名是否已存在，如果已存在则避免重命名
        if not os.path.exists(new_pdf_path):
            os.rename(pdf_path, new_pdf_path)
            print(f"文件 {pdf_file} 已重命名为 {new_pdf_file}")
        else:
            print(f"文件 {new_pdf_file} 已存在，跳过重命名")
    else:
        print(f"无法提取 {pdf_file} 的标题，跳过重命名")

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