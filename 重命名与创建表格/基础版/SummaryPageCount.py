import os
from PyPDF2 import PdfReader

def count_pdf_pages_in_directory(directory):
  total_pages = 0
  
  # 遍历文件夹及子文件夹
  for root, dirs, files in os.walk(directory):
    for file in files:
      if file.lower().endswith('.pdf'):  # 只处理PDF文件
        pdf_path = os.path.join(root, file)
        try:
          # 打开PDF文件并读取页数
          with open(pdf_path, 'rb') as f:
            reader = PdfReader(f)
            total_pages += len(reader.pages)
        except Exception as e:
          print(f"无法读取文件 {pdf_path}: {e}")
  
  return total_pages

# 使用示例
directory = r'G:\总工会\18-20'
total_pages = count_pdf_pages_in_directory(directory)
print(f"文件夹及子文件夹内所有PDF文件的总页数: {total_pages}")