import os
import win32com.client as win32
from docx import Document

def convert_to_pdf(input_path, output_path):
    # 创建Word应用程序实例
    word_app = win32.gencache.EnsureDispatch('Word.Application')
    
    try:
        # 打开Word文档
        doc = word_app.Documents.Open(input_path)
        
        # 将文档保存为PDF
        doc.SaveAs(output_path, FileFormat=17)
        
        # 关闭文档
        doc.Close()
    except Exception as e:
        print("转换失败:", str(e))
    finally:
        # 关闭Word应用程序
        word_app.Quit()

# 示例用法
input_file = 'input.docx'  # 输入的Word文档路径
output_file = 'output.pdf'  # 输出的PDF文件路径

convert_to_pdf(input_file, output_file)
