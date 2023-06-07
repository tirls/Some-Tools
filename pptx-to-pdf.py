from pptx import Presentation
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

def ppt_to_pdf(ppt_file, pdf_file):
    # 读取PPT文件
    presentation = Presentation(ppt_file)

    # 创建PDF对象
    pdf = canvas.Canvas(pdf_file, pagesize=letter)

    # 设置PDF页面尺寸
    pdf.setPageSize(letter)

    # 遍历每个幻灯片
    for slide_num, slide in enumerate(presentation.slides, start=1):
        # 创建一个新的页面
        pdf.showPage()

        # 渲染幻灯片内容到PDF页面上
        slide.export(pdf)

    # 保存PDF文件
    pdf.save()

# 示例用法
ppt_file = "presentation.pptx"
pdf_file = "presentation.pdf"
ppt_to_pdf(ppt_file, pdf_file)
