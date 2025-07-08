# DocumentsTransformer.py
import os
import pythoncom
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage
import win32com.client


def TextToWordAndPDF(text, image):
    doc = DocxTemplate("Template/WordTemplate.docx")

    if not image:
        insert_image = ""
    else:
        insert_image = InlineImage(doc, image, width=Mm(60))

    context = {
        "text": text,
        "image" : insert_image
    }

    doc.render(context)

    word_filename = "SavedWord1.docx"
    pdf_filename = "SavedPDF1.pdf"

    # 如果不存在Saved文件夹，就创建
    if not os.path.exists("Saved"):
        os.mkdir("Saved")

    word_path = "Saved/" + word_filename
    pdf_path = "Saved/" + pdf_filename

    doc.save(word_path)

    def word_to_pdf(input_path, output_path):
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        this_doc = word.Documents.Open(os.path.abspath(input_path))
        this_doc.SaveAs(os.path.abspath(output_path), FileFormat=17)
        this_doc.Close()
        word.Quit()
        pythoncom.CoUninitialize()

    word_to_pdf(word_path, pdf_path)
    return f"Word 已成功保存至 {word_path} 路径下 \n PDF 已成功保存至 {pdf_path}"

if __name__ == "__main__":
    # 示例文本和图片路径
    sample_text = "这是一个示例文本。"
    sample_image = "Image/image.png"  # 替换为实际的图片路径

    result = TextToWordAndPDF(sample_text, sample_image)
    print(result)
