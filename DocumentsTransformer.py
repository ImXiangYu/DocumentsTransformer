import os
import pythoncom
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage
import win32com.client


def generate_document(this_doc, this_context, output_path, generate_pdf=False):
    """
    生成 Word 文档并根据需要导出为 PDF。

    :param this_doc: 渲染模板
    :param this_context: 包含模板变量的字典，如 {"text": "...", "image": "..."}
    :param output_path: Word 文件输出路径
    :param generate_pdf: 是否生成 PDF
    :return: 生成结果信息
    """
    # 替换模板中的内容
    this_doc.render(this_context)

    # 确保目录存在
    directory = os.path.dirname(output_path)
    if not os.path.exists(directory):
        os.makedirs(directory)

    # 保存 Word 文件
    this_doc.save(output_path)

    this_result = f"Word 已成功保存至 {output_path} 路径下"

    # 如果需要生成 PDF，则调用转换函数
    if generate_pdf:
        pdf_path = output_path.replace(".docx", ".pdf")
        word_to_pdf(output_path, pdf_path)
        this_result += f"\nPDF 已成功保存至 {pdf_path}"

    return this_result


def word_to_pdf(input_path, output_path):
    """
    将 Word 文档转换为 PDF。
    """
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    this_doc = word.Documents.Open(os.path.abspath(input_path))
    this_doc.SaveAs(os.path.abspath(output_path), FileFormat=17)
    this_doc.Close()
    word.Quit()
    pythoncom.CoUninitialize()


if __name__ == "__main__":
    doc = DocxTemplate("Template/WordTemplate.docx")
    # 用户自定义部分

    # 文本
    sample_text = "这是一个示例文本。"

    # 图像
    tmp_image = "Image/image.png"
    sample_image = InlineImage(doc, tmp_image, width=Mm(60))

    # 列表
    sample_list = ['a', 'b', 'c']

    # 字典
    sample_dict = {"key1": "value1", "key2": "value2", "key3": "value3"}

    # 条件判断
    sample_value = 10

    # 设定变量
    sample_list_for_length = [0, 1, 2, 3, 4, 5]

    # 行插入
    sample_list_for_row = [0, 1, 2, 3, 4, 5]

    # 列插入
    sample_list_for_col = [0, 1, 2, 3, 4, 5]

    # 获取列表下标
    sample_list_for_index = [0, 1, 2, 3, 4, 5]

    # 构建 context
    context = {
        "text": sample_text,
        "image": sample_image,
        "list": sample_list,
        "dict": sample_dict,
        "value": sample_value,
        "list_for_length": sample_list_for_length,
        "list_for_row": sample_list_for_row,
        "list_for_col": sample_list_for_col,
        "list_for_index": sample_list_for_index
    }

    # 输出路径
    output_word_path = "Saved/SavedDocument.docx"

    # 是否生成对应PDF
    generate_pdf_flag = True  # 控制是否生成 PDF

    # 生成文档
    result = generate_document(doc, context, output_word_path, generate_pdf_flag)
    print(result)
