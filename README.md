# DocumentsTransformer

这是一个简单的 Python 小工具，可以将输入的文本或指定图片生成一个 Word 文档和对应的 PDF 文件。该项目基于 `docxtpl` 和 `win32com.client` 实现文档生成与转换功能。

## 功能说明
- 支持将任意文本插入到预设的 Word 模板中。
- 可选地插入一张图片到文档中。
- 自动生成 .docx 格式的 Word 文档。
- 利用 Microsoft Word 的 COM 接口将 .docx 转换为 .pdf 格式。

## 依赖库
确保安装以下依赖包：
```bash
pip install docxtpl pythoncom pywin32
```


> 注意：此工具依赖于 Windows 平台上的 Microsoft Office Word，因此仅支持在 Windows 系统上运行。

## 文件结构说明
- `text-to-word-and-pdf.py`: 主程序文件，包含核心逻辑。
- `Template/WordTemplate.docx`: Word 模板文件，用于渲染最终输出文档。
- `Image/image.png`: 示例图片路径。
- `Saved/`: 输出的 `.docx` 和 `.pdf` 文件保存目录。

## 使用方法
1. 准备好 Word 模板并放置在 `Template/` 目录下，命名为 `WordTemplate.docx`。
2. 替换 `Image/image.png` 为你自己的图片路径，或修改代码中的图片路径。
3. 运行程序：
   ```bash
   python DocumentsTransformer.py
   ```

4. 输出结果会打印在控制台上，同时文件保存至 `Saved/` 目录。

## 示例代码调用
```python
sample_text = "这是一个示例文本。"
sample_image = "Image/image.png"  # 如果不需要图片，可传入空字符串 ""
result = TextToWordAndPDF(sample_text, sample_image)
print(result)
```


## 方法说明
- [TextToWordAndPDF(text, image)](file://D:\PythonProject\text-to-word-and-pdf\text-to-word-and-pdf.py#L8-L46):
  - `text`: 需要插入到 Word 文档中的文本内容。
  - `image`: 图片路径（字符串），如果不需要插入图片，请传入空字符串 `""`。
  - 返回值：提示信息，显示 [.docx](file://D:\PythonProject\text-to-word-and-pdf\Saved\SavedWord1.docx) 和 [.pdf](file://D:\PythonProject\text-to-word-and-pdf\Saved\SavedPDF1.pdf) 文件保存路径。

## 注意事项
- 如果没有安装 Microsoft Word 或系统非 Windows，则无法使用 PDF 导出功能。
- 若模板文件或图片路径错误，可能导致程序异常中断，请确保路径正确。
- 生成的文件默认保存在 `Saved/` 目录中，如该目录不存在，程序会自动创建。

## 参考
[知乎文章：保姆级别docxtpl教程，你值得拥有](https://www.zhihu.com/tardis/zm/art/1888148733365044222?source_id=1005)

---

如有任何问题，欢迎提交 Issue 或 Pull Request！