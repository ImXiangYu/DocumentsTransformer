# DocumentsTransformer

将信息插入 Word(docx) 指定位置，并生成 PDF！

这是一个简单的 Python 小工具，可以将文字、图像、列表、字典等插入 Word(docx) 的指定位置，并生成对应的 PDF 文件。  
该项目基于 `docxtpl` 和 `win32com.client` 实现文档生成与PDF转换功能。  
自带部分常见 docxtpl Word 模板，适用于大多数场景。详见 Template/WordTemplate.docx。

## 功能说明
- 包含常见 docxtpl 模板，可以自行按需使用。
  - 可以将文字、图像、列表、字典等插入 Word 的指定位置。
  - 支持条件判断、设定变量、表格的行插入和列插入、获取列表的下标。（详见模板: Template/WordTemplate.docx）
- 自动生成 .docx 格式的 Word 文档。
- 可以选择使用 Microsoft Word 的 COM 接口生成 PDF。


## 依赖库
确保安装以下依赖包：
```bash
pip install docxtpl pythoncom pywin32
```
> 注意：此工具依赖于 Windows 平台上的 Microsoft Office Word，因此仅支持在 Windows 系统上运行。


## 文件结构说明
- `DocumentsTransformer.py`: 主程序文件，包含核心逻辑。
- `Template/WordTemplate.docx`: Word 模板文件，用于渲染最终输出文档。
- `Image/image.png`: 示例图片路径。
- `Saved/`: 输出的 `.docx` 和 `.pdf` 文件保存目录。


## 使用方法
1. 查看 Template/WordTemplate.docx 中对应的模板，编排自己想要的样式。
2. 在 DocumentsTransformer.py 中按需选择自己需要的Sample，并将所需数据写入 context 中。
3. 对 WordTemplate.docx 和 context 进行按需修改。（或导入自己的 Word 模板）
4. 运行程序：
   ```bash
   python DocumentsTransformer.py
   ```
5. Word 和 PDF 会保存至 `Saved/` 目录。
6. 注意：如有多个变量，请确保变量名不重复。（见举例说明）


## 举例说明
例如，要生成一个包含三个 text 文本的 Word 文档，并生成对应的 PDF 文件，可以按照以下步骤操作：
1. 创建一个包含三个 text 文本模板的 Word 文档。（即文档中包含{{text1}}, {{text2}}, {{text3}}）
2. 按如下方式编写context：
    ```python
   context = {
        "text1": "这是第一个文本",
        "text2": "这是第二个文本",
        "text3": "这是第三个文本"
   }
   ```
3. 设置是否要创建PDF：`generate_pdf_flag = True  # 控制是否生成 PDF`
4. 运行DocumentsTransformer.py：
   ```bash
   python DocumentsTransformer.py
   ```
5. Word 和 PDF 文件会默认保存至 `Saved/` 目录。（可在文件中进行修改）
6. 其他模板类似，只需要按照步骤操作即可。


## 注意事项
- 如果没有安装 Microsoft Word 或系统非 Windows，则无法使用 PDF 导出功能。
- 若模板文件或图片路径错误，可能导致程序异常中断，请确保路径正确。
- 生成的文件默认保存在 `Saved/` 目录中，如该目录不存在，程序会自动创建。


## 参考
[知乎文章：保姆级别docxtpl教程，你值得拥有](https://www.zhihu.com/tardis/zm/art/1888148733365044222?source_id=1005)

---

如有任何问题，欢迎提交 Issue 或 Pull Request！