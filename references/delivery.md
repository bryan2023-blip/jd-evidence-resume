# 制作与交付检查

## 制作前

确认交付语言、目标职位、必要信息和最终选材。复用已批准版式；只在用户没有提供方向时选择简洁、可阅读、易解析的布局。

应届／初级一般一页，丰富经历可两页或按用户要求。不要求统一照片、固定模块数、固定条目数或所有行业都同一格式。优先删弱相关内容，保留正常字重、可读字号和留白。

## 工具与文件

- 在当前环境发现文档和 PDF 工具。可用 documents 与 pdf 技能时，读取当前版本并按说明执行；当前工具要求的操作标记、渲染及文件引用规则优先。
- 在 Codex desktop 用 load_workspace_dependencies 查运行时及可用库。不要硬编码某位求职者的主目录、依赖版本、字体或文档路径。
- 可用 python-docx 时写原生 DOCX；转换成 PDF 后，渲染为图片查看。如果需要其他方法，两份文件仍须内容一致、Word 可编辑、PDF 文字可选。
- 多文件或长任务保留内部内容稿、证据记录和检查记录，最终只交付需要的文件。
- 使用可追溯文件名，例如“姓名_公司_岗位_简历.docx”和同名 PDF。修订时保持版本配对，避免交付旧 PDF。
- 作品链接仅来自求职者确认或材料支持的地址；保留可点击链接。能访问时检查目标与作品是否匹配，无法访问则说明访问状态，不能把确认了链接语法说成作品已审核。

## 内容审核

逐条检查经历名称、性质、时间、个人贡献、方法与产出。

- 有来源的数字是否保留单位、基准和范围？
- 尚未执行的建议是否写成建议或企划？
- 模拟结果是否仍标模拟？
- 基础练习是否被升格为熟练实际操作？
- 团队成绩是否误写成个人全部成果？
- 重要资格缺口是否保留真实信息？
- 是否还残留占位符、内部备注或没核实的联系方式？
- 一条经历能否让读者看出为什么适合这项职责？
- 是否把与目标岗位无关的缺项逐条写入简历？保留必要的经验性质说明，正文优先呈现真实的已做能力；重要差距在交付说明中告知用户。

## 文件审核

全部页面都要渲染检查，不能只查看首页。

检查页数符合约定、无溢出或空白页、标题和日期对齐、照片不遮挡、中文字体与符号正常、文本不过密。检查 DOCX 可编辑，PDF 有可提取文字，姓名、电话、岗位、日期和关键数字在两个格式中一致。

在具备 pypdf、python-docx 的环境中，可以用以下思路自动检查关键文本和链接。此检查不替代版式查看或经历核实：

```python
import re
import unicodedata
from docx import Document
from pypdf import PdfReader

def normalize(text):
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', text))

document = Document(docx_path)
word_text = '\n'.join(p.text for p in document.paragraphs)
for table in document.tables:
    word_text += '\n' + '\n'.join(c.text for row in table.rows for c in row.cells)
pdf = PdfReader(pdf_path)
pdf_text = '\n'.join(p.extract_text() or '' for p in pdf.pages)
assert len(pdf.pages) == expected_pages
assert pdf_text.strip()
for text in confirmed_key_texts:
    assert normalize(text) in normalize(word_text)
    assert normalize(text) in normalize(pdf_text)
urls = []
for page in pdf.pages:
    for ref in page.get('/Annots', []):
        action = ref.get_object().get('/A', {})
        if action.get('/URI'):
            urls.append(str(action['/URI']))
for url in confirmed_urls:
    assert url in urls
```

正文、表格以外的页眉文本和 Word 超链接标签须另查对应文档结构。对 PDF 的正常换行空格可规范化，不能为让检测通过而放宽真正缺失的内容。

必要检查失败后，修复内容或排版，重新生成 PDF 并复查受影响页面。无新增变化或未解决问题时不反复跑无意义检查。

## 交付

最后给两个经过核实的文件入口，简短说明突出哪些能力。必要时附具体资格提醒，不附长篇内部分析。下载入口遵守环境当前约定。

没有实际生成或检查的部分明确标记为未完成；不把“转换成功”称为“可直接投递”。不自行投递，也不把这份简历用于另一个求职者。
