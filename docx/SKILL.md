---
name: docx
description: "创建和编辑 Word 文档。当用户需要创建、编辑、套模板或转换 Word/docx/doc/docm/dotx 文件时使用此 skill。不适用于：仅读取/查看文档内容（用 fetch 工具）、PDF、电子表格或与文档生成无关的任务。"
---

# DOCX Skill

## 路由

| 用户想要... | 路由 |
|---|---|
| 从零创建文档（"做一个/写一个/生成一个"） | → 创建 |
| 读取 / 分析文档（"看一下/提取/分析"） | → 读取 |
| 编辑现有文档（"改一下/替换/更新"） | → 编辑 |
| 套模板 / 统一样式（"按这个模板/统一风格"） | → 模板 |
| .doc → .docx 或 → PDF（"转成/导出"） | → 转换 |

---

## 创建

**标准流程：** 确定文档类型 → 按下方模板写脚本 → 保存并告知输出路径。

**基础模板（中文文档）：**

```python
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

doc = Document()

# ── 页面设置 ──
section = doc.sections[0]
section.page_width, section.page_height = Cm(21), Cm(29.7)
section.top_margin = section.bottom_margin = Cm(2.54)
section.left_margin = section.right_margin = Cm(3.18)

# ── 中文字体必须双设：font.name 只管西文，w:eastAsia 才管中文 ──
def set_style_cjk_font(style, cn_font, en_font):
    style.font.name = en_font
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    for attr, val in [('w:eastAsia', cn_font), ('w:ascii', en_font), ('w:hAnsi', en_font)]:
        rFonts.set(qn(attr), val)

# ── 清除标题样式的蓝色主题色和 Title 底部横线 ──
def clean_heading_style(style):
    style.font.color.rgb = RGBColor(0, 0, 0)
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is not None:
        for attr in ['w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme', 'w:cstheme']:
            if rFonts.get(qn(attr)):
                del rFonts.attrib[qn(attr)]
    pPr = style.element.find(qn('w:pPr'))
    if pPr is not None:
        pBdr = pPr.find(qn('w:pBdr'))
        if pBdr is not None:
            pPr.remove(pBdr)

# ── 正文样式 ──
normal = doc.styles['Normal']
normal.font.size = Pt(12)
set_style_cjk_font(normal, '宋体', 'Times New Roman')
normal.paragraph_format.line_spacing = 1

# ── 标题样式 ──
for name in ['Title'] + [f'Heading {i}' for i in range(1, 4)]:
    set_style_cjk_font(doc.styles[name], '黑体', 'Arial')
    clean_heading_style(doc.styles[name])

# ── 标题与正文 ──
doc.add_heading('文档标题', level=0).alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_paragraph('正文内容')

# ── 保存 ──
os.makedirs(os.path.dirname(output_path), exist_ok=True)
doc.save(output_path)
```

**常用元素：**

```python
doc.add_paragraph('条目', style='List Bullet')      # 无序列表（'List Number' 为有序）
table = doc.add_table(rows=3, cols=4, style='Table Grid')
doc.add_picture('image.png', width=Cm(15))          # 图片
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)  # 分页
```

页眉页脚、页码、"第 X 页 / 共 Y 页"等见 `references/python_docx_guide.md` 第 2 节。

---

## 读取

| 场景 | 方式 |
|------|------|
| 纯文本快速读取 | `pandoc file.docx -t plain` |
| 提取修订/批注内容 | `pandoc file.docx --track-changes=all -o out.md` |
| 需要样式/结构信息 | python-docx（下方代码） |

```python
from docx import Document

doc = Document("input.docx")
print(f"段落: {len(doc.paragraphs)} | 表格: {len(doc.tables)} | 节: {len(doc.sections)}")

for i, para in enumerate(doc.paragraphs):       # 段落结构
    if para.text.strip():
        print(f"[{i}] style={para.style.name} | {para.text[:80]}")

for t, table in enumerate(doc.tables):          # 表格内容
    for row in table.rows:
        print([cell.text[:30] for cell in row.cells])
```

---

## 编辑

**标准流程：** 打开并输出段落索引和样式（同读取）→ 执行编辑 → **保存到新路径，不覆盖原文件**。

**文本替换（必须处理 run 边界，直接赋值 `paragraph.text` 会丢格式）：**

```python
def replace_text_in_paragraph(paragraph, old_text, new_text):
    full_text = paragraph.text
    if old_text not in full_text:
        return False
    for run in paragraph.runs:
        if old_text in run.text:
            run.text = run.text.replace(old_text, new_text)
            return True
    # 跨 run：合并到第一个 run
    new_full = full_text.replace(old_text, new_text)
    for run in paragraph.runs:
        run.text = ""
    paragraph.runs[0].text = new_full
    return True

# 替换所有位置（段落 + 表格内单元格）
for para in doc.paragraphs:
    replace_text_in_paragraph(para, old, new)
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                replace_text_in_paragraph(para, old, new)
```

**关键原则：**
- 插入/删除段落用 XML 操作（见参考第 3 节），`add_paragraph()` 只能追加到末尾
- 别忘了表格内的文本
- 保留原有样式，不要主动修改

---

## 模板

1. 打开模板文档，提取样式定义（结构分析代码见参考第 6 节）
2. 确定源文档样式 → 模板样式的映射关系
3. 遍历段落/表格，应用映射后的样式
4. 保存

进阶用法（复制样式 XML、不同样式名映射）见 `references/python_docx_guide.md` 第 6 节。

---

## 转换

```bash
which libreoffice || which soffice   # 先检查可用性

libreoffice --headless --convert-to docx input.doc           # .doc → .docx
libreoffice --headless --convert-to pdf input.docx           # .docx → PDF
libreoffice --headless --convert-to docx --outdir ./output/ *.doc   # 批量
```

---

## 排版规范

用户指定了格式就按用户的；没指定时用以下默认值：

| 属性 | 中文默认值 | 英文默认值 |
|------|-----------|-----------|
| 页面 | A4, 上下 2.54cm 左右 3.18cm | Letter, 1 inch |
| 正文 | 宋体 + Times New Roman, 12pt, 单倍行距 | Calibri, 11pt |
| 标题 | 黑体 + Arial | Calibri |
| 首行缩进 | 2 字符 (Pt(24)，基于 12pt 字号) | 无 |

标题层级：大标题 `level=0`（Title，居中），章 `level=1`，节 `level=2`。不要跳级。

---

## 防坑规则

1. **中文字体双设**：`font.name` 只管西文，中文必须 XML 层设 `w:eastAsia` → 用创建模板中的 `set_style_cjk_font`；run 级见参考第 4 节
2. **标题样式自带蓝色和底部横线** → 用 `clean_heading_style` 清除
3. **单元格底色没有 API**，必须用 XML：
   ```python
   from docx.oxml.ns import nsdecls
   from docx.oxml import parse_xml
   cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="4472C4"/>'))
   ```
4. **列表样式名精确**：`'List Bullet'`、`'List Number'`；`'Bullet List'`、`'bullet'` 会 KeyError
5. **目录只是占位符**：TOC 域代码不会自动生成，需告知用户打开后 Ctrl+A → F9 刷新
6. **保存前确保目录存在**：`os.makedirs(os.path.dirname(output_path), exist_ok=True)`

---

## 参考资料

`references/python_docx_guide.md`（按需读取，不要一次全读）：
1. 表格进阶（合并单元格、边框、垂直居中）
2. 页眉页脚（页码、多节不同页眉、首页不同、"第 X 页 / 共 Y 页"）
3. 段落插入与删除（XML 操作）
4. run 级中文字体设置
5. 跟踪修订（Track Changes 写入）
6. 模板结构分析、超链接
7. python-docx 做不到的事及替代方案
