# python-docx 进阶参考

当 SKILL.md 的标准流程不够用时查阅。只收录 python-docx API 不直接支持、需要 XML 操作或容易写错的场景。按节取用，不要一次全读。

---

## 1. 表格进阶

### 合并单元格

```python
table.cell(0, 0).merge(table.cell(0, 2))  # 水平合并：第一行前 3 格
table.cell(0, 0).merge(table.cell(2, 0))  # 垂直合并：第一列前 3 行
```

### 单元格边框

```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def set_cell_border(cell, **kwargs):
    """
    用法: set_cell_border(cell, top={"sz": 12, "val": "single", "color": "000000"})
    可选边: top, bottom, left, right
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge, attrs in kwargs.items():
        elem = OxmlElement(f'w:{edge}')
        for attr_name, attr_val in attrs.items():
            elem.set(qn(f'w:{attr_name}'), str(attr_val))
        tcBorders.append(elem)
```

### 垂直居中

```python
from docx.enum.table import WD_ALIGN_VERTICAL

for row in table.rows:
    for cell in row.cells:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
```

---

## 2. 页眉页脚

### 页眉 + 页脚页码（基础）

```python
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

section = doc.sections[0]

# 页眉
header = section.header
header.is_linked_to_previous = False
header.paragraphs[0].text = "页眉文本"
header.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

# 页脚页码（PAGE 域）
footer = section.footer
footer.is_linked_to_previous = False
footer_para = footer.paragraphs[0]
footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

def add_page_field(para, instr):
    """向段落追加 PAGE / NUMPAGES 等域"""
    run = para.add_run()
    run._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>'))
    run = para.add_run()
    run._r.append(parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> {instr} </w:instrText>'))
    run = para.add_run()
    run._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>'))

add_page_field(footer_para, 'PAGE')
```

### "第 X 页 / 共 Y 页"

```python
footer_para.add_run("第 ")
add_page_field(footer_para, 'PAGE')
footer_para.add_run(" 页 / 共 ")
add_page_field(footer_para, 'NUMPAGES')
footer_para.add_run(" 页")
```

### 不同节使用不同页眉

```python
new_section = doc.add_section()
new_section.header.is_linked_to_previous = False
new_section.footer.is_linked_to_previous = False
doc.sections[1].header.paragraphs[0].text = "第二章"
```

### 首页不同页眉

```python
section.different_first_page_header_footer = True
section.first_page_header.paragraphs[0].text = "首页专用页眉"
```

---

## 3. 段落插入与删除

```python
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

def insert_paragraph_after(paragraph, text, style=None):
    """在指定段落之后插入（add_paragraph 只能追加到末尾）"""
    new_p = parse_xml(f'<w:p {nsdecls("w")}/>')
    paragraph._p.addnext(new_p)
    new_para = type(paragraph)(new_p, paragraph._parent)
    new_para.text = text
    if style:
        new_para.style = style
    return new_para

def delete_paragraph(paragraph):
    paragraph._element.getparent().remove(paragraph._element)
```

---

## 4. run 级中文字体设置

SKILL.md 的 `set_style_cjk_font` 作用于**样式**；对单个 run 微调字体用：

```python
def set_cn_font(run, font_cn='宋体', font_en='Times New Roman', size=None,
                bold=None, color=None):
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), font_en)
    rFonts.set(qn('w:eastAsia'), font_cn)
    rFonts.set(qn('w:hAnsi'), font_en)
    if size is not None:
        run.font.size = size
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
```

---

## 5. 跟踪修订（Track Changes 写入）

python-docx 无修订 API，需底层 XML。python-docx 保存时会原样保留它不认识的 XML，所以此法可行。修订元素必须放在 `w:p` 段落内部。

**插入标记**（新文字正常显示，带修订痕迹）：

```xml
<w:ins w:id="1" w:author="作者" w:date="2026-01-01T00:00:00Z">
  <w:r><w:t>新文字</w:t></w:r>
</w:ins>
```

**删除标记**（旧文字划线保留；注意必须用 `w:delText` 而非 `w:t`，这是最常见的错法）：

```xml
<w:del w:id="2" w:author="作者" w:date="2026-01-01T00:00:00Z">
  <w:r><w:delText>旧文字</w:delText></w:r>
</w:del>
```

**Python 插入示例：**

```python
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

INS_XML = (
    f'<w:ins {nsdecls("w")} w:id="1" w:author="AI" w:date="2026-01-01T00:00:00Z">'
    '<w:r><w:t>新文字</w:t></w:r></w:ins>'
)
p = doc.paragraphs[0]
p._p.append(parse_xml(INS_XML))
```

**三个坑：**
1. `w:id` 必须全文唯一，多处修订时用递增编号
2. **整段删除**时，段落标记（回车符）也要标删，在段落的 `w:pPr/w:rPr` 内加 `<w:del .../>`，否则会残留空段
3. date 必须是 xsd:dateTime 格式（`2026-01-01T00:00:00Z`）

**查看他人修订**不必用 XML，直接：`pandoc file.docx --track-changes=all -o out.md`

---

## 6. 模板结构分析

套模板前先分析源文档结构（样式、节、页眉页脚概览）：

```python
from docx import Document

doc = Document(path)
styles_used = sorted({p.style.name for p in doc.paragraphs if p.text.strip()})
print(f"段落: {sum(1 for p in doc.paragraphs if p.text.strip())} | 表格: {len(doc.tables)} | 节: {len(doc.sections)}")
print(f"使用的样式: {styles_used}")

for si, s in enumerate(doc.sections):
    print(f"节 {si}: {s.page_width}x{s.page_height}, "
          f"边距 上{s.top_margin} 下{s.bottom_margin} 左{s.left_margin} 右{s.right_margin}")
    for h in s.header.paragraphs:
        if h.text.strip():
            print(f"  页眉: {h.text.strip()[:60]}"); break
    for f in s.footer.paragraphs:
        if f.text.strip():
            print(f"  页脚: {f.text.strip()[:60]}"); break
```

---

## 7. 超链接

python-docx 没有超链接 API：

```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def add_hyperlink(paragraph, text, url):
    part = paragraph.part
    r_id = part.relate_to(
        url,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        is_external=True,
    )
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '0563C1')
    rPr.append(color)
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink
```

---

## 8. 做不到的事（建议替代方案）

| 需求 | python-docx 支持度 | 替代方案 |
|------|:---:|------|
| 浮动图片（文字环绕） | ❌ | 用表格布局实现图文混排，或在模板中预设 |
| 水印 | ❌ | 在模板文档中预设水印，用代码填充内容 |
| 真正的目录内容 | ❌ | 只能插 TOC 域代码占位，告知用户打开后刷新 |
