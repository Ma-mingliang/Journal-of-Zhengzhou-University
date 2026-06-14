---
name: zzu-journal-format
description: 将任意Word文档转换为郑州大学工学报格式。读取源Word文档的内容（标题、作者、摘要、正文、参考文献等），重新生成符合工学报排版规范的新Word文档。支持中英混合字体、Heading 1/2层级标题、单栏摘要+双栏正文+双栏参考文献+单栏英文摘要的分节结构、OMML公式插入。可选从PDF提取图片。
---

# 郑州大学工学报格式转换技能

基于 python-docx 将任意 Word 文档重新格式化为郑州大学工学报投稿格式。

---

## 适用场景

用户有一个 Word 文档（论文、报告等），需要转换为郑州大学工学报的排版格式。典型场景：
- 已有翻译好的中文 Word 文档，需要调整格式
- 从其他模板生成的文档，需要对齐到工学报规范
- 手动排版的文档，需要自动化格式化

**可选前置步骤：** 如果用户只有 PDF，可以先用 PyMuPDF 从 PDF 提取图片，再进行格式转换。

---

## 工作流程（5阶段）

### 阶段0：采集用户信息（必做）

**在生成文档之前，必须向用户询问以下信息：**

- **姓名** — 作者中文姓名
- **学号** — 学号
- **班级/专业** — 年级+专业名称

**标题自动抓取：** 从源文档中自动提取中英文标题，无需用户手动输入。如果源文档标题不清晰，再向用户确认。

**作者单位格式：** 默认为 `（郑州大学 电气与信息工程学院，河南 郑州 450001）`，如用户有特殊要求可修改。

**使用 AskUserQuestion 工具采集信息，示例：**

```json
{
  "questions": [
    {
      "question": "请输入您的姓名：",
      "header": "作者姓名",
      "options": [{"label": "（示例）张三", "description": "请输入您的中文姓名"}],
      "multiSelect": false
    },
    {
      "question": "请输入您的学号：",
      "header": "学号",
      "options": [{"label": "（示例）2025XXXXXXXXX", "description": "请输入您的学号"}],
      "multiSelect": false
    },
    {
      "question": "请输入您的班级/专业信息：",
      "header": "班级专业",
      "options": [{"label": "（示例）2025级XX专业", "description": "请输入年级+专业"}],
      "multiSelect": false
    }
  ]
}
```

### 阶段1：分析源文档

**必做步骤：**

1. **读取源 Word 文档**，提取以下信息：
   - 标题（中英文）— **自动抓取匹配**
   - 作者姓名 — 来自阶段0用户输入
   - 作者单位 — 默认或用户指定
   - 摘要（中英文）
   - 关键词（中英文）
   - 正文内容（按章节）
   - 参考文献
   - 图片（如有）

2. **识别文档结构**：
   - 一级标题（如 "0 引言"、"1 背景"）
   - 二级标题（如 "2.1 方法"）
   - 正文段落
   - 图片和图表标题
   - 公式

3. **确认用户需求**：
   - 是否需要翻译（英文→中文）
   - 是否有参考模板 docx
   - 是否需要从 PDF 提取图片

### 阶段2：图片处理（可选）

**如果源文档中已有图片**，可直接从源 docx 提取。

**如果只有 PDF**，用 PyMuPDF 提取：

```python
import fitz
import os

doc = fitz.open('论文.pdf')
fig_dir = 'figures'
os.makedirs(fig_dir, exist_ok=True)

for i in range(doc.page_count):
    page = doc[i]
    images = page.get_images(full=True)
    for j, img in enumerate(images):
        xref = img[0]
        base_image = doc.extract_image(xref)
        if base_image['width'] > 100 and base_image['height'] > 100:
            fname = f'page{i+1}_img{j}.{base_image["ext"]}'
            with open(os.path.join(fig_dir, fname), 'wb') as f:
                f.write(base_image['image'])
doc.close()
```

### 阶段3：生成目标文档

**核心思路：不修改源文档，而是新建文档，按工学报格式规范重新写入所有内容。**

1. **读取源文档内容**（用 python-docx 打开源 .docx）
2. **创建新文档**，设置样式和分节
3. **按格式规范写入**每个元素

### 阶段4：验证与调整

- 检查标题层级是否正确（Heading 1 / Heading 2）
- 检查字体是否正确（中文宋体/黑体，英文 TNR）
- 检查分栏是否正确（摘要单栏，正文双栏）
- 检查颜色是否全部黑色
- 检查参考文献格式

---

## 格式规范（必须严格遵守）

### 页面设置

| 参数 | 值 |
|------|-----|
| 纸张 | A4 (11906 x 16838 twips) |
| 上下页边距 | 1440 twips |
| 左右页边距 | 1800 twips |
| 栏间距 | 425 twips |

### 分节结构

| 节 | 内容 | 分栏 | 分节符 |
|----|------|------|--------|
| Section -1（可选） | 封面页（来自封面docx） | 保持原样 | nextPage |
| Section 0 | 标题 + 作者 + 单位 + 中文摘要 + 关键词 | 单栏 | nextPage（默认） |
| Section 1 | 正文（所有章节 + 图片） | 双栏 | continuous |
| Section 2 | 参考文献 | 双栏 | continuous |
| Section 3 | 英文标题 + 英文作者 + 英文单位 + 英文摘要 + 英文关键词 | 单栏 | nextPage |

### 封面页拼接（可选）

如果工作目录下存在封面 docx 文件（文件名含"封面"，如 `封面.docx`、`张三-封面.docx`），会自动拼接到生成文档的最前面。

**封面文件识别规则：**
- 文件名包含"封面"二字
- 扩展名为 `.docx`
- 排除临时文件（以 `~` 或 `~$` 开头）

**拼接方式：**
- 读取封面 docx 的所有元素（段落、表格、图片等）
- 复制到新文档的最前面
- 封面内容保持原始格式不变
- 封面之后添加一个分节符（nextPage）
- 然后接续正常的工学报格式内容

### 字体规范

| 元素 | 中文字体 | 英文字体 | 字号 | 其他 |
|------|---------|---------|------|------|
| 论文标题 | 黑体 | Times New Roman | 22pt (sz=44) | 居中，继承自 Heading 1 样式，不加粗 |
| 作者姓名 | 宋体 | Times New Roman | 14pt (sz=28) | 居中，加粗 |
| 作者单位 | 宋体 | Times New Roman | 9pt (sz=18) | 居中 |
| 摘要标签（"摘  要："） | 黑体 | Times New Roman | 9pt (sz=18) | 加粗 |
| 摘要内容 | 宋体 | Times New Roman | 9pt (sz=18) | 左缩进 230505 EMU，右缩进 255905 EMU |
| 关键词标签（"关键词："） | 黑体 | Times New Roman | 9pt (sz=18) | 加粗 |
| 关键词内容 | 宋体 | Times New Roman | 9pt (sz=18) | |
| 一级标题（"0 引言"） | 宋体(数字) + 黑体(文字) | Times New Roman | 14pt (sz=28) | Heading 1 样式，不加粗 |
| 二级标题（"2.1 方法"） | 黑体 | Times New Roman | 11pt (sz=21) | Heading 2 样式，不加粗 |
| 正文 | 宋体 | Times New Roman | 10.5pt (sz=21) | 首行缩进 266700 EMU |
| 图表标题 | 黑体 | Times New Roman | 9pt (sz=18) | 居中，加粗 |
| 参考文献 | 宋体 | Times New Roman | 10.5pt (sz=21) |
| 英文摘要标签（"Abstract: "） | — | Times New Roman | 9pt (sz=18) | 加粗 |
| 英文摘要内容 | — | Times New Roman | 9pt (sz=18) | |

### 段落间距

| 元素 | 段前 | 段后 | 行距 |
|------|------|------|------|
| 论文标题 | 0 twips | 0 twips | 240/auto（单倍） |
| Heading 1 样式 | 0 twips | 0 twips | — |
| Heading 1 段落 | — | — | 240/auto（单倍） |
| Heading 2 样式 | 0 twips | 0 twips | — |
| Heading 2 段落 | — | — | 240/auto（单倍） |
| 正文 | — | — | 继承 Normal 样式 |

**重要：** 论文标题和所有 Heading 1/2 的段前段后必须为 0 磅，标题段落需在段落级别显式设置 `before=0, after=0`。

### 颜色

- **所有文字颜色必须为黑色 (RGBColor(0, 0, 0))**
- 覆盖 Heading 1/2 样式默认的蓝色（#365F91 / #4F81BD）

---

## 核心函数模板

```python
from docx import Document
from docx.shared import Pt, Cm, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re

# ============================================================
# 格式常量
# ============================================================
SIZE_AUTHOR = 177800        # 14pt
SIZE_AFFILIATION = 114300   # 9pt
SIZE_HEADING1 = 177800      # 14pt
SIZE_HEADING2 = 133350      # 11pt
SIZE_BODY = 133350          # 10.5pt
SIZE_ABSTRACT = 114300      # 9pt
SIZE_CAPTION = 114300       # 9pt
SIZE_REF = 133350           # 10.5pt
SIZE_ENG_TITLE = 152400     # 12pt

FI_BODY = 266700            # 首行缩进
RI_BODY = 255905            # 右缩进

PAGE_W = 11906
PAGE_H = 16838
MARGIN_TB = 1440
MARGIN_LR = 1800
COL_SPACE = 425


# ============================================================
# 工具函数
# ============================================================
def set_font(run, ea='宋体', fn='Times New Roman', size=None, bold=False):
    """设置字体，强制黑色"""
    run.font.name = fn
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    if size:
        run.font.size = Emu(size)
    r = run._element
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        r.insert(0, rPr)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), ea)
    rFonts.set(qn('w:ascii'), fn)
    rFonts.set(qn('w:hAnsi'), fn)


def set_spacing(para, line=None, rule=None):
    pPr = para._element.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        para._element.insert(0, pPr)
    sp = pPr.find(qn('w:spacing'))
    if sp is None:
        sp = OxmlElement('w:spacing')
        pPr.append(sp)
    if line is not None:
        sp.set(qn('w:line'), str(line))
    if rule is not None:
        sp.set(qn('w:lineRule'), rule)


def set_indent(para, first=None, left=None, right=None):
    pPr = para._element.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        para._element.insert(0, pPr)
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    if first is not None:
        ind.set(qn('w:firstLine'), str(int(first / 635)))
    if left is not None:
        ind.set(qn('w:left'), str(int(left / 635)))
    if right is not None:
        ind.set(qn('w:right'), str(int(right / 635)))


def set_continuous(section):
    sectPr = section._sectPr
    type_el = sectPr.find(qn('w:type'))
    if type_el is None:
        type_el = OxmlElement('w:type')
        pgSz = sectPr.find(qn('w:pgSz'))
        if pgSz is not None:
            pgSz.addprevious(type_el)
        else:
            sectPr.append(type_el)
    type_el.set(qn('w:val'), 'continuous')


def setup_section(section, cols=1, col_space=COL_SPACE):
    sectPr = section._sectPr
    pgSz = sectPr.find(qn('w:pgSz'))
    if pgSz is None:
        pgSz = OxmlElement('w:pgSz')
        sectPr.append(pgSz)
    pgSz.set(qn('w:w'), str(PAGE_W))
    pgSz.set(qn('w:h'), str(PAGE_H))
    pgMar = sectPr.find(qn('w:pgMar'))
    if pgMar is None:
        pgMar = OxmlElement('w:pgMar')
        sectPr.append(pgMar)
    pgMar.set(qn('w:top'), str(MARGIN_TB))
    pgMar.set(qn('w:bottom'), str(MARGIN_TB))
    pgMar.set(qn('w:left'), str(MARGIN_LR))
    pgMar.set(qn('w:right'), str(MARGIN_LR))
    old_cols = sectPr.findall(qn('w:cols'))
    for c in old_cols:
        sectPr.remove(c)
    cols_elem = OxmlElement('w:cols')
    if cols > 1:
        cols_elem.set(qn('w:num'), str(cols))
        cols_elem.set(qn('w:space'), str(col_space))
    sectPr.append(cols_elem)


def make_heading_1(doc, text):
    p = doc.add_paragraph()
    p.style = doc.styles['Heading 1']
    m = re.match(r'^(\d+\s+)(.+)$', text)
    if m:
        run1 = p.add_run(m.group(1))
        set_font(run1, '宋体', 'Times New Roman', SIZE_HEADING1)
        run2 = p.add_run(m.group(2))
        set_font(run2, '黑体', 'Times New Roman', SIZE_HEADING1)
    else:
        run = p.add_run(text)
        set_font(run, '黑体', 'Times New Roman', SIZE_HEADING1)
    set_spacing(p, 240, 'auto')
    return p


def make_heading_2(doc, text):
    p = doc.add_paragraph()
    p.style = doc.styles['Heading 2']
    run = p.add_run(text)
    set_font(run, '黑体', 'Times New Roman', SIZE_HEADING2)
    set_spacing(p, 240, 'auto')
    return p


def make_body(doc, text):
    p = doc.add_paragraph()
    set_indent(p, first=FI_BODY, right=RI_BODY)
    run = p.add_run(text)
    set_font(run, '宋体', 'Times New Roman', SIZE_BODY)
    return p


def make_caption(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    set_font(run, '黑体', 'Times New Roman', SIZE_CAPTION, bold=True)
    return p


def make_image(doc, path, width_in=3.2):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(path, width=Cm(width_in * 2.54))
    return p


def make_ref(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_font(run, '宋体', 'Times New Roman', SIZE_REF)
    return p


# ============================================================
# 样式初始化
# ============================================================
def init_styles(doc):
    # Heading 1: sz=44 (22pt), 黑色, 不加粗, 段前段后0
    h1 = doc.styles['Heading 1']
    h1_rPr = h1.element.find(qn('w:rPr'))
    if h1_rPr is not None:
        sz = h1_rPr.find(qn('w:sz'))
        if sz is not None:
            sz.set(qn('w:val'), '44')
        szCs = h1_rPr.find(qn('w:szCs'))
        if szCs is not None:
            szCs.set(qn('w:val'), '44')
        color = h1_rPr.find(qn('w:color'))
        if color is not None:
            color.set(qn('w:val'), '000000')
            for a in ['w:themeColor', 'w:themeShade']:
                if color.get(qn(a)) is not None:
                    del color.attrib[qn(a)]
        for tag in ['w:b', 'w:bCs']:
            el = h1_rPr.find(qn(tag))
            if el is not None:
                h1_rPr.remove(el)
    # Heading 1 段间距: 段前段后0
    h1_pPr = h1.element.find(qn('w:pPr'))
    if h1_pPr is not None:
        sp = h1_pPr.find(qn('w:spacing'))
        if sp is None:
            sp = OxmlElement('w:spacing')
            h1_pPr.append(sp)
        sp.set(qn('w:before'), '0')
        sp.set(qn('w:after'), '0')

    # Heading 2: 黑色, 段前段后0
    h2 = doc.styles['Heading 2']
    h2_rPr = h2.element.find(qn('w:rPr'))
    if h2_rPr is not None:
        color = h2_rPr.find(qn('w:color'))
        if color is not None:
            color.set(qn('w:val'), '000000')
            for a in ['w:themeColor', 'w:themeShade']:
                if color.get(qn(a)) is not None:
                    del color.attrib[qn(a)]
    h2_pPr = h2.element.find(qn('w:pPr'))
    if h2_pPr is not None:
        sp = h2_pPr.find(qn('w:spacing'))
        if sp is not None:
            sp.set(qn('w:before'), '0')
            sp.set(qn('w:after'), '0')
```

---

## 从源 Word 文档读取内容

```python
def read_source_doc(path):
    """读取源Word文档，提取结构化内容"""
    doc = Document(path)
    content = {
        'title': '',
        'authors': '',
        'affiliation': '',
        'abstract_zh': '',
        'keywords_zh': '',
        'abstract_en': '',
        'keywords_en': '',
        'sections': [],      # [(heading, text), ...]
        'references': [],
        'images': [],        # [(path, caption), ...]
    }

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        style = para.style.name if para.style else 'None'

        # 根据实际文档结构识别各部分
        # （需要根据具体文档调整判断逻辑）
        if style == 'Heading 1':
            content['sections'].append(('h1', text))
        elif style == 'Heading 2':
            content['sections'].append(('h2', text))
        else:
            content['sections'].append(('body', text))

    return content
```

---

## 公式处理（OMML）

使用 pandoc 将 LaTeX 公式转换为 Word 原生 OMML 格式：

```python
import subprocess
import tempfile

def insert_formula(paragraph, latex):
    with tempfile.NamedTemporaryFile(mode='w', suffix='.tex', delete=False, encoding='utf-8') as f:
        f.write(f'\\documentclass{{article}}\\begin{{document}}${latex}$\\end{{document}}')
        tex_path = f.name
    docx_path = tex_path.replace('.tex', '.docx')
    subprocess.run(['pandoc', tex_path, '-o', docx_path, '--from=latex', '--to=docx'], check=True)
    temp_doc = Document(docx_path)
    for p in temp_doc.paragraphs:
        for elem in p._element:
            if elem.tag.endswith('}oMath') or elem.tag.endswith('}oMathPara'):
                paragraph._element.append(elem)
    os.unlink(tex_path)
    os.unlink(docx_path)

def insert_formula_centered(doc, latex):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    insert_formula(p, latex)
    return p
```

---

## 完整生成脚本结构

```python
def find_cover_docx(work_dir):
    """在工作目录下查找封面docx文件（文件名含'封面'）"""
    for f in os.listdir(work_dir):
        if f.endswith('.docx') and '封面' in f and not f.startswith('~'):
            return os.path.join(work_dir, f)
    return None


def merge_cover_page(doc, cover_path):
    """将封面docx的所有内容拼接到文档开头，保持原始格式，正确处理图片"""
    import io
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.opc.part import Part as OpcPart
    from docx.opc.packuri import PackURI

    cover = Document(cover_path)
    body = doc.element.body

    # 提取封面的图片数据
    cover_images = {}
    for rel_id, rel in cover.part.rels.items():
        if 'image' in rel.reltype and hasattr(rel.target_part, 'blob'):
            cover_images[rel_id] = (rel.target_part.blob, rel.target_part.content_type)

    first_element = body[0]

    for element in cover.element.body:
        if element.tag.endswith('}sectPr'):
            continue
        first_element.addprevious(element)

    # 修复图片引用
    for elem in body:
        if elem == first_element:
            break
        if elem.tag.endswith('}p'):
            blips = elem.findall('.//' + qn('a:blip'))
            for blip in blips:
                old_embed = blip.get(qn('r:embed'))
                if old_embed and old_embed in cover_images:
                    img_bytes, img_ct = cover_images[old_embed]
                    ext_map = {'image/png': 'png', 'image/jpeg': 'jpg', 'image/jpg': 'jpg',
                               'image/x-emf': 'emf', 'image/x-wmf': 'wmf'}
                    ext = ext_map.get(img_ct, 'png')
                    if 'wmf' in img_ct:
                        try:
                            from PIL import Image
                            img = Image.open(io.BytesIO(img_bytes))
                            png_buf = io.BytesIO()
                            img.save(png_buf, 'PNG')
                            img_bytes = png_buf.getvalue()
                            img_ct = 'image/png'
                            ext = 'png'
                        except Exception:
                            pass
                    max_num = 0
                    for rid in doc.part.rels.keys():
                        if rid.startswith('rId'):
                            try:
                                max_num = max(max_num, int(rid[3:]))
                            except ValueError:
                                pass
                    new_rid = f'rId{max_num + 1}'
                    img_partname = PackURI(f'/word/media/cover_{max_num + 1}.{ext}')
                    img_part = OpcPart(img_partname, img_ct, img_bytes, doc.part.package)
                    doc.part.relate_to(img_part, RT.IMAGE)
                    blip.set(qn('r:embed'), new_rid)

    # 封面后分节符
    new_sect = OxmlElement('w:sectPr')
    pgSz = OxmlElement('w:pgSz')
    pgSz.set(qn('w:w'), str(PAGE_W))
    pgSz.set(qn('w:h'), str(PAGE_H))
    new_sect.append(pgSz)
    first_element.addprevious(new_sect)


def make_title(doc, title, spacing_before=0, spacing_after=0):
    """论文标题：Heading 1 样式，22pt 黑体，居中，段前段后0"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.style = doc.styles['Heading 1']
    # 段落级别显式设置段前段后为0
    pPr = p._element.find(qn('w:pPr'))
    sp = pPr.find(qn('w:spacing'))
    if sp is None:
        sp = OxmlElement('w:spacing')
        pPr.append(sp)
    sp.set(qn('w:before'), str(spacing_before))
    sp.set(qn('w:after'), str(spacing_after))
    run = p.add_run(title)
    set_font(run, '黑体', 'Times New Roman')  # 继承样式 sz=44 (22pt)
    return p


def main():
    # 0. 配置（用户信息来自阶段0采集）
    WORK_DIR = r'工作目录'
    SOURCE_PATH = os.path.join(WORK_DIR, '源文档.docx')
    OUTPUT_PATH = os.path.join(WORK_DIR, '输出文档.docx')

    # 用户信息（从 AskUserQuestion 采集）
    AUTHOR_NAME = '用户姓名'       # ← 替换为采集到的姓名
    STUDENT_ID = '用户学号'        # ← 替换为采集到的学号
    CLASS_INFO = '年级+专业'       # ← 替换为采集到的班级
    AFFILIATION = '（郑州大学 电气与信息工程学院，河南 郑州 450001）'

    # 1. 读取源文档，自动抓取标题
    source = read_source_doc(SOURCE_PATH)
    title_zh = source.get('title', '')  # 自动从源文档提取
    title_en = source.get('title_en', '')

    # 2. 创建新文档
    doc = Document()
    init_styles(doc)
    if doc.paragraphs:
        el = doc.paragraphs[0]._element
        el.getparent().remove(el)

    # 3. 封面页拼接（可选）
    cover_path = find_cover_docx(WORK_DIR)
    if cover_path:
        merge_cover_page(doc, cover_path)

    # 4. Section 0: 标题 + 作者 + 摘要（单栏）
    setup_section(doc.sections[0], cols=1)

    # 标题（自动抓取，段前段后0）
    make_title(doc, title_zh)

    # 作者（用户提供的姓名）
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(AUTHOR_NAME)
    set_font(run, '宋体', 'Times New Roman', SIZE_AUTHOR, bold=True)

    # 单位
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(AFFILIATION)
    set_font(run, '宋体', 'Times New Roman', SIZE_AFFILIATION)

    # 中文摘要 + 关键词 ...
    # 英文摘要 + 关键词 ...

    # 8. 保存
    doc.save(OUTPUT_PATH)
```

---

## 常见问题

### Q1: 源文档结构不清晰，无法自动识别标题/摘要
**解决：** 让用户手动确认各部分内容，或读取源文档的段落样式名来判断

### Q2: 源文档中的图片如何保留
**解决：** 从源 docx 中提取图片（遍历 `doc.inline_shapes`），保存到临时目录，再插入新文档

### Q3: Heading 样式默认蓝色
**解决：** `init_styles()` 修改样式颜色 + `set_font()` 强制 `RGBColor(0, 0, 0)`

### Q4: Python 执行报错 exit code 49
**解决：** 使用完整路径 `E:/Anaconda/python.exe <file>`

### Q5: 公式渲染不正确
**解决：** 确保安装 pandoc：`conda install -c conda-forge pandoc`

### Q6: 二级标题段前段后间距不对
**解决：** 在 `init_styles()` 中设置 Heading 2 样式的 `before=0, after=0`

### Q12: 论文标题段前段后间距不为0
**解决：** 需要在两处设置：
1. `init_styles()` 中设置 Heading 1 样式的 `before=0, after=0`
2. 标题段落本身需在段落级别显式设置 `sp.set(qn('w:before'), '0'); sp.set(qn('w:after'), '0')`

仅修改样式不够，因为段落可能有独立的间距覆盖。

### Q7: 封面docx没有被拼接
**解决：** 确保封面文件名包含"封面"二字（如 `封面.docx`、`张三-封面.docx`），且不是临时文件（不以 `~` 或 `~$` 开头）

### Q8: 封面拼接后格式错乱
**解决：** 封面内容是原样复制的，不会重新格式化。如果封面格式有问题，需要先调整封面docx本身的格式

### Q9: 双栏布局中图片只占一栏宽度
**解决：** 使用 `make_full_width_image()` 函数，通过连续分节符临时切换单栏来插入全宽图片和图注。详见下方"双栏全宽图片"章节。

### Q10: 封面docx中的图片没有显示
**解决：** `merge_cover_page()` 使用 `element.addprevious()` 复制XML元素时，图片的 rId 引用关系不会自动转移。需要在复制后遍历封面段落中的 `a:blip` 元素，将图片数据重新注册到新文档的 rels 中，并更新 rId 引用。如果是 WMF 格式，需先转为 PNG。

### Q11: 如何判断图片使用全宽还是单栏
**解决：** 根据图片的宽高比（aspect ratio）判断：
- aspect > 2.0 → 使用 `make_full_width_image()`（跨栏显示）
- aspect ≤ 2.0 → 使用 `make_image()` + `make_caption()`（单栏显示）

```python
from PIL import Image
img = Image.open(path)
w, h = img.size
if w / h > 2.0:
    make_full_width_image(doc, path, caption)
else:
    make_image(doc, path, width_in=3.0)
    make_caption(doc, caption)
```

---

## 文件结构

```
工作目录/
├── 源文档.docx              # 原始Word文档（只读）
├── 封面.docx                # 封面页（可选，文件名含"封面"）
├── 参考模板.docx            # 格式参考（可选）
├── figures/                 # 图片目录（可选）
├── generate_formatted.py    # 生成脚本
└── 输出文档.docx            # 最终输出
```

---

## 依赖

```bash
pip install python-docx lxml
# 可选：从PDF提取图片
pip install PyMuPDF
# 可选：OMML公式
conda install -c conda-forge pandoc
```

---

**技能版本：** 1.3.0
**兼容平台：** Claude Code (CLI/Desktop/Web)
**核心依赖：** python-docx, lxml
**Python环境：** Anaconda (E:/Anaconda/python.exe)
