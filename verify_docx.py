# -*- coding: utf-8 -*-
"""Verify generated docx matches reference format."""
from docx import Document
from docx.oxml.ns import qn
import sys

def analyze(path, label):
    doc = Document(path)
    print(f'\n{"="*60}')
    print(f'  {label}: {path}')
    print(f'{"="*60}')

    # Sections
    print(f'\n--- SECTIONS ({len(doc.sections)}) ---')
    for i, sec in enumerate(doc.sections):
        cols = sec._sectPr.find(qn('w:cols'))
        col_num = cols.get(qn('w:num')) if cols is not None else '1'
        col_space = cols.get(qn('w:space')) if cols is not None else 'N/A'
        pgSz = sec._sectPr.find(qn('w:pgSz'))
        w = pgSz.get(qn('w:w')) if pgSz is not None else 'N/A'
        h = pgSz.get(qn('w:h')) if pgSz is not None else 'N/A'

        # section break type
        stype = 'continuous'
        pgMar = sec._sectPr.find(qn('w:pgMar'))
        t = pgMar.get(qn('w:top')) if pgMar is not None else '?'
        b = pgMar.get(qn('w:bottom')) if pgMar is not None else '?'
        l = pgMar.get(qn('w:left')) if pgMar is not None else '?'
        r = pgMar.get(qn('w:right')) if pgMar is not None else '?'

        # check if previous section has type attr
        type_el = sec._sectPr.find(qn('w:type'))
        if type_el is not None:
            stype = type_el.get(qn('w:val'), 'nextPage')
        print(f'  Sec {i}: {col_num}col, space={col_space}, page={w}x{h}, mar T={t} B={b} L={l} R={r}, type={stype}')

    # Paragraphs (first 80 with content)
    print(f'\n--- PARAGRAPHS (first 80 with text) ---')
    count = 0
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        count += 1
        if count > 80:
            break

        style = para.style.name if para.style else 'None'
        pf = para.paragraph_format
        fi = pf.first_line_indent
        li = pf.left_indent
        ri = pf.right_indent
        jc = para.alignment

        # Get run fonts
        run_info = []
        for r in para.runs[:3]:
            ea = r._element.find(qn('w:rPr'))
            fn = ea.find(qn('w:rFonts')).get(qn('w:eastAsia')) if ea is not None and ea.find(qn('w:rFonts')) is not None else '?'
            sz = r._element.find(qn('w:rPr'))
            szv = None
            if sz is not None:
                szel = sz.find(qn('w:sz'))
                if szel is not None:
                    szv = szel.get(qn('w:val'))
                else:
                    szv2 = sz.find(qn('w:szCs'))
                    if szv2 is not None:
                        szv = szv2.get(qn('w:val'))
            bold = r.bold
            fname = r.font.name or '?'
            run_info.append(f'ea={fn} fn={fname} sz={szv} b={bold}')

        txt = text[:60].replace('\n', ' ')
        print(f'  [{i}] {style} | {" | ".join(run_info)} | {txt}')

    print()

analyze(r'E:\Code\Embed\作业\智能机器人控制\马明亮-202522852018965.docx', 'REFERENCE')
analyze(r'E:\Code\Embed\作业\智能机器人控制\TD-MPC2-中文翻译-郑州大学工学报格式.docx', 'GENERATED')
