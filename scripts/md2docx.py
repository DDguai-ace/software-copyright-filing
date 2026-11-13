# -*- coding: utf-8 -*-
"""极简 Markdown -> DOCX 转换器（支持标题/表格/列表/引用/代码块/粗体/分隔线）"""
import re
import sys
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

CN_FONT = "微软雅黑"
CODE_FONT = "Consolas"


def set_run_font(run, size=10.5, bold=False, mono=False, color=None):
    run.font.size = Pt(size)
    run.font.bold = bold
    name = CODE_FONT if mono else CN_FONT
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    if color:
        run.font.color.rgb = RGBColor(*color)


def add_inline(par, text, size=10.5, bold=False):
    """处理 **粗体** 与 `代码`"""
    tokens = re.split(r"(\*\*.+?\*\*|`.+?`)", text)
    for tk in tokens:
        if not tk:
            continue
        if tk.startswith("**") and tk.endswith("**") and len(tk) > 4:
            r = par.add_run(tk[2:-2])
            set_run_font(r, size, True)
        elif tk.startswith("`") and tk.endswith("`") and len(tk) > 2:
            r = par.add_run(tk[1:-1])
            set_run_font(r, size - 0.5, bold, mono=True)
        else:
            r = par.add_run(tk)
            set_run_font(r, size, bold)


def flush_table(doc, rows):
    if not rows:
        return
    data = [[c.strip() for c in r.strip().strip("|").split("|")] for r in rows]
    data = [d for d in data if not all(re.fullmatch(r":?-{2,}:?", c or "-") for c in d)]
    if not data:
        return
    ncol = max(len(d) for d in data)
    data = [d + [""] * (ncol - len(d)) for d in data]
    tb = doc.add_table(rows=len(data), cols=ncol)
    tb.style = "Table Grid"
    tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(data):
        for j, cell in enumerate(row):
            c = tb.cell(i, j)
            c.text = ""
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            add_inline(p, cell, size=8.5, bold=(i == 0))
    doc.add_paragraph()


def apply_header_footer(doc, header_text="", with_page_number=True):
    """页眉：软件全称 + 版本号（软著鉴别材料要求）；页脚：第 X 页 共 Y 页"""
    from docx.oxml import OxmlElement

    def _field(par, code):
        r = par.add_run()
        f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
        it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = code
        f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
        r._r.append(f1); r._r.append(it); r._r.append(f2)
        r.font.size = Pt(9); r.font.name = "宋体"
        r._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    for sec in doc.sections:
        sec.header_distance = Cm(0.8)
        sec.footer_distance = Cm(0.8)
        if header_text:
            hp = sec.header.paragraphs[0]
            hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            hr = hp.add_run(header_text)
            hr.font.size = Pt(9)
            hr.font.name = "宋体"
            hr._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
            pPr = hp._p.get_or_add_pPr()
            pbdr = OxmlElement("w:pBdr")
            b = OxmlElement("w:bottom")
            b.set(qn("w:val"), "single"); b.set(qn("w:sz"), "4")
            b.set(qn("w:space"), "1"); b.set(qn("w:color"), "808080")
            pbdr.append(b); pPr.append(pbdr)
        if with_page_number:
            fp = sec.footer.paragraphs[0]
            fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for seg, code in (("第 ", " PAGE "), (" 页 共 ", None), (None, " NUMPAGES "), (" 页", None)):
                if seg:
                    r = fp.add_run(seg)
                    r.font.size = Pt(9); r.font.name = "宋体"
                    r._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
                if code:
                    _field(fp, code)


def convert(md_path, docx_path, header_text="", with_page_number=True):
    with open(md_path, "r", encoding="utf-8") as f:
        lines = f.read().splitlines()

    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = CN_FONT
    st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), CN_FONT)

    sec = doc.sections[0]
    sec.page_width = Cm(21.0)      # A4 —— python-docx 默认是 Letter（612×792pt），
    sec.page_height = Cm(29.7)     # 软著要求 A4（595×842pt），不设会一页少 50pt 高度、校验不过
    sec.top_margin = Cm(2.0)
    sec.bottom_margin = Cm(1.9)
    sec.left_margin = Cm(2.2)
    sec.right_margin = Cm(2.0)

    i, n = 0, len(lines)
    while i < n:
        line = lines[i]
        s = line.rstrip()

        if s.startswith("```"):
            i += 1
            buf = []
            while i < n and not lines[i].startswith("```"):
                buf.append(lines[i])
                i += 1
            i += 1
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(8)
            p.paragraph_format.left_indent = Pt(12)
            r = p.add_run("\n".join(buf))
            set_run_font(r, 9, mono=True)
            continue

        if s.startswith("|"):
            buf = []
            while i < n and lines[i].strip().startswith("|"):
                buf.append(lines[i])
                i += 1
            flush_table(doc, buf)
            continue

        if re.fullmatch(r"-{3,}|\*{3,}", s.strip()):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(6)
            pPr = p._p.get_or_add_pPr()
            from docx.oxml import OxmlElement
            pbdr = OxmlElement("w:pBdr")
            bottom = OxmlElement("w:bottom")
            bottom.set(qn("w:val"), "single")
            bottom.set(qn("w:sz"), "6")
            bottom.set(qn("w:color"), "BFBFBF")
            pbdr.append(bottom)
            pPr.append(pbdr)
            i += 1
            continue

        m = re.match(r"^(#{1,6})\s+(.*)$", s)
        if m:
            lvl = len(m.group(1))
            txt = m.group(2).strip()
            sizes = {1: 18, 2: 14, 3: 12, 4: 10.5, 5: 10.5, 6: 10}
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.space_before = Pt(8 if lvl == 1 else 6 if lvl == 2 else 4)
            pf.space_after = Pt(5 if lvl <= 2 else 3)
            pf.keep_with_next = lvl <= 2
            if lvl == 1:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_inline(p, txt, size=sizes.get(lvl, 11), bold=True)
            i += 1
            continue

        if s.startswith(">"):
            txt = s.lstrip(">").strip()
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Pt(14)
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
            add_inline(p, txt, size=9.5)
            i += 1
            continue

        m = re.match(r"^(\s*)[-*+]\s+(.*)$", s)
        if m:
            depth = len(m.group(1)) // 2
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.left_indent = Pt(18 + depth * 16)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            add_inline(p, m.group(2))
            i += 1
            continue

        m = re.match(r"^(\s*)\d+[.)]\s+(.*)$", s)
        if m:
            depth = len(m.group(1)) // 2
            p = doc.add_paragraph(style="List Number")
            p.paragraph_format.left_indent = Pt(18 + depth * 16)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            add_inline(p, m.group(2))
            i += 1
            continue

        if not s.strip():
            i += 1
            continue

        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.05
        p.paragraph_format.first_line_indent = Pt(0)
        add_inline(p, s.strip())
        i += 1

    apply_header_footer(doc, header_text, with_page_number)
    doc.save(docx_path)
    print("OK ->", docx_path)


if __name__ == "__main__":
    args = [a for a in sys.argv[1:]]
    header = ""
    if "--header" in args:
        k = args.index("--header")
        header = args[k + 1]
        del args[k:k + 2]
    convert(args[0], args[1], header_text=header)
