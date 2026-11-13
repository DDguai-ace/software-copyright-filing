# -*- coding: utf-8 -*-
"""把「源代码.docx」重排为符合软著要求的源程序鉴别材料

用法：
    python build_source.py <源docx> <输出docx> <软件全称> [版本号]

规则（实测可稳定得到 60 行/页）：
- A4 21.0×29.7cm，边距 上下1.2 左1.5 右1.2cm
- Consolas 8pt，固定行距 13.2pt，段前段后 0
- 去掉空行（保证每一行都是真实代码行，行数可核）
- 页眉「软件全称  V1.0」，页脚「第 X 页 共 Y 页」
"""
import os
import re
import sys
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT = "Consolas"
META = re.compile(r"源程序|全量提交|全部提交|鉴别材料|共\s*\d+\s*行|约\s*\d+\s*页")


def add_field(par, code):
    r = par.add_run()
    f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = code
    f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
    r._r.append(f1); r._r.append(it); r._r.append(f2)
    return r


def build(src, out, title, ver="V1.0"):
    d = Document(src)
    # 段落内可能含 <w:br/>（p.text 返回 \n），必须按行拆开，否则行数统计与分页都会失真
    lines = [l.rstrip() for p in d.paragraphs for l in p.text.split("\n")]

    # 去掉开头的标题/统计行
    while lines:
        t = lines[0].strip()
        if not t:
            lines.pop(0)
        elif META.search(t) or t.startswith(title) or t == f"{title} {ver}":
            lines.pop(0)
        else:
            break
    # 去空行
    lines = [l for l in lines if l.strip()]

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin = Cm(1.2); sec.bottom_margin = Cm(1.2)
    sec.left_margin = Cm(1.5); sec.right_margin = Cm(1.2)
    sec.header_distance = Cm(0.55); sec.footer_distance = Cm(0.55)

    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(8)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    pf = st.paragraph_format
    pf.space_before = Pt(0); pf.space_after = Pt(0)
    pf.line_spacing = Pt(13.2); pf.line_spacing_rule = 4

    hp = sec.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hr = hp.add_run(f"{title}  {ver}")
    hr.font.size = Pt(9); hr.font.name = "宋体"
    hr._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    pPr = hp._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    b.set(qn("w:val"), "single"); b.set(qn("w:sz"), "4")
    b.set(qn("w:space"), "1"); b.set(qn("w:color"), "808080")
    pbdr.append(b); pPr.append(pbdr)

    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for seg, code in (("第 ", " PAGE "), (" 页 共 ", None), (None, " NUMPAGES "), (" 页", None)):
        if seg:
            r = fp.add_run(seg); r.font.size = Pt(9); r.font.name = "宋体"
            r._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
        if code:
            add_field(fp, code)

    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    for l in lines:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = Pt(13.2)
        p.paragraph_format.line_spacing_rule = 4
        r = p.add_run(l)
        r.font.size = Pt(8); r.font.name = FONT
        r._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    doc.save(out)
    print(f"OK {out}\n   行数 {len(lines)}，最长行 {max((len(l) for l in lines), default=0)}")


if __name__ == "__main__":
    if len(sys.argv) < 4:
        print(__doc__); sys.exit(1)
    build(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4] if len(sys.argv) > 4 else "V1.0")
