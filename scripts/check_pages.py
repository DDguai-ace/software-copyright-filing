# -*- coding: utf-8 -*-
"""软著鉴别材料 PDF 自检：页数、每页行数、页眉关键字

用法：
    python check_pages.py <pdf> [--min-lines 50] [--expect-pages 60] [--header 软件全称]

- 源程序：--min-lines 50
- 说明书：--min-lines 30
- 不传 --expect-pages 则只报告实际页数
输出：每页行数概览 + 不达标页清单 + 结论
"""
import sys
import argparse

try:
    import pymupdf
except ImportError:
    import fitz as pymupdf


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("--min-lines", type=int, default=0)
    ap.add_argument("--expect-pages", type=int, default=0)
    ap.add_argument("--header", default="")
    ap.add_argument("--last-page-exempt", action="store_true", default=True)
    a = ap.parse_args()

    doc = pymupdf.open(a.pdf)
    n = doc.page_count
    print(f"文件：{a.pdf}")
    print(f"页数：{n}" + (f"（要求 {a.expect_pages} 页）" if a.expect_pages else ""))

    rows = []
    bad = []
    for i in range(n):
        txt = doc[i].get_text("text")
        lines = [ln for ln in txt.splitlines() if ln.strip()]
        rows.append(len(lines))
        is_last = i == n - 1
        if a.min_lines and not (is_last and a.last_page_exempt):
            if len(lines) < a.min_lines:
                bad.append((i + 1, len(lines)))

    print("每页行数：" + " ".join(f"{x}" for x in rows))
    if a.min_lines:
        print(f"最低要求：{a.min_lines} 行/页（末页豁免）")
        if bad:
            print("不达标页：")
            for p, c in bad:
                print(f"  - 第 {p} 页：{c} 行")
        else:
            print("行数：全部达标")

    if a.expect_pages and n != a.expect_pages:
        print(f"页数异常：实际 {n}，要求 {a.expect_pages}"
              + ("（不足 60 页时应全部提交，不算错）" if n < a.expect_pages else ""))

    if a.header:
        hit = sum(1 for i in range(n) if a.header in doc[i].get_text("text"))
        print(f"页眉关键字「{a.header}」命中 {hit}/{n} 页" + ("（正常）" if hit == n else "（有缺页，需检查）"))

    first = doc[0].get_text("text").strip().splitlines()
    if first:
        print("首页首行：" + first[0][:80])

    ok = not bad and (not a.expect_pages or n >= min(a.expect_pages, n))
    print("\n结论：" + ("通过自检" if ok else "存在待修项，见上"))
    return 0 if ok else 2


if __name__ == "__main__":
    sys.exit(main())
