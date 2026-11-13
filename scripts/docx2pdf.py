# -*- coding: utf-8 -*-
"""批量把 .docx/.doc 转为 PDF（Windows，走 Word COM，失败回退 WPS）

用法：
    python docx2pdf.py <输入目录或文件> [输出目录]

- 输入目录：递归/非递归均可（默认只处理该目录下的文件），输出默认同目录
- 输入文件：输出到同名 .pdf
- 输出目录不存在会自动创建
"""
import os
import sys
import glob
import time


def _word_app():
    import win32com.client as win32
    try:
        app = win32.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        return app
    except Exception:
        app = win32.DispatchEx("KWPS.Application")  # WPS 兜底
        app.Visible = False
        app.DisplayAlerts = 0
        return app


def convert(files, out_dir):
    if not files:
        print("没有可转换的文件")
        return []
    app = _word_app()
    done = []
    try:
        for src in files:
            src = os.path.abspath(src)
            stem = os.path.splitext(os.path.basename(src))[0]
            dst = os.path.join(out_dir, stem + ".pdf")
            if os.path.exists(dst) and os.path.getmtime(dst) >= os.path.getmtime(src):
                print("跳过（已是最新）:", dst)
                done.append(dst)
                continue
            doc = app.Documents.Open(src, ReadOnly=True)
            try:
                doc.ExportAsFixedFormat(dst, 17)  # 17 = wdExportFormatPDF
            finally:
                doc.Close(False)
            print("OK:", dst)
            done.append(dst)
            time.sleep(0.2)
    finally:
        try:
            app.Quit()
        except Exception:
            pass
    return done


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    target = sys.argv[1]
    if os.path.isdir(target):
        files = sorted(
            glob.glob(os.path.join(target, "*.docx"))
            + glob.glob(os.path.join(target, "*.doc"))
        )
        files = [f for f in files if not os.path.basename(f).startswith("~$")]
        out_dir = sys.argv[2] if len(sys.argv) > 2 else target
    else:
        files = [target]
        out_dir = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(os.path.abspath(target))
    out_dir = os.path.abspath(out_dir)
    os.makedirs(out_dir, exist_ok=True)
    done = convert(files, out_dir)
    print(f"\n完成 {len(done)} 个 → {out_dir}")


if __name__ == "__main__":
    main()
