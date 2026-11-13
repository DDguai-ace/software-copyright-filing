# -*- coding: utf-8 -*-
"""软著一条龙流水线：查阅文档 → 读源码 → 建材料 → 检查 → 填表 → 存草稿 → 复验 → 退出

用法：
    python run_project.py <项目目录> [--source 源代码.docx] [--doc 技术设计文档.docx]
                          [--software 软件全称] [--date YYYY-MM-DD] [--category 嵌入式软件]
                          [--no-fill] [--no-verify] [--cdp 9222]

五个阶段（任何一环 FAIL 即停，错误信息直接给出修法）：
  1 extract  把源代码.docx / 技术设计文档.docx 一次性转成纯文本 dump + 摘要
             （01_申请表/源码摘要.txt、技术文档摘要.txt）——写功能描述/说明书时
             只读这两个文件，不再反复解析 docx，提速「查阅文档、阅读源码」
  2 build    源程序重排（build_source.py）→ 批量转 PDF（docx2pdf.py）→ 页数行数校验
             （check_pages.py：源程序≥50 行/页、说明书≥30 行/页）
  3 prep     prep_fill.py 生成 fill_plan.json（含技术特点标签校验）
  4 fill     fill_r11.py 一条命令填完五步并存草稿（未登录会停在登录握手处）
  5 verify   重开草稿硬核验：两份 PDF 的 fileList 已挂进 <流水号>/ 正式目录、
             技术特点标签已勾选 → 合并写进 05_提交记录/填表结果.json，然后退出

说明：阶段 1 的「写功能描述/说明书文案」本身是创作活，仍由助手完成——
本脚本把机械环节全部自动化并落盘中间产物，助手只读摘要即可动笔。
"""
import argparse
import json
import os
import re
import subprocess
import sys
import time

SCRIPTS = os.path.dirname(os.path.abspath(__file__))
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)

PY = sys.executable
NAME_DIR_RE = re.compile(r'^软著_(.+)_V\d')


def sh(script, *args, **kw):
    """调同目录下的兄弟脚本，失败即抛。"""
    cmd = [PY, os.path.join(SCRIPTS, script)] + [str(a) for a in args]
    print(f'  $ {os.path.basename(script)} {" ".join(str(a) for a in args)}')
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace', **kw)
    out = (r.stdout or '') + (r.stderr or '')
    if r.returncode != 0:
        sys.exit(f'✗ {script} 失败（exit {r.returncode}）：\n{out[-2000:]}')
    return out


def docx_lines(path):
    """读 docx 全部行（含 <w:br/> 拆分），与 build_source 口径一致。"""
    from docx import Document
    d = Document(path)
    return [l.rstrip() for p in d.paragraphs for l in p.text.split('\n')]


def stage_extract(root, src_docx, tech_docx):
    out_dir = os.path.join(root, '01_申请表')
    os.makedirs(out_dir, exist_ok=True)
    made = []
    if src_docx and os.path.exists(src_docx):
        lines = docx_lines(src_docx)
        nonblank = [l for l in lines if l.strip()]
        # 函数/符号名（C 系口径），保序去重
        seen, funcs = set(), []
        for l in nonblank:
            m = re.match(r'^[ \t]*(?:static\s+|const\s+|extern\s+|unsigned\s+|signed\s+)*'
                         r'[A-Za-z_][\w\s\*]*?\**\s*([A-Za-z_]\w+)\s*\(', l)
            if m and m.group(1) not in ('if', 'for', 'while', 'switch', 'return', 'sizeof'):
                n = m.group(1)
                if n not in seen:
                    seen.add(n)
                    funcs.append(n)
        dump = os.path.join(out_dir, '源码全文.txt')
        with open(dump, 'w', encoding='utf-8') as fh:
            fh.write('\n'.join(nonblank))
        summary = os.path.join(out_dir, '源码摘要.txt')
        with open(summary, 'w', encoding='utf-8') as fh:
            fh.write(f'源码文件：{src_docx}\n总行数（去空行）：{len(nonblank)}\n'
                     f'函数/符号数（去重）：{len(funcs)}\n\n函数与符号清单（按出现顺序）：\n')
            fh.write('\n'.join(funcs[:200]))
            fh.write(f'\n\n全文已落盘：{dump}（写说明书/功能描述直接读它）\n\n开头 30 行样例：\n')
            fh.write('\n'.join(nonblank[:30]))
        made.append('源码摘要.txt + 源码全文.txt')
        print(f'  源码：{len(nonblank)} 行（去空行），{len(funcs)} 个函数/符号')
    if tech_docx and os.path.exists(tech_docx):
        from docx import Document
        d = Document(tech_docx)
        parts, tree = [], []
        for p in d.paragraphs:
            t = p.text.strip()
            if not t:
                continue
            style = (p.style.name or '') if p.style is not None else ''
            m = re.match(r'Heading (\d)', style) or re.match(r'标题 (\d)', style)
            if m:
                tree.append('  ' * (int(m.group(1)) - 1) + t)
            parts.append(t)
        dump = os.path.join(out_dir, '技术文档摘要.txt')
        with open(dump, 'w', encoding='utf-8') as fh:
            fh.write(f'技术文档：{tech_docx}\n段落数：{len(parts)}\n\n标题树：\n')
            fh.write('\n'.join(tree) if tree else '（无标题样式）')
            fh.write('\n\n全文：\n')
            fh.write('\n'.join(parts))
        made.append('技术文档摘要.txt')
        print(f'  技术文档：{len(parts)} 段，{len(tree)} 个标题')
    if made:
        print(f'  ✓ 摘要落盘：{"、".join(made)}（写材料只读这些，不回翻 docx）')
    else:
        print('  （未提供 --source/--doc，跳过提取）')


def stage_build(root, src_docx):
    src_pdf = os.path.join(root, '02_源程序', '源程序鉴别材料_V1.0.pdf')
    doc_pdf = os.path.join(root, '03_说明书', '软件设计说明书_V1.0.pdf')
    name = plan_name(root)
    if not os.path.exists(src_pdf):
        src_doc = os.path.join(root, '02_源程序', '源程序鉴别材料_V1.0.docx')
        if not os.path.exists(src_doc):
            if not (src_docx and os.path.exists(src_docx)):
                sys.exit('✗ 缺源程序材料：没有现成 docx/pdf，也没给 --source，无法重排')
            out_doc = os.path.join(root, '02_源程序')
            os.makedirs(out_doc, exist_ok=True)
            sh('build_source.py', src_docx,
               os.path.join(out_doc, '源程序鉴别材料_V1.0.docx'), name, 'V1.0')
    # 凡是 docx 缺 pdf 就批量转
    have = []
    for sub in ('01_申请表', '02_源程序', '03_说明书'):
        p = os.path.join(root, sub)
        if not os.path.isdir(p):
            continue
        for f in os.listdir(p):
            if f.endswith('.docx') and not f.startswith('~$'):
                pdf = os.path.join(p, f[:-5] + '.pdf')
                if not os.path.exists(pdf):
                    have.append(os.path.join(p, f))
    for f in have:
        sh('docx2pdf.py', f, os.path.dirname(f))
    if not os.path.exists(src_pdf) or not os.path.exists(doc_pdf):
        sys.exit(f'✗ 缺鉴别材料 PDF：{src_pdf} / {doc_pdf}')
    print('  校验源程序（≥50 行/页）…')
    sh('check_pages.py', src_pdf, '--min-lines', '50')
    print('  校验说明书（≥30 行/页）…')
    sh('check_pages.py', doc_pdf, '--min-lines', '30')


def plan_name(root):
    base = os.path.basename(root.rstrip('\\/'))
    m = NAME_DIR_RE.match(base)
    return m.group(1) if m else base


def stage_prep(root):
    sh('prep_fill.py', root)


def stage_fill(root, date, category):
    args = [PY, os.path.join(SCRIPTS, 'fill_r11.py'), root]
    if date:
        args += ['--date', date]
    if category:
        args += ['--category', category]
    print('  $ fill_r11.py（五步填完 → 存草稿 → 取流水号）')
    r = subprocess.run(args, capture_output=True, text=True, encoding='utf-8', errors='replace')
    out = (r.stdout or '') + (r.stderr or '')
    print('\n'.join('  ' + x for x in out.strip().splitlines()[-12:]))
    if r.returncode != 0:
        sys.exit('✗ fill_r11.py 失败。若是「未登录」，先按 §5.0 请用户登录后重跑本命令；'
                 '其余错误按输出信息修。')
    return out


def stage_verify(root, cdp):
    """重开草稿硬核验：材料 fileList + 技术特点标签。复用 fill_r11 的浏览器助手。"""
    import fill_r11 as fr
    rec = os.path.join(root, '05_提交记录', '填表结果.json')
    serial = None
    if os.path.exists(rec):
        data0 = json.load(open(rec, encoding='utf-8'))
        serial = data0.get('serial') or data0.get('flow_number')  # 兼容手写的历史结果文件
    if not serial:
        print('  （无流水号——fill 未跑或 --no-fill，跳过复验）')
        return None
    fr.CDP = str(cdp)
    fr.AB = fr.find_ab()
    fr.run('open', f'https://register.ccopyright.com.cn/r11.html#/identity?flowNumber={serial}')
    time.sleep(4)
    fr.ev("location.hash='#/features?flowNumber=%s'" % serial)
    time.sleep(3)
    fl = fr.ev("(()=>{const o=[];document.querySelectorAll('.upLoadBox').forEach((b,i)=>{"
               "const vm=b.querySelector('.hdUpload').__vue__;"
               "const l=(vm&&vm.$props&&vm.$props.fileList)||[];"
               "o.push(l[0]?{name:l[0].name,url:String(l[0].url)}:'EMPTY')});return o})()")
    ok_files = (len(fl) >= 2 and isinstance(fl[0], dict) and isinstance(fl[1], dict)
                and serial in (fl[0].get('url') or '') and serial in (fl[1].get('url') or ''))
    checked = fr.ev("(()=>{return [...document.querySelectorAll('label.hd-checkbox-button')]"
                    ".filter(l=>{const i=l.querySelector('input');return i&&i.checked})"
                    ".map(l=>l.innerText.trim())})()") or []
    plan = json.load(open(os.path.join(root, '01_申请表', 'fill_plan.json'), encoding='utf-8'))
    want_tags = plan['fields'].get('tech_tags') or []
    missing_tags = [t for t in want_tags if t not in checked]
    res = {
        'serial': serial,
        'files_ok': ok_files,
        'files': fl[:2],
        'checked_labels': checked,
        'tags_ok': not missing_tags,
        'missing_tags': missing_tags,
        'verified_at': time.strftime('%Y-%m-%d %H:%M:%S'),
    }
    data = json.load(open(rec, encoding='utf-8')) if os.path.exists(rec) else {}
    data['verify'] = res
    json.dump(data, open(rec, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
    print(f'  材料绑定：{"✓ 两份 PDF 已挂进正式目录" if ok_files else "✗ " + json.dumps(fl[:2], ensure_ascii=False)}')
    print(f'  技术特点标签：{"✓ " + "、".join(want_tags) if res["tags_ok"] else "✗ 缺 " + "、".join(missing_tags)}')
    if not (ok_files and res['tags_ok']):
        sys.exit('✗ 复验不通过：材料或标签没存上，回功能页重传/重勾后重新保存草稿')
    print(f'  ✓ 复验通过，结果写入 05_提交记录/填表结果.json')
    return res


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('project')
    ap.add_argument('--source', default=None, help='源代码.docx 路径（阶段1提取 + 缺料时重排）')
    ap.add_argument('--doc', default=None, help='技术设计文档.docx 路径（阶段1提取）')
    ap.add_argument('--date', default=None)
    ap.add_argument('--category', default=None)
    ap.add_argument('--no-fill', action='store_true', help='只跑到阶段3（材料+体检），不开浏览器')
    ap.add_argument('--no-verify', action='store_true', help='填完不重开草稿复验')
    ap.add_argument('--cdp', default='9222')
    args = ap.parse_args()
    root = os.path.abspath(args.project)
    if not os.path.isdir(root):
        sys.exit(f'✗ 项目目录不存在：{root}')

    t0 = time.time()
    print(f'【1/5】提取源码与技术文档 →')
    stage_extract(root, args.source, args.doc)
    print(f'【2/5】建材料 + 转PDF + 校验 →')
    stage_build(root, args.source)
    print(f'【3/5】填表体检 →')
    stage_prep(root)
    if args.no_fill:
        print('（--no-fill：到体检为止）')
        return
    print(f'【4/5】在线填表 + 存草稿 →')
    stage_fill(root, args.date, args.category)
    if not args.no_verify:
        print(f'【5/5】草稿复验 →')
        stage_verify(root, args.cdp)
    print(f'\n✓ 全链完成，耗时 {round(time.time()-t0,1)}s。剩余人工动作：确认填报、签章页手抄盖章。')


if __name__ == '__main__':
    main()
