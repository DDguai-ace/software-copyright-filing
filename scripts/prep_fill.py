# -*- coding: utf-8 -*-
"""把「功能和技术特点_填报版（按字段）.txt」+ 两份 PDF 汇总成 fill_plan.json。

用法：
    python prep_fill.py "./output/软著_XXX_V1.0"

产出：
    <项目>/01_申请表/fill_plan.json   —— 填表阶段只读这个，不回翻长文档
    终端打印字段清单 + 字数/页数体检结果（任何一项 FAIL 先修材料再开浏览器）
"""
import json
import os
import re
import sys

LABEL_RE = re.compile(r'^(?P<label>[^（【—]{2,40})（[≤]?(?P<num>\d+|[0-9]+–[0-9]+)(?P<extra>[^）]*)）\s*$')
KEYMAP = [
    ('开发目的', 'dev_purpose'),
    ('面向领域', 'domain'),
    ('软件的主要功能', 'main_func'),
    ('软件的技术特点', 'tech_feature'),
    ('开发的硬件环境', 'dev_hw'),
    ('运行的硬件环境', 'run_hw'),
    ('开发该软件的操作系统', 'dev_os'),
    ('软件开发环境', 'dev_tool'),
    ('该软件的运行平台', 'run_platform'),
    ('软件运行支撑环境', 'support_sw'),
]
LIMITS = {
    'dev_purpose': 50, 'domain': 50, 'main_func': (500, 1300),
    'tech_feature': 100, 'dev_hw': 50, 'run_hw': 50,
    'dev_os': 50, 'dev_tool': 50, 'run_platform': 50, 'support_sw': 50,
}
# 「软件的技术特点」类别标签组（必填*，2026-11-12 实测：漏选也能存草稿，属带病数据）
TECH_TAGS = ['APP', '游戏软件', '教育软件', '金融软件', '医疗软件', '地理信息软件',
             '云计算软件', '信息安全软件', '大数据软件', '人工智能软件', 'VR软件',
             '5G软件', '小程序', '物联网软件', '智慧城市软件']


def parse_fields(txt):
    fields, cur, buf = {}, None, []
    for raw in txt.splitlines():
        line = raw.strip()
        if not line:
            continue
        if line.startswith('【') or line.startswith('——'):
            continue
        m = LABEL_RE.match(line)
        if m:
            if cur:
                fields[cur] = ' '.join(buf).strip()
            label = m.group('label').strip()
            cur, buf = None, []
            for prefix, key in KEYMAP:
                if label.startswith(prefix):
                    cur = key
                    break
            continue
        if line.startswith('技术特点标签'):
            # 技术特点标签：医疗软件、物联网软件（必填*，先于 LABEL_RE 处理，防止被当成上一字段正文）
            val = line.split('：', 1)[1] if '：' in line else line.split(':', 1)[-1]
            # 去掉行尾括号说明（如「（必填*，从 15 个合法选项选…）」），否则整段会被当成标签值
            val = re.sub(r'[（(][^）)]*[）)]\s*$', '', val).strip()
            tags = [x.strip() for x in re.split(r'[、,，\s]+', val) if x.strip()]
            if tags:
                fields['tech_tags'] = tags
            if cur:  # 先把上一字段的缓冲落库再清，否则 tech_feature 会被吞掉
                fields[cur] = ' '.join(buf).strip()
            cur, buf = None, []
            continue
        if line.startswith('编程语言'):
            mt = re.search(r'编程语言[:：]\s*([^\s　]+)', line)
            ms = re.search(r'源程序量[:：]\s*(\d+)', line)
            if mt:
                fields['lang'] = mt.group(1)
            if ms:
                fields['source_lines'] = int(ms.group(1))
            continue
        if cur:
            buf.append(line)
    if cur:
        fields[cur] = ' '.join(buf).strip()
    return fields


def pdf_pages(path):
    try:
        import pymupdf
        with pymupdf.open(path) as d:
            return d.page_count
    except Exception as exc:  # noqa: BLE001
        return f'ERR:{exc}'


def main():
    root = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else '.')
    form_txt = os.path.join(root, '01_申请表', '功能和技术特点_填报版（按字段）.txt')
    if not os.path.exists(form_txt):
        sys.exit(f'✗ 缺少填报版字段文件：{form_txt}\n  先去写材料（scripts/gen_forms.py 会生成）')

    fields = parse_fields(open(form_txt, encoding='utf-8').read())

    checks = {}
    for key, limit in LIMITS.items():
        val = fields.get(key, '')
        n = len(val)
        if isinstance(limit, tuple):
            ok = limit[0] <= n <= limit[1]
            checks[key] = {'chars': n, 'limit': f'{limit[0]}-{limit[1]}', 'ok': ok}
        else:
            checks[key] = {'chars': n, 'limit': f'<={limit}', 'ok': n <= limit}

    tags = fields.get('tech_tags') or []
    unknown = [x for x in tags if x not in TECH_TAGS]
    checks['tech_tags'] = {
        'tags': tags, 'ok': bool(tags) and not unknown,
        'msg': '' if tags else '缺「技术特点标签」行（必填*）；合法值：' + '/'.join(TECH_TAGS),
    }
    if unknown:
        checks['tech_tags']['msg'] = f'不在 15 个合法选项里：{unknown}'

    src_pdf = os.path.join(root, '02_源程序', '源程序鉴别材料_V1.0.pdf')
    doc_pdf = os.path.join(root, '03_说明书', '软件设计说明书_V1.0.pdf')
    files = {}
    for key, p in (('source_pdf', src_pdf), ('doc_pdf', doc_pdf)):
        files[key] = {'path': p if os.path.exists(p) else None,
                      'pages': pdf_pages(p) if os.path.exists(p) else None}
        if not os.path.exists(p):
            checks[key] = {'ok': False, 'msg': '缺 PDF，先跑 docx2pdf.py'}

    sources = [f for f in files.values() if f['path']]
    checks['files_present'] = {'ok': len(sources) == 2, 'count': len(sources)}

    plan = {
        'project_dir': root,
        'software_name': os.path.basename(root).replace('软著_', '').rsplit('_V', 1)[0],
        'version': 'V1.0',
        'fields': fields,
        'checks': checks,
        'files': files,
    }
    out = os.path.join(root, '01_申请表', 'fill_plan.json')
    with open(out, 'w', encoding='utf-8') as fh:
        json.dump(plan, fh, ensure_ascii=False, indent=2)

    bad = 0
    print(f'软件全称：{plan["software_name"]}  {plan["version"]}')
    print(f'{"字段":<12}{"字数":>5}  {"上限":<10}结果')
    for key, c in checks.items():
        if key in ('source_pdf', 'doc_pdf', 'files_present', 'tech_tags'):
            continue
        flag = 'OK' if c['ok'] else 'FAIL'
        bad += 0 if c['ok'] else 1
        print(f'{key:<12}{c["chars"]:>5}  {c["limit"]:<10}{flag}')
    print(f'\n源程序 PDF：{files["source_pdf"]["pages"]} 页')
    print(f'说明书 PDF：{files["doc_pdf"]["pages"]} 页')
    print(f'编程语言：{fields.get("lang", "?")}　源程序量：{fields.get("source_lines", "?")} 行')
    tc = checks['tech_tags']
    print(f'技术特点标签：{"、".join(tc["tags"]) or "（缺，必填*）"}　{"OK" if tc["ok"] else "FAIL " + tc.get("msg", "")}')
    print(f'\n→ {out}')
    if bad or not checks['files_present']['ok'] or not checks['tech_tags']['ok']:
        reasons = []
        if bad:
            reasons.append(f'{bad} 项字数不达标')
        if not checks['files_present']['ok']:
            reasons.append('缺 PDF')
        if not checks['tech_tags']['ok']:
            reasons.append('技术特点标签不合规')
        print(f'\n✗ 先修材料再开浏览器：' + '、'.join(reasons))
        sys.exit(1)
    print('\n✓ 字段与文件全部就绪，可以直接开始填表')


if __name__ == '__main__':
    main()
