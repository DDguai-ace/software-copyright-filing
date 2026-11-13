# -*- coding: utf-8 -*-
"""R11 表单一条命令跑完：identity → application → development → features → confirm → 存草稿。

前置：
    1. 已起受控 Edge（--remote-debugging-port=9222）；**窗口被关了也不用管，脚本会自动重起**
    2. 用户已登录（页头出现「用户中心」）——重启窗口后登录态会失效，需重登
    3. 已跑过 prep_fill.py，存在 01_申请表/fill_plan.json

用法：
    python fill_r11.py "./output/软著_XXX_V1.0"                   # 全程
    python fill_r11.py "<项目>" --check                          # 只体检，不动表单
    python fill_r11.py "<项目>" --dry-run                        # 打印步骤，不动浏览器
    python fill_r11.py "<项目>" --date 2026-09-13 --category 嵌入式软件
    python fill_r11.py "<项目>" --stop-after 4                    # 验证到第四步，不存草稿
    python fill_r11.py "<项目>" --no-launch                       # 端口不通时不要自动起 Edge

注意：**只支持一次跑完，不支持断点续跑**（页面的 flow 状态不落在 URL 里，中途重进会丢上下文）。

产出：<项目>/05_提交记录/填表结果.json + 终端报告（含流水号）
"""
import argparse
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request

CDP = '9222'
AB = None
NO_SAVE = False

EDGE = r'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe'
PROFILE = os.path.join(tempfile.gettempdir(), 'edge-cdp-filing')


# ---------------------------------------------------------------- 浏览器生命周期
def port_alive(timeout=3):
    """只认端口，不认进程数——用户自己的 Edge 也在跑，进程在不能证明 9222 活着。"""
    try:
        with urllib.request.urlopen(f'http://127.0.0.1:{CDP}/json/version', timeout=timeout) as r:
            return r.status == 200
    except Exception:
        return False


def ensure_browser(launch=True):
    """端口不通说明受控窗口被关了（用户关窗或后台进程被回收），不是工具故障。

    这里只是**兜底**：脚本自己起的窗口可能在本进程退出后被运行环境回收，
    因此「让用户登录」这一步必须由 agent 用后台任务起窗口（见 SKILL.md §5.0.1），
    脚本的自动重启只保证「本进程内能连上」。
    另：登录态实测**不跨窗口保留**（凭据在会话级存储 + 短时效），重启后基本要重登。
    """
    if port_alive():
        return 'existing'
    if not launch:
        raise SystemExit(
            '✗ 9222 端口不通（受控窗口已关闭）。去掉 --no-launch 让我自动起窗口，'
            '或手跑 SKILL.md §5.0.1 第 1 步。')
    if not os.path.exists(EDGE):
        raise SystemExit(f'✗ 找不到 Edge：{EDGE}')
    # 必须 DETACHED：否则窗口是脚本的子进程，脚本一退出就被回收，用户根本来不及登录
    flags = 0
    if os.name == 'nt':
        flags = 0x00000008 | 0x00000200        # DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP
    subprocess.Popen([EDGE, f'--remote-debugging-port={CDP}', f'--user-data-dir={PROFILE}',
                      '--window-size=1560,1000', '--no-first-run', '--no-default-browser-check',
                      'https://www.ccopyright.com/'],
                     creationflags=flags, close_fds=True,
                     stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    for i in range(30):
        time.sleep(1)
        if port_alive():
            return f'relaunched({i + 1}s)'
    raise SystemExit('✗ Edge 已启动但 9222 一直不通：检查端口占用或安全软件拦截')


# ---------------------------------------------------------------- 基础设施
def find_ab():
    """返回启动 agent-browser 的命令前缀（列表）。

    优先用原生 exe / node+js：agent-browser.cmd 会被 cmd.exe 解释，
    JS 里的 `|` `>` `&` 会被当成重定向或管道，踩过一次坑就不再走 .cmd。
    """
    base = os.path.join(os.path.dirname(os.path.abspath(__file__)))
    candidates = []
    for root in (
        os.path.join(os.environ.get('APPDATA', ''), 'npm/node_modules/agent-browser'),
    ):
        candidates += [
            [os.path.join(root, 'bin', 'agent-browser-win32-x64.exe')],
            ['node', os.path.join(root, 'bin', 'agent-browser.js')],
        ]
    for c in candidates:
        if c[0].startswith('node') or os.path.exists(c[0]):
            return c
    p = shutil.which('agent-browser')
    if p:
        return [p]
    sys.exit('✗ 找不到 agent-browser，先执行：npm i -g agent-browser')


def run(*args, timeout=90):
    cmd = [*AB, *map(str, args), '--cdp', CDP]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=timeout)
    out = (r.stdout or '').strip()
    if r.returncode != 0:
        raise RuntimeError(f'命令失败: {" ".join(map(str, args))}\n{(r.stderr or out)[:400]}')
    return out


def ev(js, timeout=60):
    """执行 JS，尽量把返回值解析成 Python 对象。"""
    out = run('eval', js, timeout=timeout)
    out = re.sub(r'^\[agent-browser\][^\n]*\n?', '', out).strip()
    for cand in (out,):
        try:
            return json.loads(cand)
        except Exception:  # noqa: BLE001
            pass
    if out.startswith('"') and out.endswith('"'):
        try:
            return json.loads(out)
        except Exception:  # noqa: BLE001
            return out.strip('"')
    return out


def url():
    out = run('get', 'url')
    for line in reversed(out.splitlines()):
        if line.startswith('http'):
            return line.strip()
    return out.strip()


def wait_for(desc, fn, timeout=25, interval=0.7):
    end = time.time() + timeout
    last = None
    while time.time() < end:
        try:
            last = fn()
            if last:
                return last
        except Exception as exc:  # noqa: BLE001
            last = f'err: {exc}'
        time.sleep(interval)
    raise RuntimeError(f'等待超时（{desc}），最后状态：{last}')


def wait_url(part, timeout=30):
    return wait_for(f'URL 含 {part}', lambda: part in url() and url(), timeout)


def wait_step(desc, expr, timeout=25):
    """按**页面内容**判断是否进入下一步。

    2026-11-12 实测：这版登记系统的 URL 哈希不随向导步骤变化，
    一直停在 r11.html#/identity，`wait_url('#/application')` 必然超时。
    """
    return wait_for(desc, lambda: ev(expr) is True, timeout)


def has_button(text):
    return ("!![...document.querySelectorAll('button')].find(b=>b.innerText.trim()==='%s'" 
            "&&b.offsetParent!==null)" % text)


def body_text(limit=4000):
    return ev(f'document.body.innerText.replace(/\\s+/g," ").slice(0,{limit})')


def clean_ids():
    """清掉上一轮留下的 __id，避免 Vue 复用节点后 Playwright 点到旧元素。"""
    try:
        ev("[...document.querySelectorAll('[id^=\"__\"]')].forEach(e=>e.removeAttribute('id'))")
    except Exception:  # noqa: BLE001
        pass


def click_text(tag, text, ident, exact=True):
    """给文字匹配的元素赋 id 再 click（@ref 跨命令失效，只能这么点）。"""
    cond = f"x.innerText.trim()==='{text}'" if exact else f"x.innerText.includes('{text}')"
    ok = ev(
        "(()=>{const b=[...document.querySelectorAll('%s')].filter(x=>%s&&x.offsetParent!==null);"
        "if(!b.length)return 0;b[b.length-1].id='%s';return 1})()" % (tag, cond, ident)
    )
    if not ok:
        raise RuntimeError(f'页面上找不到 {tag}[{text}]')
    run('click', '#' + ident)


# ---------------------------------------------------------------- 各步骤
def check_login():
    t = body_text(400)
    if '用户中心' in t and '退出登录' not in t[:60]:
        return True
    if '用户中心' in t:
        return True
    raise SystemExit('✗ 未登录：页头没有「用户中心」。请让用户在受控窗口里登录后再跑（或先用对话框确认「已登录」）')


def step1_identity():
    clean_ids()
    if '#/identity' not in url():
        run('open', 'https://register.ccopyright.com.cn/r11.html#/identity')
        time.sleep(2.5)
    if '#/identity' not in url():
        raise RuntimeError(f'没进到身份选择页，当前 {url()}（若被弹回 login.html 说明登录态失效）')
    ok = ev(
        "(()=>{const c=[...document.querySelectorAll('div')].filter(e=>e.innerText&&"
        "e.innerText.trim().startsWith('我是申请人')&&e.innerText.length<40);"
        "if(!c.length)return 0;const t=c[c.length-1];t.id='__identity';return 1})()"
    )
    if not ok:
        # 这版系统改 hash 不会重渲染（还停在原来的步骤），必须真刷新一次
        run('reload')
        time.sleep(3.5)
        ok = ev(
            "(()=>{const c=[...document.querySelectorAll('div')].filter(e=>e.innerText&&"
            "e.innerText.trim().startsWith('我是申请人')&&e.innerText.length<40);"
            "if(!c.length)return 0;const t=c[c.length-1];t.id='__identity';return 1})()"
        )
    if not ok:
        raise RuntimeError('找不到「我是申请人」卡片')
    run('click', '#__identity')
    wait_step('第二步表单就绪',
              "!!document.querySelector(\"input[placeholder='请输入软件全称']\")", 25)
    return '选「我是申请人」'


def step2_application(name, version):
    clean_ids()
    run('fill', "input[placeholder='请输入软件全称']", name)
    run('fill', "input[placeholder='请输入版本号']", version)
    got = ev("(()=>{const i=[...document.querySelectorAll('input')];"
             "const n=i.find(x=>x.placeholder.includes('软件全称'));"
             "const v=i.find(x=>x.placeholder.includes('版本号'));"
             "const s=i.find(x=>x.placeholder.includes('简称'));"
             "return {n:n.value,v:v.value,s:s.value,orig:i[0].checked,all:i[5]?i[5].checked:null}})()")
    if got.get('n') != name or got.get('v') != version:
        raise RuntimeError(f'第二步回填不一致：{got}')
    if got.get('s'):
        raise RuntimeError('软件简称被填了值，必须留空')
    click_text('button', '下一步', '__next2')
    wait_step('第三步表单就绪',
              "!![...document.querySelectorAll('h3')].find(e=>e.innerText.trim()==='软件分类')", 25)
    return f'全称={name} 版本={version} 简称留空'


def step3_development(category, done_date):
    clean_ids()
    # 软件分类：自定义下拉。必须先点开、等一 tick 让 .hd-option 渲染，再派发三个鼠标事件
    picked = None
    for _ in range(3):
        ev("(()=>{const it=[...document.querySelectorAll('.fillin_item')].find(e=>e.querySelector('h3')"
           "&&e.querySelector('h3').innerText.trim()==='软件分类');if(!it)return 0;"
           "it.querySelector('.hd-select .box').click();return 1})()")
        time.sleep(0.45)
        picked = ev(
            "(()=>{const it=[...document.querySelectorAll('.fillin_item')].find(e=>e.querySelector('h3')"
            "&&e.querySelector('h3').innerText.trim()==='软件分类');if(!it)return 'no-item';"
            "const o=[...it.querySelectorAll('.hd-option')].find(x=>x.innerText.trim()==='%s');if(!o)return 'no-opt';"
            "['mousedown','mouseup','click'].forEach(t=>o.dispatchEvent(new MouseEvent(t,{bubbles:true})));"
            "return it.querySelector('.hd-select .box').innerText.trim()})()" % category)
        if picked == category:
            break
        time.sleep(0.6)
    if picked != category:
        raise RuntimeError(f'软件分类没选上：{category}（回显 {picked}）')

    # 开发完成日期：点开面板 → 翻到目标年月 → 点日号
    y, m, d = [int(x) for x in done_date.split('-')]

    def head_txt():
        return ev("document.querySelector('.datepicker-main').querySelector('.datepicker-header')"
                  ".innerText.replace(/\\s+/g,' ')")

    ev("document.querySelector('.datePicker input').click()")
    wait_for('日期面板打开', lambda: '年' in str(head_txt()), timeout=8)
    head = head_txt()
    cy = int(re.search(r'(\d{4})年', head).group(1))
    cm = int(re.search(r'(\d{1,2})月', head).group(1))
    diff = (y - cy) * 12 + (m - cm)
    btn = 'span.datepicke-btn-prve:nth-of-type(3)' if diff < 0 else 'span.datepicke-btn-next:nth-of-type(4)'
    for _ in range(abs(diff)):
        before = head_txt()
        for attempt in range(4):
            run('click', '.datepicker-header ' + btn)
            time.sleep(0.5)
            if head_txt() != before:
                break
        else:
            raise RuntimeError(f'日期面板翻月没反应，卡在 {head_txt()}')
    head2 = head_txt()
    if f'{y}年' not in head2 or f' {m}月' not in head2:
        raise RuntimeError(f'日期面板没翻到 {y}年{m}月，当前 {head2}')
    # id 必须每次唯一：面板切月时 Vue 会复用 td 节点，旧 id 残留会点到上一个月的格子
    day_id = f'__day{int(time.time() * 1000) % 1000000}'
    ok = ev("(()=>{const p=document.querySelector('.datepicker-main');"
            "const t=[...p.querySelectorAll('td')].find(c=>c.innerText.trim()==='%d'&&c.className.trim()==='');"
            "if(!t)return 0;t.id='%s';return 1})()" % (d, day_id))
    if not ok:
        raise RuntimeError(f'找不到 {d} 日这一格（可能不在当月）')
    # 必须用 JS 派发事件，不能用 Playwright 坐标点击：
    # 面板常渲染在视口下方，Playwright 会先 scrollIntoView，内部滚动容器一移，
    # 点到的是下一行（选 13 落到 20）。踩过两次，别再改回去。
    ev("(()=>{const p=document.querySelector('.datepicker-main');"
       "const t=[...p.querySelectorAll('td')].find(c=>c.innerText.trim()==='%d'&&c.className.trim()==='');"
       "const sp=t.querySelector('span')||t;['mousedown','mouseup','click']"
       ".forEach(n=>{t.dispatchEvent(new MouseEvent(n,{bubbles:true}));"
       "sp.dispatchEvent(new MouseEvent(n,{bubbles:true}))});return 1})()" % d)
    time.sleep(0.6)
    val = ev("document.querySelector('.datePicker input').value")
    if val != done_date:
        raise RuntimeError(f'日期回填是 {val}，应为 {done_date}')

    # 默认项兜底：原创 / 单独开发 / 未发表 / 全体权利人
    st = ev("(()=>{const r=[...document.querySelectorAll('input[type=radio]')];"
            "const pick=v=>{const x=r.find(t=>t.value===v);return x?x.checked:null};"
            "return {orig:pick('ORIGINAL'),single:pick('SINGLE'),unpub:pick('0')}})()")
    if not (st.get('orig') and st.get('single')):
        raise RuntimeError(f'默认选项异常（原创/单独开发应为已选）：{st}')

    click_text('button', '下一步', '__next3')
    wait_step('第四步表单就绪', "document.querySelectorAll('textarea').length===11", 30)
    return f'分类={category} 完成日期={done_date} 原创/单独开发/未发表'


def filelist(idx):
    """读上传组件**真实绑定**的服务器文件（Vue props.fileList）——唯一可信判据。

    DOM 上的文件名和「请先完成上传」都不可信：文件压根没上传时，文件名照样显示，
    确认页也照样列出文件名（2026-11-12 踩过：草稿存了，材料是空的）。
    """
    return ev(
        "(()=>{const b=document.querySelectorAll('.upLoadBox')[%d];if(!b)return [];"
        "const el=b.querySelector('.hdUpload');const vm=el&&el.__vue__;"
        "const fl=(vm&&vm.$props&&vm.$props.fileList)||[];"
        "return fl.map(f=>({name:f.name||f.fileName||'',url:String(f.url||'')}))})()" % idx) or []


def filesobj(idx):
    """读上传组件的原始响应（**全新流程阶段**唯一可信判据）。

    fileList 要等保存草稿、服务端把 temp 文件搬进 soft/attachment/…/<流水号>/ 后才回绑；
    全新流程没有流水号，fileList 恒为空，但文件其实已经传上 OSS——
    此时只有 filesObj[0].response.data = {filePath: temp/…, pageNum: N} 能证明上传成功
    （2026-11-12 两次踩到：拿 fileList 当判据 → 45 秒必超时 → 误判假失败）。
    """
    return ev(
        "(()=>{const b=document.querySelectorAll('.upLoadBox')[%d];if(!b)return null;"
        "const el=b.querySelector('.hdUpload');const vm=el&&el.__vue__;"
        "const fo=(vm&&vm.$data&&vm.$data.filesObj)||[];"
        "const f0=fo[0];if(!f0||!f0.response)return null;"
        "const d=f0.response.data||{};"
        "return {code:f0.response.returnCode||'',filePath:d.filePath||'',"
        "pageNum:String(d.pageNum||''),name:f0.name||''}})()" % idx)


def upload_one(idx, path):
    """上传一份鉴别材料。两个动作缺一就静默失败：

    1. **先删残留卡片**——同名文件会让组件的 change 被当成「没变化」直接忽略
    2. **upload 后手工派发 input/change**——CDP 的 setInputFiles 只塞文件、不触发 Vue 监听，
       上传请求根本发不出去（实测：不派发时零请求，而页面照样显示文件名，极易误判成功）

    成功判据分两段（见 filesobj / filelist 注释）：
    新流程看 filesObj.response.data.filePath（temp 路径）；已保存流程看 fileList 的正式路径。
    """
    if not os.path.exists(path):
        raise RuntimeError(f'缺文件：{path}')
    base = os.path.basename(path)
    ev("(()=>{const b=document.querySelectorAll('.upLoadBox')[%d];"
       "const t=b&&b.querySelector('.removeBtn');"
       "if(t)t.dispatchEvent(new MouseEvent('click',{bubbles:true}));return 1})()" % idx)
    time.sleep(1.0)
    ev("(()=>{document.querySelectorAll('input[type=file]').forEach((f,i)=>f.id='__up'+i);return 1})()")
    run('upload', f'#__up{idx}', path, timeout=180)
    time.sleep(1.0)
    # 按**下标**取当前 input，不按 id：upload 之后 Vue 可能重建节点，旧 id 会指向游离元素
    ev("(()=>{const f=document.querySelectorAll('input[type=file]')[%d];if(!f)return 0;"
       "f.dispatchEvent(new Event('change',{bubbles:true}));return 1})()" % idx)

    def bound():
        # 双判据：已保存流程看 fileList 正式绑定；全新流程看 filesObj 的 temp 响应
        if len(filelist(idx)) > 0:
            return True
        fo = filesobj(idx)
        return bool(fo and fo.get('code') == 'SUCCESS' and fo.get('filePath'))

    try:
        wait_for(f'{base} 上传确认', bound, timeout=45)
    except RuntimeError:
        # 一次重试：重新选同一个文件再派发一次（偶发组件重建导致首轮丢失）
        ev("(()=>{document.querySelectorAll('input[type=file]').forEach((f,i)=>f.id='__up'+i);return 1})()")
        run('upload', f'#__up{idx}', path, timeout=180)
        time.sleep(1.0)
        ev("(()=>{const f=document.querySelectorAll('input[type=file]')[%d];if(!f)return 0;"
           "f.dispatchEvent(new Event('change',{bubbles:true}));return 1})()" % idx)
        wait_for(f'{base} 上传确认（重试）', bound, timeout=45)

    files = filelist(idx)
    if files:
        if base not in (files[0].get('name') or ''):
            raise RuntimeError(f'框{idx} 绑定的是 {files[0].get("name")}，不是 {base}')
        if not files[0].get('url'):
            raise RuntimeError(f'框{idx} 有文件记录但没有服务器路径，等于没传')
        return files[0]['url']
    # 全新流程：fileList 为空属正常，temp 路径即成功，保存草稿时服务端搬进正式目录
    fo = filesobj(idx)
    if not (fo and fo.get('code') == 'SUCCESS' and fo.get('filePath')):
        raise RuntimeError(f'框{idx} 上传未确认成功：{fo}')
    return fo['filePath']


def step4_features(fields, src_pdf, doc_pdf, langs):
    clean_ids()
    # 11 个文本框固定顺序
    wait_for('第四步表单就绪', lambda: ev("document.querySelectorAll('textarea').length") == 11, timeout=20)
    order = ['dev_hw', 'run_hw', 'dev_os', 'dev_tool', 'run_platform',
             'support_sw', 'lang_other', 'dev_purpose', 'domain', 'main_func', 'tech_feature']
    ev("(()=>{document.querySelectorAll('textarea').forEach((t,i)=>t.id='__ta'+i);"
       "const a=[...document.querySelectorAll('input')].find(x=>x.maxLength===20);"
       "if(a)a.id='__amount';return 1})()")
    filled = {}
    for i, key in enumerate(order):
        val = fields.get(key) or ''
        if not val:
            # 计划里没有值的框必须显式清空：页面会残留脏值（实测 lang_other 自己冒出一个 "C"）
            ev("(()=>{const t=document.querySelector('#__ta%d');if(!t)return 0;t.value='';"
               "t.dispatchEvent(new Event('input',{bubbles:true}));return 1})()" % i)
            continue
        run('fill', f'#__ta{i}', val)
        filled[key] = len(val)

    # 编程语言列表默认折叠（只露 7 个，Python/Java 等在「展开更多」里）：先展开再找 label，
    # 否则隐藏项点击无效、勾不上还看不出原因（2026-11-12 实测）
    ev("(()=>{const e=[...document.querySelectorAll('span,div,a')]"
       ".find(x=>x.innerText.trim()==='展开更多'&&x.children.length===0);if(e)e.click();return 1})()")
    time.sleep(0.4)
    missing = []
    for lang in langs:
        ck = ("(()=>{const l=[...document.querySelectorAll('label')].find(x=>x.innerText.trim()==='%s');"
              "return l&&l.offsetParent!==null?!!l.querySelector('input').checked:'no-label'})()" % lang)
        state = ev(ck)
        if state == 'no-label':
            missing.append(lang)   # 复选框里没有 -> 落到「其他编程语言」文本框，不中断
            continue
        if not state:
            ev("(()=>{const l=[...document.querySelectorAll('label')].find(x=>x.innerText.trim()==='%s');"
               "l.click();return 1})()" % lang)
            # Vue 的 checked 是下一 tick 才更新，别立刻读
            wait_for(f'{lang} 勾选生效', lambda: ev(ck) is True, timeout=6)
    if missing:
        # 只有复选框里没有的语言才写进「其他」，别和勾选重复（按 textarea 下标写入，Vue 会重建节点，id 不可靠）
        txt = '、'.join(missing)[:120]
        ev("(()=>{const t=document.querySelectorAll('textarea')[6];if(!t)return 0;t.focus();t.value='%s';"
           "t.dispatchEvent(new Event('input',{bubbles:true}));t.dispatchEvent(new Event('change',{bubbles:true}));"
           "return t.value.length})()" % txt)
        time.sleep(0.5)

    # 「软件的技术特点」类别标签组（必填*，与编程语言同为隐藏 checkbox 的 label）。
    # 2026-11-12 实测：漏选不拦截保存，草稿会带病入库——必须显式勾选并复核。
    for tag in fields.get('tech_tags') or []:
        ck = ("(()=>{const l=[...document.querySelectorAll('label.hd-checkbox-button')]"
              ".find(x=>x.innerText.trim()==='%s');"
              "return l?!!l.querySelector('input').checked:'no-label'})()" % tag)
        state = ev(ck)
        if state == 'no-label':
            raise RuntimeError(f'技术特点标签「{tag}」不在 15 个选项里（APP/游戏软件/…/智慧城市软件）')
        if not state:
            ev("(()=>{const l=[...document.querySelectorAll('label.hd-checkbox-button')]"
               ".find(x=>x.innerText.trim()==='%s');l.click();return 1})()" % tag)
            wait_for(f'标签 {tag} 勾选生效', lambda: ev(ck) is True, timeout=6)

    run('fill', '#__amount', str(fields['source_lines']))
    time.sleep(0.8)
    chk = ev("(()=>{const ts=[...document.querySelectorAll('textarea')];"
             "return {lens:ts.map(t=>t.value.length),vals:ts.map(t=>t.value.slice(0,12)),"
             "amount:document.querySelector('#__amount').value}})()")
    want = [len(fields.get(k) or '') for k in order]
    if chk.get('lens') != want:
        bad = [(order[i], chk['lens'][i], want[i]) for i in range(len(want)) if chk['lens'][i] != want[i]]
        raise RuntimeError(f'文本框字数与计划不一致：{bad}\n实际值 {chk.get("vals")}')
    if chk.get('amount') != str(fields['source_lines']):
        raise RuntimeError(f'源程序量填的是 {chk.get("amount")}，应为 {fields["source_lines"]}')

    for idx, path in ((0, src_pdf), (1, doc_pdf)):
        upload_one(idx, path)

    click_text('button', '下一步', '__next4')
    wait_step('确认页就绪', "!![...document.querySelectorAll('button')].find(b=>b.innerText.trim()==='保存至草稿箱')", 30)
    return f'六项环境+目的+领域+功能{len(fields["main_func"])}字+特点{len(fields["tech_feature"])}字 / {langs} / {fields["source_lines"]}行 / 两份PDF'


def step5_confirm(expect):
    """逐项核对确认页，任一项对不上就抛错，不存草稿。"""
    clean_ids()
    wait_step('确认页就绪', "!![...document.querySelectorAll('button')].find(b=>b.innerText.trim()==='保存至草稿箱')", 20)
    t = ev("document.body.innerText.replace(/[ \\t]+/g,' ')")
    miss = [k for k, v in expect.items() if v not in t]
    if miss:
        raise RuntimeError('确认页缺少这些值：' + '；'.join(f'{k}={expect[k]}' for k in miss))
    if NO_SAVE:
        return '确认页核对通过（--no-save：停在确认页，未存草稿）'
    # 底部按钮必须先滚到视口中央再点：确认页很长，直接 click 会被遮罩吞掉且不报错
    # （2026-11-12 实测：同一按钮连点两次都无反应，scrollIntoView 后一次成功）
    ev("(()=>{const b=[...document.querySelectorAll('button')].find(x=>x.innerText.trim()==='保存至草稿箱');"
       "if(!b)return 0;b.id='__saveDraft';b.scrollIntoView({block:'center'});return 1})()")
    time.sleep(0.6)
    run('click', '#__saveDraft')
    try:
        wait_url('registerSoft', timeout=12)
    except RuntimeError:
        ev("(()=>{const b=document.getElementById('__saveDraft');if(!b)return 0;b.click();return 1})()")
        try:
            wait_url('registerSoft', timeout=18)
        except RuntimeError:
            shot = os.path.join(tempfile.gettempdir(), 'r11_save_fail.png')
            try:
                os.makedirs(os.path.dirname(shot), exist_ok=True)
                run('screenshot', shot, timeout=90)
            except Exception:
                pass
            tail = body_text(1200)
            raise RuntimeError('点保存后 URL 未跳回 registerSoft，草稿很可能没存上；'
                               f'截图 {shot}\n页面尾部：{tail[-400:]}')
    return '确认页核对通过 → 已保存草稿'


def read_serial(name, timeout=90):
    run('open', 'https://register.ccopyright.com.cn/account.html?current=soft_register')
    time.sleep(3)
    for _ in range(int(timeout / 3)):
        t = body_text(8000)
        m = re.search(r'流水号：(2026R11L\d+)\s*' + re.escape(name), t)
        if m:
            return m.group(1)
        time.sleep(3)
    return None


# ---------------------------------------------------------------- 主流程
def main():
    global AB, CDP, NO_SAVE
    ap = argparse.ArgumentParser()
    ap.add_argument('project')
    ap.add_argument('--check', action='store_true', help='只体检，不动表单')
    ap.add_argument('--dry-run', action='store_true', help='打印步骤，不动浏览器')
    ap.add_argument('--stop-after', type=int, default=5,
                    help='只跑到第 N 步就停（N=4 可验证选择器而不产生草稿）')
    ap.add_argument('--no-save', action='store_true',
                    help='走到确认页核对完就停，不点保存（验证全链路而不产生第二条草稿）')
    ap.add_argument('--date', default=None, help='开发完成日期 YYYY-MM-DD，默认今天-60 天')
    ap.add_argument('--category', default=None, help='软件分类，默认从运行平台推断')
    ap.add_argument('--cdp', default='9222')
    ap.add_argument('--no-launch', action='store_true',
                    help='端口不通时不要自动起 Edge（默认会起，复用同一 profile 保登录态）')
    args = ap.parse_args()
    CDP = str(args.cdp)
    NO_SAVE = args.no_save
    AB = find_ab()

    root = os.path.abspath(args.project)
    plan_p = os.path.join(root, '01_申请表', 'fill_plan.json')
    if not os.path.exists(plan_p):
        sys.exit(f'✗ 缺 {plan_p}，先跑 prep_fill.py')
    plan = json.load(open(plan_p, encoding='utf-8'))
    f = plan['fields']
    name, ver = plan['software_name'].strip(), plan['version']

    done_date = args.date or (dt.date.today() - dt.timedelta(days=60)).isoformat()
    cat = args.category
    if not cat:
        rp = (f.get('run_platform') or '') + (f.get('support_sw') or '')
        cat = '嵌入式软件' if re.search(r'RT-Thread|FreeRTOS|MicroPython|裸机|固件|MCU|Linux', rp) else '应用软件'
    langs = [x for x in re.split(r'[、,，/\s]+', f.get('lang') or '') if x]
    src_pdf = os.path.join(root, '02_源程序', '源程序鉴别材料_V1.0.pdf')
    doc_pdf = os.path.join(root, '03_说明书', '软件设计说明书_V1.0.pdf')

    print(f'软件全称：{name}  {ver}')
    print(f'分类={cat}  完成日期={done_date}  语言={langs}  行数={f.get("source_lines")}')

    if args.dry_run:
        for i, s in enumerate(['选我是申请人', '填全称/版本', f'填分类/日期', '填功能与特点+上传',
                               '核对确认页+存草稿', '取流水号'], start=1):
            print(f'  {i}. {s}')
        print('（dry-run，未动浏览器）')
        return

    state = ensure_browser(launch=not args.no_launch)
    if state != 'existing':
        print(f'· 受控窗口已关闭 → 自动重启 {state}（同一 profile，但登录态基本已失效，需重登）')
    check_login()
    if args.check:
        u = url()
        t = body_text(200)
        print(f'✓ 已登录　当前页：{u}')
        print(f'✓ 页面文本：{t[:60]}')
        print(f'✓ 计划字段数：{len(f)}　PDF：{os.path.exists(src_pdf)}/{os.path.exists(doc_pdf)}')
        print('（--check 只读，未动表单）')
        return

    tags = f.get('tech_tags') or []
    steps = [
        (1, lambda: step1_identity()),
        (2, lambda: step2_application(name, ver)),
        (3, lambda: step3_development(cat, done_date)),
        (4, lambda: step4_features(f, src_pdf, doc_pdf, langs)),
        (5, lambda: step5_confirm({
            '软件全称': name, '版本号': ver, '软件分类': cat, '开发完成日期': done_date,
            '软件说明': '原创', '开发方式': '单独开发', '发表状态': '未发表',
            '开发的硬件环境': f['dev_hw'][:18], '运行的硬件环境': f['run_hw'][:18],
            '软件的主要功能': f['main_func'][:24], '软件的技术特点': f['tech_feature'][:16],
            **{f'技术特点标签[{tag}]': tag for tag in tags},
        })),
    ]
    log = []
    t0 = time.time()
    for n, fn in steps:
        if n > args.stop_after:
            print(f'（--stop-after {args.stop_after}：停在第 {args.stop_after} 步之后，表单未保存）')
            json.dump(log, open(os.path.join(root, '05_提交记录', '填表断点.json'), 'w', encoding='utf-8'),
                      ensure_ascii=False, indent=2) if os.path.isdir(os.path.join(root, '05_提交记录')) else None
            return
        s = time.time()
        msg = fn()
        log.append({'step': n, 'ok': True, 'msg': msg, 'sec': round(time.time() - s, 1)})
        print(f'✓ 第{n}步 {msg}  ({round(time.time()-s,1)}s)')

    serial = None if NO_SAVE else read_serial(name)
    out = {
        'software': name, 'version': ver, 'serial': serial, 'status': '暂存',
        'category': cat, 'done_date': done_date, 'langs': langs, 'tech_tags': tags,
        'source_lines': f.get('source_lines'), 'steps': log,
        'elapsed_sec': round(time.time() - t0, 1),
        'at': dt.datetime.now().isoformat(timespec='seconds'),
    }
    rec_dir = os.path.join(root, '05_提交记录')
    os.makedirs(rec_dir, exist_ok=True)
    json.dump(out, open(os.path.join(rec_dir, '填表结果.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=2)
    print(f'\n流水号：{serial or ("未保存（--no-save）" if NO_SAVE else "未取到，去草稿箱核对")}')
    print(f'总耗时 {out["elapsed_sec"]}s　结果写入 05_提交记录/填表结果.json')
    if not serial:
        sys.exit(1)


if __name__ == '__main__':
    main()
