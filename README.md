# 软著申请一键部署

> 计算机软件著作权（软著）登记全流程代办技能 —— 写功能描述、出源程序与说明书、批量转 PDF、接管浏览器填表。

面向个人开发者与中小企业研发团队的自助申报工具。按中国版权保护中心 2026 版申请表要求，把一套原本要来回折腾两三天的材料准备工作压到一次对话里完成。

**它不代提交。** 登录、滑块验证、点「确认填报」始终由本人完成，技能只负责把材料备齐、把表单填好并**保存到草稿箱**。

---

## 能做什么

| 环节 | 产出 |
|---|---|
| 功能描述 | 500–1300 字、「主要功能和技术特点」一段到底不分点，按开发目的 / 面向领域 / 主要功能 / 技术特点 / 技术标签五维度组织 |
| 源程序鉴别材料 | 前 30 页 + 后 30 页共 60 页，每页 ≥50 行，自动加页眉版本号 |
| 软件说明书 | 设计说明书或使用说明书，每页 ≥30 行，含功能模块与流程说明 |
| Word → PDF | 批量转换（走 Word COM），按 `01_申请表 / 02_源程序 / 03_说明书` 归档 |
| 在线填表 | 接管 R11 五步表单，逐字段填写、逐步断言、确认页核对、存草稿取流水号 |
| 提交前自检 | 页数与每页行数、页眉版本号一致性、同批多件开发完成日期错开、技术环境六字段差异化 |

## 安装

技能支持两个层级，按需选一个：

```bash
# 用户级（所有项目可用，推荐）
git clone https://github.com/DDguai-ace/software-copyright-filing.git \
  ~/.workbuddy/skills/software-copyright-filing

# 项目级（仅当前工程可用）
git clone https://github.com/DDguai-ace/software-copyright-filing.git \
  .workbuddy/skills/software-copyright-filing
```

装好后无需重启，技能列表里会出现「软著申请一键部署」。

**依赖**（仅用到脚本时需要）：

```bash
pip install python-docx pywin32 pymupdf
npm i -g agent-browser        # 仅在自动填表时需要，不必装 Chromium
```

- Word → PDF 依赖本机已装 Microsoft Word 或 WPS；**不需要 pandoc、不需要 LibreOffice**
- 自动填表依赖本机 Edge 或 Chrome，用 `--remote-debugging-port` 起受控窗口

## 快速开始

装上之后直接说人话即可：

```
帮我办软著，软件叫「XX 设备管理软件」，Python 写的，代码在 D:\proj\src
```

技能会先问一次主体信息（著作权人、证件号、地址、联系人），然后按默认流程跑完：

1. 采集字段 → 生成功能描述
2. 生成源程序鉴别材料 + 说明书（Markdown）
3. 转 PDF、校验页数与每页行数
4. 汇总出 `fill_plan.json`
5. 起受控浏览器 → **等你手动过登录** → 自动填完五步 → 存草稿

中途只会因为三种情况停下来问你：主体信息缺失、登录需要本人、材料页数不达标。

### 脚本单独用

```bash
PY=python    # 换成你的解释器

# 1) Markdown → Word（带页眉版本号）
"$PY" scripts/md2docx.py "说明书.md" "说明书.docx" --header "XX软件  V1.0"

# 2) 源程序目录 → 60 页鉴别材料
"$PY" scripts/build_source.py "./src" "./output/02_源程序"

# 3) Word → PDF（目录批量 / 单文件）
"$PY" scripts/docx2pdf.py "./output/软著_XXX_V1.0"

# 4) 校验页数与每页行数
"$PY" scripts/check_pages.py "源程序.pdf" --min-lines 50 --header "XX软件"

# 5) 填表前体检（各字段字数、技术标签、材料齐备）
"$PY" scripts/prep_fill.py "./output/软著_XXX_V1.0"

# 6) 接管 R11 填表
"$PY" scripts/fill_r11.py "./output/软著_XXX_V1.0"

# 7) 五阶段一条龙
"$PY" scripts/run_project.py "./output/软著_XXX_V1.0"
```

`fill_r11.py` 的常用开关：

| 开关 | 用途 |
|---|---|
| `--check` | 只体检环境（登录态、计划、PDF），不动表单 |
| `--dry-run` | 打印步骤，不启动浏览器 |
| `--stop-after 4` | 验证到第四步就停，不产生草稿（调选择器用） |
| `--no-save` | 走到确认页核对完就停，不留垃圾草稿 |
| `--date` / `--category` | 覆盖默认的开发完成日期 / 软件分类 |

## 目录结构

```
.
├── SKILL.md                      # 技能主文件，9 个章节的完整 SOP
├── references/
│   ├── 原始五步流程.md            # 五步流程原话 + 默认值（不得偏离）
│   ├── 字段填写对照表.md          # 每个字段的取值范围与字数上限
│   ├── 功能描述模板.md            # 五维度结构 + 可改写范文
│   ├── R11填表实录.md             # 真实界面字段名、报错原文、控件坑位
│   └── 材料规范与补正.md          # 页数规则、周期基准、补正应对
└── scripts/
    ├── md2docx.py                # Markdown → docx，A4、页眉、页码
    ├── build_source.py           # 源程序 → 60 页鉴别材料
    ├── docx2pdf.py               # 批量转 PDF（Word COM，WPS 自动回退）
    ├── check_pages.py            # 页数 / 每页行数 / 页眉一致性校验
    ├── prep_fill.py              # 汇总 fill_plan.json + 填表前体检
    ├── fill_r11.py               # R11 五步自动填表（含断言与存草稿）
    └── run_project.py            # 五阶段编排
```

## 几条设计取舍

**默认值写死，不问第二遍。** 版本号 `V1.0`、软件简称留空、软件说明「原创」、开发方式「单独开发」、权利取得「原始取得」、交存方式「一般交存」——这些是所有申请都一样的字段，反复确认只浪费时间。

**开发完成日期按申请日往前推 60–70 天。** 同批多件必须错开，且与源程序里的注释时间、说明书里的编写日期自洽，否则容易被要求补正。

**技术环境六字段逐件差异化。** 同一批申报多件时，六项技术环境字段完全复制粘贴会出现明显批量痕迹，技能会检查并提示改写角度。

**只存草稿，不点确认填报。** 这是硬纪律。填完把确认页截图给你，核对无误后由你本人点提交。

## 已知限制

- 只在 Windows 上验证过（Word COM 转 PDF、Edge 受控窗口）；macOS 需要把 `docx2pdf.py` 换成 LibreOffice 或 `pandoc + wkhtmltopdf`
- R11 页面结构若改版，选择器会失效，需按 `references/R11填表实录.md` 重新取一次
- 功能描述的作者视角、行业术语需人工过一遍，工具保证的是结构与字数，不是内容质量

## 隐私

仓库内不含任何真实的著作权人信息、证件号、联系方式、真实申请流水号。`SKILL.md` §1 的主体信息是取值表，首次运行时由使用者自己填写，不落盘到仓库。

## License

[MIT](LICENSE)
