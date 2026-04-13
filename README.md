# govdoc-skill

> 中国公文撰写与生成 Claude Code 技能，遵循 GB/T 9704 国家标准，支持导入 .docx 模板和自定义格式。

## 功能

- **内容撰写辅助**：通知、请示、报告、学术论文正文起草
- **一键生成 .docx**：自动排版，字体、行距、页边距、页码全部符合对应规范
- **导入 .docx 为模板**：从现有文档自动解析排版参数，生成可复用的格式配置
- **文本解析格式**：粘贴中文排版规范说明，自动生成格式配置文件
- **示例库管理**：导入、管理自有参考文档
- **自定义格式**：通过 `.yaml` 或 `.json` 文件导入机构专属排版规范

## 快速开始

### 安装依赖

```bash
pip install python-docx
pip install pyyaml   # 可选，使用 YAML 格式文件时需要
```

### 安装到 Claude Code

```bash
# 个人全局安装
cp -r govdoc-skill ~/.claude/skills/govdoc

# 或项目级安装
cp -r govdoc-skill .claude/skills/govdoc
```

重启 Claude Code 后，输入 `/govdoc` 即可使用。

### 直接调用（无需 Claude Code）

```python
from templates.tongzhi import build

build(
    title="关于开展2024年安全生产检查的通知",
    doc_number="办发〔2024〕8号",
    issuer="某某机关办公室",
    issue_date="2024年4月1日",
    recipients="各部门、各单位",
    sections=[
        {"heading": "一、检查范围", "body": ["本次检查覆盖全体在职人员及所有办公场所。"]},
        {"heading": "二、检查内容", "body": ["重点检查消防设施、用电安全、危险品管理三个方面。"]},
    ],
    closing="请遵照执行。",
    output_path="output/安全检查通知.docx"
)
```

---

## 工具命令

### 从规范文本直接生成格式配置

将中文排版规范说明（字体字号、页边距、行间距等）直接解析为可用的格式配置文件：

```bash
# 直接传入文本
python3 tools/import_format_text.py --name "中国进出口银行" --text "
字体、字号：一级标题（三号黑体）
二级标题（三号楷体-GB2312 加粗）
其余各级标题及正文（三号仿宋-GB2312）
结构层次序数：一、（一）1.（1）
行间距：固定值28磅
页边距：上3.67cm，下3.49cm，左右2.68cm
页码：\"—1—\"，四号宋体，单页靠右，双页靠左
"

# 从 .txt 文件读取规范
python3 tools/import_format_text.py 规范说明.txt --name "单位名称"

# 交互式粘贴（运行后直接粘贴文本，Ctrl+D 结束）
python3 tools/import_format_text.py --name "单位名称"

# 在某个已有格式基础上叠加（只覆盖差异字段）
python3 tools/import_format_text.py --name "定制格式" --base formats/default_gbt9704.yaml --text "..."
```

输出示例：
```
📋 解析结果：
  一级标题：黑体 16.0pt  对齐:left
  二级标题：楷体_GB2312 16.0pt 加粗  对齐:left
  正文：仿宋_GB2312 16.0pt  对齐:left
  页边距：上3.67 下3.49 左2.68 右2.68 cm
  行间距：28.0 磅（固定值）
  页码格式：—{n}—  单页:right 双页:left
  层次序数：一、  （一）  1.  （1）
```

---

### 导入 .docx 作为格式模板

自动解析 .docx 文件的排版参数（页边距、字体、字号、行距），生成可复用的格式配置：

```bash
# 基本用法（格式名自动取文件名）
python3 tools/import_template.py ~/Downloads/我的模板.docx

# 指定格式名称
python3 tools/import_template.py ~/Downloads/模板.docx --name "财务部公文格式"
```

执行后：
- 在 `formats/` 生成 `<名称>.json` 格式配置文件
- 备份原始文件为 `formats/<名称>_reference.docx`
- 输出解析结果预览（页边距、字体、行距等）

使用解析出的格式：

```python
build(..., format_path="formats/财务部公文格式.json")
```

> 解析结果仅供参考，复杂样式可能需要手动微调生成的 JSON 文件。

---

### 导入示例文档

```bash
# 导入单个文件
python3 tools/import_example.py ~/Downloads/通知示例.docx

# 附加描述标签
python3 tools/import_example.py ~/Downloads/通知示例.docx --label "2024年安全通知示例"

# 批量导入目录下所有 .docx
python3 tools/import_example.py ~/Downloads/公文示例/
```

导入的文件保存在 `examples/`，索引记录在 `examples/index.json`。

---

## 目录结构

```
govdoc-skill/
├── SKILL.md                    # Claude Code 技能入口
├── README.md
├── utils/
│   └── docx_utils.py           # 排版工具函数（字体、行距、页码等）
├── templates/
│   ├── tongzhi.py              # 通知模板
│   ├── qingshi.py              # 请示模板
│   ├── baogao.py               # 报告模板
│   └── lunwen.py               # 学术论文模板（毕业/学位/期刊）
├── tools/
│   ├── import_format_text.py   # 从中文规范文本解析生成格式配置
│   ├── import_template.py      # 从 .docx 解析并导入为格式配置
│   └── import_example.py       # 导入用户示例文档
├── formats/
│   ├── default_gbt9704.yaml    # GB/T 9704 国家标准（默认）
│   ├── academic_thesis.json    # 学术论文格式（高校通用）
│   └── custom_template.yaml    # 自定义格式模板（可复制修改）
└── examples/
    ├── index.json              # 示例库索引（导入后自动生成）
    ├── 示例_通知.docx
    ├── 示例_请示.docx
    ├── 示例_报告.docx
    └── 示例_论文.docx
```

---

## 自定义格式说明

复制 `formats/custom_template.yaml`，按需修改后传入 `format_path` 参数：

```python
build(..., format_path="formats/my_org.yaml")
```

只需填写与默认值不同的字段，其余自动继承 GB/T 9704 默认值。

### 格式参数说明

| 参数 | 说明 | GB/T 9704 默认值 |
|------|------|----------------|
| `page.top_cm` | 上边距 | 3.7 cm |
| `page.bottom_cm` | 下边距 | 3.5 cm |
| `page.left_cm` | 左边距 | 2.8 cm |
| `page.right_cm` | 右边距 | 2.6 cm |
| `fonts.heading1.name` | 文件标题字体 | 方正小标宋简体 |
| `fonts.heading1.size_pt` | 文件标题字号 | 22pt（二号） |
| `fonts.body.name` | 正文字体 | 仿宋_GB2312 |
| `fonts.body.size_pt` | 正文字号 | 16pt（三号） |
| `paragraph.line_spacing_pt` | 行间距 | 28磅（固定值） |
| `page_number.format` | 页码格式 | `—{n}—` |

## 内置格式

| 格式文件 | 适用场景 |
|---------|---------|
| `formats/default_gbt9704.yaml` | 通用国家机关公文 |
| `formats/exim_bank.yaml` | 中国进出口银行内部公文 |

## 注意事项

- 需要安装中文字体：`仿宋_GB2312`、`楷体_GB2312`、`方正小标宋简体`（macOS 默认不含，可从 Windows 系统复制或下载）
- Python 版本要求：3.8+

## License

MIT
