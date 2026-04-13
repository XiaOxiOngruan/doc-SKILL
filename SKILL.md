---
name: govdoc
description: 中国公文撰写与生成技能。当用户需要起草通知、请示、报告、批复等公文，或将已有内容排版生成标准 .docx 文件时使用。默认遵循 GB/T 9704 国家标准，支持导入自定义格式文件。
user-invocable: true
allowed-tools:
  - Bash
  - Read
  - Write
  - Glob
---

# 中国公文撰写与生成技能

## 能力范围

1. **公文内容撰写**：根据用户提供的主题、背景、要求，辅助撰写标准公文正文
2. **公文排版生成**：将内容排版为符合国标的 .docx 文件
3. **多类型支持**：通知、请示、报告、批复、函、决定等
4. **格式灵活**：内置 GB/T 9704 标准，支持导入自定义 .yaml/.json 格式文件
5. **模板导入**：从用户提供的 .docx 文件自动解析排版参数，生成格式配置文件
6. **示例库管理**：导入、管理用户自有示例文档

---

## 用户命令

### 从规范文本直接生成格式配置

当用户粘贴一段公文排版规范说明（如单位内部规范文档中的字体字号页边距描述），执行：

```bash
# 直接传入文本
python3 tools/import_format_text.py --name "单位名称" --text "
字体、字号：一级标题（三号黑体）
二级标题（三号楷体-GB2312 加粗）
其余各级标题及正文（三号仿宋-GB2312）
结构层次序数：一、（一）1.（1）
行间距：固定值28磅
页边距：上3.67cm，下3.49cm，左右2.68cm
页码："—1—"，四号宋体，单页靠右，双页靠左
"

# 从 .txt 文件读取规范
python3 tools/import_format_text.py 规范说明.txt --name "单位名称"

# 交互式粘贴（直接运行后粘贴文本，Ctrl+D 结束）
python3 tools/import_format_text.py --name "单位名称"

# 在某个已有格式基础上叠加修改（只覆盖差异字段）
python3 tools/import_format_text.py --name "定制格式" --base formats/default_gbt9704.yaml --text "..."
```

支持解析的字段：一级/二级/三级标题字体字号、正文字体字号、页边距、行间距、页码格式、结构层次序数。  
未能识别的字段自动继承基础格式，解析完成后会打印预览报告。

---

### 导入 .docx 作为新格式模板

当用户说"把这个 docx 设为模板"、"用这个文件的格式"、"导入我们单位的模板"时，执行：

```bash
# 基本用法（自动从文件名生成格式名）
python3 tools/import_template.py /path/to/模板文件.docx

# 指定格式名称
python3 tools/import_template.py /path/to/模板文件.docx --name "财务部公文格式"
```

执行后会：
1. 解析 .docx 的页边距、字体、字号、行距等排版参数
2. 在 `formats/` 目录生成 `<名称>.json` 格式配置文件
3. 备份原始 .docx 为 `formats/<名称>_reference.docx`
4. 输出解析结果预览，告知用户解析到的参数

**注意**：解析结果仅供参考，复杂样式可能需要人工微调 `formats/` 下生成的 JSON 文件。

---

### 导入示例 .docx 文档

当用户说"上传示例"、"把这个加到示例库"、"导入参考文档"时，执行：

```bash
# 导入单个文件
python3 tools/import_example.py /path/to/示例.docx

# 导入单个文件并添加描述标签
python3 tools/import_example.py /path/to/示例.docx --label "2023年度工作报告示例"

# 批量导入目录下所有 .docx
python3 tools/import_example.py /path/to/示例目录/
```

执行后会：
1. 将文件复制到 `examples/` 目录（自动处理重名冲突）
2. 更新 `examples/index.json` 示例索引
3. 输出示例库当前文件总数

---

## 文档类型与适用场景

| 类型 | 适用场景 | 模板文件 |
|------|---------|---------|
| 通知 | 转发文件、部署工作、传达事项 | `templates/tongzhi.py` |
| 请示 | 向上级申请批准事项 | `templates/qingshi.py` |
| 报告 | 工作汇报、情况说明、调研成果 | `templates/baogao.py` |
| 学术论文 | 毕业论文、学位论文、期刊投稿初稿 | `templates/lunwen.py` |

---

## 执行流程

### 第一步：理解用户需求

收集以下信息（未提供的逐一询问）：

- **公文类型**：通知 / 请示 / 报告 / 其他
- **标题**：如"关于XXX的通知"
- **发文机关**：单位全称
- **主送机关**：接收单位
- **发文字号**：如"办发〔2024〕1号"（可选）
- **核心内容**：主要事项、背景、要求
- **保存路径**：默认当前目录，文件名建议拼音/英文
- **是否使用自定义格式**：如有，提供 .yaml 或 .json 文件路径

### 第二步：辅助撰写正文

公文正文写作要求：
- **语言**：庄重、准确、简洁，避免口语化
- **结构**：背景/依据 → 主体事项 → 要求/期限
- **层次序数**：严格使用 `一、`→`（一）`→`1.`→`（1）`
- **结语**：通知用"请遵照执行"，请示用"当否，请批示"，报告用"特此报告"

### 第三步：生成 .docx 文件

根据文档类型调用对应模板脚本，示例：

```bash
# 生成通知
python3 templates/tongzhi.py

# 生成请示
python3 templates/qingshi.py

# 生成报告（使用自定义格式）
python3 templates/baogao.py

# 生成学术论文
python3 templates/lunwen.py
```

也可直接用 Python 调用函数：

```python
import sys
sys.path.insert(0, "/path/to/govdoc-skill")
from templates.tongzhi import build as build_tongzhi
from templates.lunwen import build as build_lunwen

# 公文
build_tongzhi(
    title="关于开展XXX工作的通知",
    doc_number="办发〔2024〕1号",
    issuer="某某机关办公室",
    issue_date="2024年1月1日",
    recipients="各部门、各单位",
    sections=[
        {"heading": "一、工作背景", "body": ["……"]},
        {"heading": "二、主要任务", "body": ["……"]},
    ],
    closing="请遵照执行。",
    output_path="output/通知.docx",
    format_path=None  # 传入 .json/.yaml 路径可使用自定义格式
)

# 学术论文
build_lunwen(
    title="论文标题",
    author="作者姓名",
    institution="某某大学",
    supervisor="导师姓名 教授",
    degree="工学硕士",
    date="2024年6月",
    abstract_zh="中文摘要内容……",
    keywords_zh=["关键词1", "关键词2"],
    title_en="Thesis Title in English",
    abstract_en="English abstract...",
    keywords_en=["keyword1", "keyword2"],
    chapters=[
        {
            "title": "第一章  绪论",
            "sections": [
                {"title": "1.1 研究背景", "body": ["……"]},
                {"title": "1.2 研究意义", "body": ["……"]},
            ]
        },
    ],
    references=["[1] 作者. 书名[M]. 出版社, 年份."],
    acknowledgement="感谢……",
    output_path="output/论文.docx"
)
```

### 论文模板结构

论文模板自动生成以下结构：

```
封面（标题、作者、机构、导师、学位、日期）
  ↓
中文摘要 + 关键词
  ↓
英文摘要 + Keywords（可选）
  ↓
目录（Word 域代码，打开后右键更新域即可）
  ↓
正文各章（三级标题：章 → 节 → 小节）
  ↓
参考文献（悬挂缩进格式）
  ↓
致谢（可选）
```

---

## 自定义格式说明

### 导入方式

将 `formats/custom_template.yaml` 复制并修改，然后在调用时传入路径：

```python
build(..., format_path="formats/my_org.yaml")
```

### 格式文件结构

```yaml
name: "格式名称"

page:
  top_cm: 3.7        # 上边距
  bottom_cm: 3.5     # 下边距
  left_cm: 2.8       # 左边距
  right_cm: 2.6      # 右边距

fonts:
  heading1:
    name: "方正小标宋简体"   # 字体名
    size_pt: 22             # 字号（磅）
    bold: false
    align: center           # left/center/right/justify
  body:
    name: "仿宋_GB2312"
    size_pt: 16
    align: justify

paragraph:
  line_spacing_pt: 28       # 行间距（固定值）
  first_line_indent_chars: 2

page_number:
  format: "—{n}—"           # {n} 为页码占位符
  odd_align: right           # 单页页码对齐
  even_align: left           # 双页页码对齐
```

**只需填写与默认值不同的字段**，未填写的自动继承 GB/T 9704 默认值。

---

## 内置格式文件

| 文件 | 说明 |
|------|------|
| `formats/default_gbt9704.yaml` | GB/T 9704 国家标准（默认） |
| `formats/exim_bank.yaml` | 中国进出口银行公文格式 |
| `formats/custom_template.yaml` | 自定义格式模板（可复制修改） |

---

## 执行规则

1. **默认格式**：不指定格式文件时，使用 GB/T 9704 国家标准
2. **不擅自改内容**：已有文档若无明确修改指令，不删除或变更原有内容
3. **保存路径**：文件名用拼音或英文，如 `tongzhi_20240101.docx`
4. **生成后验证**：输出绝对路径，并列出文档章节结构
5. **字体提示**：若系统缺少 `仿宋_GB2312`、`楷体_GB2312`、`方正小标宋简体`，提示用户安装或替换为系统已有字体

---

## 示例文档

`examples/` 目录包含三个开箱即用的示例：

- `examples/示例_通知.docx`
- `examples/示例_请示.docx`
- `examples/示例_报告.docx`
