"""
import_format_text.py
将中文公文规范说明文字解析为格式配置文件（JSON）。

支持两种输入方式：
  1. 直接传入文本字符串（命令行 --text 参数）
  2. 传入 .txt 文件路径

用法：
    # 从文件读取规范文本
    python3 tools/import_format_text.py spec.txt --name "中国进出口银行"

    # 直接输入文本（用引号包裹）
    python3 tools/import_format_text.py --text "字体、字号：一级标题（三号黑体）..." --name "自定义格式"

    # 交互式输入（不传任何文本参数，直接回车后粘贴）
    python3 tools/import_format_text.py --name "自定义格式"

输出：
    在 formats/ 目录生成 <名称>.json，可直接用于 build(..., format_path=...)
"""

import sys
import os
import re
import json
import argparse

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.docx_utils import DEFAULT_FORMAT, _deep_merge

# ─────────────────────────────────────────────
# 字号映射表（中文字号 → 磅值 pt）
# ─────────────────────────────────────────────
FONT_SIZE_MAP = {
    "初号":  42.0,
    "小初":  36.0,
    "一号":  26.0,
    "小一":  24.0,
    "二号":  22.0,
    "小二":  18.0,
    "三号":  16.0,
    "小三":  15.0,
    "四号":  14.0,
    "小四":  12.0,
    "五号":  10.5,
    "小五":   9.0,
    "六号":   7.5,
    "小六":   6.5,
    "七号":   5.5,
    "八号":   5.0,
}

# ─────────────────────────────────────────────
# 字体名称规范化（处理常见别名/错误拼写）
# ─────────────────────────────────────────────
FONT_ALIAS = {
    "仿宋":          "仿宋_GB2312",
    "仿宋gb2312":    "仿宋_GB2312",
    "仿宋-gb2312":   "仿宋_GB2312",
    "仿宋_gb2312":   "仿宋_GB2312",
    "仿宋cb2312":    "仿宋_GB2312",   # 原文 "CB2312" 为笔误
    "仿宋-cb2312":   "仿宋_GB2312",
    "楷体":          "楷体_GB2312",
    "楷体gb2312":    "楷体_GB2312",
    "楷体-gb2312":   "楷体_GB2312",
    "楷体_gb2312":   "楷体_GB2312",
    "黑体":          "黑体",
    "宋体":          "宋体",
    "微软雅黑":      "微软雅黑",
    "方正小标宋":    "方正小标宋简体",
    "方正小标宋简体": "方正小标宋简体",
    "arial":         "Arial",
    "times new roman": "Times New Roman",
}

def normalize_font(raw: str) -> str:
    """规范化字体名称"""
    key = raw.strip().lower().replace(" ", "")
    return FONT_ALIAS.get(key, raw.strip())

def parse_font_size(raw: str) -> object:
    """解析字号文字，返回 pt 值"""
    raw = raw.strip()
    # 中文字号
    for name, pt in FONT_SIZE_MAP.items():
        if name in raw:
            return pt
    # 直接写磅值，如 "16pt" / "16磅" / "16"
    m = re.search(r"(\d+(?:\.\d+)?)\s*(?:pt|磅|点)?$", raw)
    if m:
        return float(m.group(1))
    return None

def parse_cm(raw: str) -> object:
    """解析厘米值，如 '3.67cm' / '3.67'"""
    m = re.search(r"(\d+(?:\.\d+)?)\s*(?:cm|厘米)?", raw)
    return float(m.group(1)) if m else None

def parse_align(raw: str) -> str:
    """解析对齐方式"""
    if any(k in raw for k in ["居中", "center", "中"]):
        return "center"
    if any(k in raw for k in ["靠右", "右对齐", "right"]):
        return "right"
    if any(k in raw for k in ["靠左", "左对齐", "left"]):
        return "left"
    if any(k in raw for k in ["两端", "justify"]):
        return "justify"
    return "left"

# ─────────────────────────────────────────────
# 核心解析函数
# ─────────────────────────────────────────────

def parse_spec_text(text: str) -> dict:
    """
    将规范说明文本解析为格式配置字典（只包含解析到的字段）。
    未解析到的字段不写入，由调用方与默认值合并。
    """
    result = {}

    lines = [l.strip() for l in text.replace("。", "\n").replace("；", "\n").splitlines()]
    full  = " ".join(lines)   # 也保留完整文本用于整体匹配

    # ── 1. 一级标题 ──────────────────────────
    heading1 = _parse_heading_level(full, ["一级标题", "文件标题", "标题一"])
    if heading1:
        result.setdefault("fonts", {})["heading1"] = heading1

    # ── 2. 二级标题 ──────────────────────────
    heading2 = _parse_heading_level(full, ["二级标题", "标题二"])
    if heading2:
        result.setdefault("fonts", {})["heading2"] = heading2

    # ── 3. 三级标题 ──────────────────────────
    heading3 = _parse_heading_level(full, ["三级标题", "标题三"])
    if heading3:
        result.setdefault("fonts", {})["heading3"] = heading3

    # ── 4. 正文（"其余各级标题及正文" / "正文"）────
    body = _parse_body_font(full)
    if body:
        result.setdefault("fonts", {})["body"] = body

    # ── 5. 页边距 ────────────────────────────
    margins = _parse_margins(full)
    if margins:
        result["page"] = margins

    # ── 6. 行间距 ────────────────────────────
    ls = _parse_line_spacing(full)
    if ls is not None:
        result.setdefault("paragraph", {})["line_spacing_pt"] = ls

    # ── 7. 字间距 ────────────────────────────
    # 标准字间距不需要特殊设置，仅记录
    if "字间距" in full and "标准" in full:
        result.setdefault("paragraph", {})["char_spacing"] = "standard"

    # ── 8. 页码格式 ──────────────────────────
    pn = _parse_page_number(full)
    if pn:
        result["page_number"] = pn

    # ── 9. 结构层次序数 ──────────────────────
    levels = _parse_structure_levels(full)
    if levels:
        result["structure"] = {"levels": levels}

    return result


def _parse_heading_level(text: str, keywords: list) -> object:
    """从文本中解析某一标题级别的字体配置"""
    for kw in keywords:
        # 匹配模式：一级标题（三号黑体）或 一级标题：三号黑体加粗
        pattern = rf"{re.escape(kw)}[（(：:\s]+([^）)]+)[）)]?"
        m = re.search(pattern, text)
        if m:
            spec = m.group(1).strip()
            return _parse_font_spec(spec)
    return None


def _parse_body_font(text: str) -> object:
    """解析正文字体（'其余各级标题及正文' 或 '正文'）"""
    patterns = [
        r"其余各级标题及正文[（(：:\s]+([^）)\n]+)",
        r"其余标题及正文[（(：:\s]+([^）)\n]+)",
        r"正文[（(：:\s]+([^）)\n，,]+)",
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            return _parse_font_spec(m.group(1).strip())
    return None


def _parse_font_spec(spec: str) -> dict:
    """
    从字体规格字符串中提取字体名、字号、加粗、对齐。
    例如：'三号黑体'、'三号楷体-GB2312 加粗'、'仿宋_GB2312 16pt 居中'
    """
    cfg = {"name": None, "size_pt": None, "bold": False, "align": "left"}

    # 加粗
    if any(k in spec for k in ["加粗", "bold", "Bold"]):
        cfg["bold"] = True

    # 对齐
    cfg["align"] = parse_align(spec)

    # 字号（先匹配，避免字号关键词干扰字体名提取）
    size_pt = parse_font_size(spec)
    if size_pt:
        cfg["size_pt"] = size_pt
        # 从 spec 中去掉字号文字，再提取字体名
        for name in FONT_SIZE_MAP:
            spec = spec.replace(name, "")

    # 字体名（从剩余文字中匹配已知字体）
    # 按长度降序匹配，避免"楷体"误匹配"楷体_GB2312"
    candidates = sorted(FONT_ALIAS.keys(), key=len, reverse=True)
    spec_lower = spec.lower().replace(" ", "")
    for alias in candidates:
        if alias in spec_lower:
            cfg["name"] = FONT_ALIAS[alias]
            break

    # 若未匹配到已知字体，尝试直接提取中文字体名
    if not cfg["name"]:
        m = re.search(r"([\u4e00-\u9fff_\-A-Za-z0-9]+体[\w\-_]*)", spec)
        if m:
            cfg["name"] = normalize_font(m.group(1))

    return {k: v for k, v in cfg.items() if v is not None and v is not False}


def _parse_margins(text: str) -> object:
    """
    解析页边距，支持多种写法：
    - 上3.67cm，下3.49cm，左右2.68cm
    - 上3.67cm，下3.49cm，左2.68cm，右2.68cm
    - 页边距：上3.7 下3.5 左2.8 右2.6（cm）
    """
    result = {}

    # 上边距
    m = re.search(r"上\s*(\d+(?:\.\d+)?)\s*(?:cm|厘米)?", text)
    if m: result["top_cm"] = float(m.group(1))

    # 下边距
    m = re.search(r"下\s*(\d+(?:\.\d+)?)\s*(?:cm|厘米)?", text)
    if m: result["bottom_cm"] = float(m.group(1))

    # 左右统一
    m = re.search(r"左右\s*(\d+(?:\.\d+)?)\s*(?:cm|厘米)?", text)
    if m:
        result["left_cm"]  = float(m.group(1))
        result["right_cm"] = float(m.group(1))
    else:
        # 左右分开
        m = re.search(r"(?<![左右上下])左\s*(\d+(?:\.\d+)?)\s*(?:cm|厘米)?", text)
        if m: result["left_cm"] = float(m.group(1))
        m = re.search(r"(?<![左右上下])右\s*(\d+(?:\.\d+)?)\s*(?:cm|厘米)?", text)
        if m: result["right_cm"] = float(m.group(1))

    return result if result else None


def _parse_line_spacing(text: str) -> object:
    """解析行间距，如 '固定值28磅' / '28pt' / '1.5倍行距'"""
    # 固定值磅数
    m = re.search(r"(?:固定值|行距|行间距)[^\d]*(\d+(?:\.\d+)?)\s*磅", text)
    if m: return float(m.group(1))

    m = re.search(r"行[间距距]*[：:\s]*(\d+(?:\.\d+)?)\s*(?:磅|pt)", text)
    if m: return float(m.group(1))

    return None


def _parse_page_number(text: str) -> object:
    """解析页码格式，如 '—1—，四号宋体，单页靠右，双页靠左'"""
    result = {}

    # 页码格式字符串
    m = re.search(u'[\u201c\u201d\u2018\u2019\"\']([^\u201c\u201d\u2018\u2019\"\']{1,20})[\u201c\u201d\u2018\u2019\"\']', text)
    if m:
        fmt_raw = m.group(1)
        # 将数字替换为 {n} 占位符
        fmt_normalized = re.sub(r"\d+", "{n}", fmt_raw)
        result["format"] = fmt_normalized

    # 页码字体
    m = re.search(r"页码[^，,。\n]*?([一二三四五六七八小初]号)", text)
    if not m:
        m = re.search(r"(?:页码|页数)[^，,。\n]*?([\u4e00-\u9fff]+体[\w\-_]*)", text)
    if m:
        result["font_name"] = normalize_font(m.group(1))

    # 页码字号
    m = re.search(r"页码[^，,。\n]*?([一二三四五六七八小初]号)", text)
    if m:
        size = parse_font_size(m.group(1))
        if size:
            result["font_size_pt"] = size

    # 奇偶对齐
    if "单页靠右" in text or "奇数页靠右" in text:
        result["odd_align"] = "right"
    if "双页靠左" in text or "偶数页靠左" in text:
        result["even_align"] = "left"
    if "单页靠左" in text:
        result["odd_align"] = "left"
    if "双页靠右" in text:
        result["even_align"] = "right"
    if "居中" in text and "页码" in text:
        result["odd_align"] = "center"
        result["even_align"] = "center"

    return result if result else None


def _parse_structure_levels(text: str) -> object:
    """解析结构层次序数，如 '一、（一）1.（1）'"""
    m = re.search(r"结构层次序数[：:\s]*(.{4,40}?)(?:\s|行|字|页|$)", text)
    if not m:
        return None
    raw = m.group(1).strip()
    # 按常见分隔符拆分
    parts = re.split(r"[→\-\s]+", raw)
    parts = [p.strip() for p in parts if p.strip()]
    return parts if len(parts) >= 2 else None


# ─────────────────────────────────────────────
# 生成人类可读的解析报告
# ─────────────────────────────────────────────

def format_report(merged: dict) -> str:
    lines = ["", "📋 解析结果："]

    fonts = merged.get("fonts", {})
    for level, label in [("heading1", "一级标题"), ("heading2", "二级标题"),
                          ("heading3", "三级标题"), ("body", "正文")]:
        f = fonts.get(level, {})
        if f:
            bold_str = " 加粗" if f.get("bold") else ""
            lines.append(f"  {label}：{f.get('name', '?')} {f.get('size_pt', '?')}pt{bold_str}  对齐:{f.get('align','?')}")

    page = merged.get("page", {})
    if page:
        lines.append(f"  页边距：上{page.get('top_cm')} 下{page.get('bottom_cm')} "
                     f"左{page.get('left_cm')} 右{page.get('right_cm')} cm")

    para = merged.get("paragraph", {})
    if para.get("line_spacing_pt"):
        lines.append(f"  行间距：{para['line_spacing_pt']} 磅（固定值）")

    pn = merged.get("page_number", {})
    if pn:
        lines.append(f"  页码格式：{pn.get('format', '—{{n}}—')}  "
                     f"单页:{pn.get('odd_align','?')} 双页:{pn.get('even_align','?')}")

    struct = merged.get("structure", {})
    if struct.get("levels"):
        lines.append(f"  层次序数：{'  '.join(struct['levels'])}")

    return "\n".join(lines)


# ─────────────────────────────────────────────
# 主函数
# ─────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="将中文公文规范说明解析为格式配置文件")
    parser.add_argument("file", nargs="?", help=".txt 规范文件路径（可选）")
    parser.add_argument("--text", "-t", default=None, help="直接传入规范文本字符串")
    parser.add_argument("--name", "-n", required=True, help="格式名称，如 '中国进出口银行'")
    parser.add_argument(
        "--formats-dir", default=None,
        help="格式文件保存目录（默认：本项目 formats/ 目录）"
    )
    parser.add_argument("--base", default=None,
        help="基础格式文件路径（默认继承 GB/T 9704），新解析的字段会覆盖基础格式")
    args = parser.parse_args()

    # ── 读取输入文本 ──────────────────────────
    if args.text:
        spec_text = args.text
    elif args.file:
        path = os.path.expanduser(args.file)
        if not os.path.exists(path):
            print(f"❌ 文件不存在：{path}")
            sys.exit(1)
        with open(path, "r", encoding="utf-8") as f:
            spec_text = f.read()
    else:
        print("📝 请粘贴公文规范文本，输入完成后按 Enter 然后 Ctrl+D（macOS/Linux）或 Ctrl+Z+Enter（Windows）：")
        try:
            spec_text = sys.stdin.read()
        except KeyboardInterrupt:
            print("\n已取消")
            sys.exit(0)

    if not spec_text.strip():
        print("❌ 未输入任何内容")
        sys.exit(1)

    # ── 加载基础格式 ──────────────────────────
    if args.base:
        base_path = os.path.expanduser(args.base)
        import json as _json
        with open(base_path, "r", encoding="utf-8") as f:
            base_fmt = _json.load(f)
        print(f"📎 基础格式：{base_path}")
    else:
        base_fmt = DEFAULT_FORMAT.copy()
        print("📎 基础格式：GB/T 9704 国家标准（默认）")

    # ── 解析文本 ──────────────────────────────
    print(f"🔍 解析规范文本 → 格式名称：{args.name}")
    parsed = parse_spec_text(spec_text)

    if not parsed:
        print("⚠️  未从文本中解析到任何格式信息，请检查输入格式")
        sys.exit(1)

    # ── 与基础格式合并 ────────────────────────
    merged = _deep_merge(base_fmt, parsed)
    merged["name"] = args.name

    # ── 保存 JSON ─────────────────────────────
    formats_dir = args.formats_dir or os.path.join(os.path.dirname(__file__), "..", "formats")
    formats_dir = os.path.abspath(formats_dir)
    os.makedirs(formats_dir, exist_ok=True)

    slug = re.sub(r"[^\w\u4e00-\u9fff\-]", "_", args.name).strip("_") or "custom"
    output_path = os.path.join(formats_dir, f"{slug}.json")

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(merged, f, ensure_ascii=False, indent=2)

    print(f"\n✅ 格式配置已保存：{output_path}")
    print(f"\n使用方式：")
    print(f'  build(..., format_path="{output_path}")')

    # ── 打印解析报告 ──────────────────────────
    print(format_report(merged))

    print("\n💡 提示：可直接编辑上述 JSON 文件微调未能自动识别的参数")
    return output_path


if __name__ == "__main__":
    main()
