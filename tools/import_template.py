"""
import_template.py
将用户提供的 .docx 文件解析为格式配置文件，并注册为可用模板。

用法：
    python3 tools/import_template.py <docx路径> [--name 格式名称]

示例：
    python3 tools/import_template.py ~/Downloads/我的公文模板.docx
    python3 tools/import_template.py ~/Downloads/模板.docx --name "财务部公文格式"
"""

import sys
import os
import json
import shutil
import argparse
import re

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn


def pt_to_cm(emu):
    """EMU → cm（python-docx 内部单位是 EMU：1cm = 360000 EMU）"""
    if emu is None:
        return None
    return round(emu / 360000, 2)


def emu_to_pt(emu):
    """EMU → pt（1pt = 12700 EMU）"""
    if emu is None:
        return None
    return round(emu / 12700, 1)


def half_pt_to_pt(half):
    """half-point → pt"""
    if half is None:
        return None
    return round(half / 2, 1)


def detect_align(para):
    align_map = {
        "CENTER": "center",
        "LEFT": "left",
        "RIGHT": "right",
        "JUSTIFY": "justify",
        "DISTRIBUTE": "justify",
        None: "left"
    }
    al = para.alignment
    name = al.name if al is not None else None
    return align_map.get(name, "left")


def get_run_font_name(run):
    """优先取 eastAsia 中文字体名"""
    rPr = run._r.find(qn("w:rPr"))
    if rPr is not None:
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is not None:
            ea = rFonts.get(qn("w:eastAsia"))
            if ea:
                return ea
            ascii_ = rFonts.get(qn("w:ascii"))
            if ascii_:
                return ascii_
    if run.font.name:
        return run.font.name
    return None


def analyze_paragraph(para):
    """提取段落的字体、字号、加粗、对齐、行距信息"""
    info = {
        "font_name": None,
        "size_pt": None,
        "bold": False,
        "align": detect_align(para),
        "line_spacing_pt": None,
        "first_line_indent_chars": 0,
    }

    # 行距
    pf = para.paragraph_format
    if pf.line_spacing is not None:
        info["line_spacing_pt"] = emu_to_pt(pf.line_spacing)

    # 首行缩进（字符数估算）
    if pf.first_line_indent and pf.first_line_indent > 0:
        # 用字号估算：缩进量 / 字号 ≈ 字符数
        info["_first_line_indent_emu"] = pf.first_line_indent

    # 从 runs 取字体信息（取第一个有效 run）
    for run in para.runs:
        if run.text.strip():
            name = get_run_font_name(run)
            if name:
                info["font_name"] = name
            size_emu = run.font.size
            if size_emu:
                info["size_pt"] = emu_to_pt(size_emu)
            if run.bold:
                info["bold"] = True
            break

    return info


def extract_format(doc_path: str, format_name: str) -> dict:
    """从 .docx 文件中提取格式信息，生成格式配置字典"""
    doc = Document(doc_path)
    section = doc.sections[0]

    fmt = {
        "name": format_name,
        "page": {
            "top_cm":    pt_to_cm(section.top_margin),
            "bottom_cm": pt_to_cm(section.bottom_margin),
            "left_cm":   pt_to_cm(section.left_margin),
            "right_cm":  pt_to_cm(section.right_margin),
            "page_size": "A4"
        },
        "fonts": {},
        "paragraph": {
            "line_spacing_pt": 28,
            "first_line_indent_chars": 2,
            "space_before_pt": 0,
            "space_after_pt": 0
        },
        "structure": {
            "levels": ["一、", "（一）", "1.", "（1）"]
        },
        "page_number": {
            "format": "—{n}—",
            "odd_align": "right",
            "even_align": "left"
        }
    }

    # 逐段分析，按位置推断层级
    paragraphs = [p for p in doc.paragraphs if p.text.strip()]

    heading1_info = None
    heading2_info = None
    body_info = None
    line_spacings = []

    for i, para in enumerate(paragraphs):
        info = analyze_paragraph(para)
        if info["line_spacing_pt"]:
            line_spacings.append(info["line_spacing_pt"])

        style_name = para.style.name.lower() if para.style else ""

        if "heading 1" in style_name or "标题 1" in style_name:
            heading1_info = info
        elif "heading 2" in style_name or "标题 2" in style_name:
            heading2_info = info
        elif "normal" in style_name or "正文" in style_name or not info.get("bold"):
            if body_info is None and info["font_name"]:
                body_info = info

        # 前三段若无样式名，按顺序推断
        if i == 0 and not heading1_info and info["font_name"]:
            heading1_info = info
        elif i == 1 and not heading2_info and info["font_name"]:
            heading2_info = info
        elif i >= 2 and not body_info and info["font_name"]:
            body_info = info

    # 最常见行距作为全局行距
    if line_spacings:
        from collections import Counter
        most_common = Counter(line_spacings).most_common(1)[0][0]
        fmt["paragraph"]["line_spacing_pt"] = most_common

    def build_font_cfg(info, default_name, default_size, default_bold, default_align):
        if not info:
            return {"name": default_name, "size_pt": default_size, "bold": default_bold, "align": default_align}
        size = info["size_pt"] or default_size
        # 首行缩进字符数估算
        if info.get("_first_line_indent_emu") and size:
            chars = round(info["_first_line_indent_emu"] / (size * 12700))
            if 1 <= chars <= 4:
                fmt["paragraph"]["first_line_indent_chars"] = chars
        return {
            "name":    info["font_name"] or default_name,
            "size_pt": size,
            "bold":    info["bold"] or default_bold,
            "align":   info["align"] or default_align
        }

    fmt["fonts"]["heading1"] = build_font_cfg(heading1_info, "方正小标宋简体", 22, False, "center")
    fmt["fonts"]["heading2"] = build_font_cfg(heading2_info, "黑体",          16, False, "left")
    fmt["fonts"]["heading3"] = {"name": "楷体_GB2312", "size_pt": 16, "bold": False, "align": "left"}
    fmt["fonts"]["body"]     = build_font_cfg(body_info,     "仿宋_GB2312",   16, False, "justify")
    fmt["fonts"]["header"]   = {"name": "仿宋_GB2312", "size_pt": 14, "bold": False}
    fmt["fonts"]["footer"]   = {"name": "宋体",        "size_pt": 14, "bold": False}

    return fmt


def save_format(fmt: dict, output_dir: str, slug: str) -> str:
    """将格式配置保存为 JSON 文件"""
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"{slug}.json")
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(fmt, f, ensure_ascii=False, indent=2)
    return output_path


def slugify(name: str) -> str:
    """将名称转为合法文件名"""
    name = re.sub(r"[^\w\u4e00-\u9fff\-]", "_", name)
    return name.strip("_") or "custom"


def main():
    parser = argparse.ArgumentParser(description="将 .docx 模板解析为公文格式配置文件")
    parser.add_argument("docx_path", help=".docx 文件路径")
    parser.add_argument("--name", "-n", default=None, help="格式名称（默认使用文件名）")
    parser.add_argument(
        "--formats-dir", default=None,
        help="格式文件保存目录（默认：本项目 formats/ 目录）"
    )
    args = parser.parse_args()

    docx_path = os.path.expanduser(args.docx_path)
    if not os.path.exists(docx_path):
        print(f"❌ 文件不存在：{docx_path}")
        sys.exit(1)
    if not docx_path.endswith(".docx"):
        print(f"❌ 仅支持 .docx 格式文件")
        sys.exit(1)

    format_name = args.name or os.path.splitext(os.path.basename(docx_path))[0]
    formats_dir = args.formats_dir or os.path.join(os.path.dirname(__file__), "..", "formats")
    formats_dir = os.path.abspath(formats_dir)

    print(f"📄 解析文件：{docx_path}")
    print(f"📐 格式名称：{format_name}")

    fmt = extract_format(docx_path, format_name)

    slug = slugify(format_name)
    output_path = save_format(fmt, formats_dir, slug)

    # 同时保留原始 .docx 作为参考
    ref_path = os.path.join(formats_dir, f"{slug}_reference.docx")
    shutil.copy2(docx_path, ref_path)

    print(f"\n✅ 格式配置已保存：{output_path}")
    print(f"📎 原始模板已备份：{ref_path}")
    print(f"\n使用方式：")
    print(f'  build(..., format_path="{output_path}")')
    print(f"\n解析结果预览：")
    print(f"  页边距：上{fmt['page']['top_cm']}cm 下{fmt['page']['bottom_cm']}cm "
          f"左{fmt['page']['left_cm']}cm 右{fmt['page']['right_cm']}cm")
    print(f"  标题字体：{fmt['fonts']['heading1']['name']} {fmt['fonts']['heading1']['size_pt']}pt")
    print(f"  正文字体：{fmt['fonts']['body']['name']} {fmt['fonts']['body']['size_pt']}pt")
    print(f"  行间距：{fmt['paragraph']['line_spacing_pt']}pt")

    return output_path


if __name__ == "__main__":
    main()
