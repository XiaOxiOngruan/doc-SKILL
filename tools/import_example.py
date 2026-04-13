"""
import_example.py
将用户提供的 .docx 文件导入为示例文档。

用法：
    python3 tools/import_example.py <docx路径> [--label 描述标签]
    python3 tools/import_example.py <目录路径>   # 批量导入目录下所有 .docx

示例：
    python3 tools/import_example.py ~/Downloads/通知示例.docx
    python3 tools/import_example.py ~/Downloads/通知示例.docx --label "财务部通知示例"
    python3 tools/import_example.py ~/Downloads/公文示例/        # 批量导入
"""

import sys
import os
import shutil
import argparse
import json
from datetime import datetime


EXAMPLES_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "examples"))
INDEX_PATH   = os.path.join(EXAMPLES_DIR, "index.json")


def load_index() -> list:
    if os.path.exists(INDEX_PATH):
        with open(INDEX_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_index(index: list):
    with open(INDEX_PATH, "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=2)


def import_one(src_path: str, label: str = None) -> dict:
    """导入单个 .docx 文件到 examples/ 目录"""
    os.makedirs(EXAMPLES_DIR, exist_ok=True)

    filename = os.path.basename(src_path)
    dest_path = os.path.join(EXAMPLES_DIR, filename)

    # 若目标已存在，加时间戳避免覆盖
    if os.path.exists(dest_path):
        stem, ext = os.path.splitext(filename)
        ts = datetime.now().strftime("%Y%m%d%H%M%S")
        filename = f"{stem}_{ts}{ext}"
        dest_path = os.path.join(EXAMPLES_DIR, filename)

    shutil.copy2(src_path, dest_path)

    record = {
        "filename": filename,
        "label":    label or os.path.splitext(filename)[0],
        "source":   src_path,
        "imported_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    return record, dest_path


def main():
    parser = argparse.ArgumentParser(description="导入 .docx 示例文档到 examples/ 目录")
    parser.add_argument("path", help=".docx 文件路径 或 包含 .docx 文件的目录路径")
    parser.add_argument("--label", "-l", default=None, help="示例描述标签（单文件导入时有效）")
    args = parser.parse_args()

    src = os.path.expanduser(args.path)

    if not os.path.exists(src):
        print(f"❌ 路径不存在：{src}")
        sys.exit(1)

    index = load_index()
    imported = []

    # ── 目录：批量导入 ──────────────────────────
    if os.path.isdir(src):
        files = [f for f in os.listdir(src) if f.endswith(".docx") and not f.startswith("~")]
        if not files:
            print(f"⚠️  目录下没有找到 .docx 文件：{src}")
            sys.exit(0)

        print(f"📂 批量导入目录：{src}")
        print(f"   共找到 {len(files)} 个文件\n")

        for fname in sorted(files):
            fpath = os.path.join(src, fname)
            try:
                record, dest = import_one(fpath)
                index.append(record)
                imported.append(dest)
                print(f"  ✅ {fname} → examples/{record['filename']}")
            except Exception as e:
                print(f"  ❌ {fname} 导入失败：{e}")

    # ── 单文件导入 ──────────────────────────────
    else:
        if not src.endswith(".docx"):
            print(f"❌ 仅支持 .docx 格式文件")
            sys.exit(1)

        print(f"📄 导入文件：{src}")
        record, dest = import_one(src, args.label)
        index.append(record)
        imported.append(dest)
        print(f"✅ 已导入：{dest}")
        if args.label:
            print(f"   标签：{args.label}")

    save_index(index)

    print(f"\n📋 示例库现共有 {len(index)} 个文件")
    print(f"   索引位置：{INDEX_PATH}")

    if imported:
        print(f"\n已导入文件：")
        for p in imported:
            print(f"  {p}")


if __name__ == "__main__":
    main()
