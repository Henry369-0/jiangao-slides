"""比较用同一环境导出的两份幻灯片 PDF，检查修改的视觉影响范围。"""

import argparse
import sys
from pathlib import Path

import pymupdf


def page_list(numbers):
    return "、".join(str(n) for n in sorted(numbers)) or "无"


def compare(before, after, allowed_pages):
    for label, document in (("修改前", before), ("修改后", after)):
        if not document.is_pdf or document.needs_pass or len(document) == 0:
            raise ValueError(f"{label}的输入须为可直接打开且至少有一页的 PDF。")

    if len(before) != len(after):
        print(f"页数发生变化：{len(before)} → {len(after)}。无法按原页码对照。")
        return 1

    invalid = {n for n in allowed_pages if not 1 <= n <= len(before)}
    if invalid:
        raise ValueError(f"页码须在 1–{len(before)} 范围内：{page_list(invalid)}。")

    changed = set()
    resized = set()
    for number, (left, right) in enumerate(zip(before, after), start=1):
        if left.rect != right.rect:
            changed.add(number)
            resized.add(number)
            print(f"第 {number} 页尺寸变化（pt）："
                  f"{left.rect.width:g} × {left.rect.height:g} → "
                  f"{right.rect.width:g} × {right.rect.height:g}")
            continue
        original = left.get_pixmap(dpi=144, colorspace=pymupdf.csRGB, alpha=False)
        updated = right.get_pixmap(dpi=144, colorspace=pymupdf.csRGB, alpha=False)
        if original.samples != updated.samples:
            changed.add(number)

    unexpected = changed - allowed_pages
    print(f"实际有视觉变化的页：{page_list(changed)}")
    print(f"声明范围外变化的页：{page_list(unexpected)}")
    if allowed_pages - changed:
        print(f"声明范围内未观察到视觉变化的页：{page_list(allowed_pages - changed)}")
    return 1 if unexpected or resized else 0


def main():
    parser = argparse.ArgumentParser(
        description=__doc__,
        epilog="退出码：0 表示范围内或无可见变化，1 表示范围外或尺寸/页数变化，2 表示输入有误。",
    )
    parser.add_argument("before", type=Path, help="修改前的 PDF")
    parser.add_argument("after", type=Path, help="修改后的 PDF")
    parser.add_argument("--changed", nargs="+", type=int, required=True,
                        metavar="页码", help="允许修改的页码，从 1 开始，多个页码用空格分隔")
    args = parser.parse_args()
    try:
        with pymupdf.open(args.before) as before, pymupdf.open(args.after) as after:
            return compare(before, after, set(args.changed))
    except (OSError, RuntimeError, ValueError, pymupdf.mupdf.FzErrorBase) as error:
        parser.exit(2, f"无法比较：{error}\n")


if __name__ == "__main__":
    sys.exit(main())
