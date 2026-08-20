"""命令行入口: 修复 WPS 单元格内嵌图片 (DISPIMG) 在 Excel 中不显示的问题。

用法示例:
    python -m wps_dispimg_fixer 报告.xlsx
    python -m wps_dispimg_fixer 报告.xlsx -o 修复后.xlsx
    python -m wps_dispimg_fixer 文件夹/                 # 批量处理目录下所有 xlsx
    python -m wps_dispimg_fixer 报告.xlsx --check        # 只检测不修复
"""

from __future__ import annotations

import argparse
import os
import sys

from .converter import FixReport, fix_workbook
from . import __version__


def _format_report(r: FixReport, check_only: bool) -> str:
    lines = []
    lines.append(f"文件: {r.input_path}")
    if r.errors:
        for e in r.errors:
            lines.append(f"  [!] {e}")
        return "\n".join(lines)
    lines.append(f"  检测: 包含 WPS 单元格内嵌图片 (cellimages.xml), 共 {r.total_images} 张")
    for s in r.sheets:
        lines.append(f"  工作表 {os.path.basename(s.sheet_file)}:")
        lines.append(f"    转换图片: {len(s.anchors)} 个")
        if s.cleared_cells:
            lines.append(f"    清除公式: {len(s.cleared_cells)} 个单元格"
                         f" (首行 {s.cleared_cells[0]}...)" if len(s.cleared_cells) > 1
                         else f"    清除公式: {len(s.cleared_cells)} 个单元格")
        if s.missing_ids:
            lines.append(f"    [!] 缺少对应图片: {len(s.missing_ids)} 个引用 "
                         f"(如 {s.missing_ids[0]}), 未转换")
    if check_only:
        lines.append(f"  结论: DISPIMG 单元格 {r.total_dispimg} 个, "
                     f"可转换 {r.fixed} 个, 缺失 {r.missing} 个")
    else:
        lines.append(f"  完成: 转换 {r.fixed} 张图片, 清除 {r.total_dispimg - r.missing} 个公式, "
                     f"缺失 {r.missing} 个")
        lines.append(f"  输出: {r.output_path}")
    return "\n".join(lines)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="wps-dispimg-fixer",
        description="修复 WPS 单元格内嵌图片 (DISPIMG) 在 Excel 中不显示、只剩一串代码的问题",
    )
    parser.add_argument("input", nargs="+", help="要修复的 .xlsx 文件或包含 xlsx 的文件夹")
    parser.add_argument("--check", action="store_true", help="只检测并报告, 不生成输出文件")
    parser.add_argument("-o", "--output", help="输出文件路径 (仅单文件输入时可用)")
    parser.add_argument("--size", choices=["keep", "cell"], default="keep",
                        help="图片尺寸模式: keep=保留 WPS 原始显示尺寸(默认, 最忠实), "
                             "cell=按所在行高缩放(贴合单元格, 需行高已设置)")
    parser.add_argument("--keep-wps-part", action="store_true",
                        help="保留原 cellimages.xml 文件 (默认删除, 与手工修复行为一致)")
    parser.add_argument("-V", "--version", action="version", version=f"%(prog)s {__version__}")
    args = parser.parse_args(argv)

    if args.output and len(args.input) > 1:
        print("错误: -o 只能用于单个输入文件", file=sys.stderr)
        return 2

    files = []
    for item in args.input:
        if os.path.isdir(item):
            files.extend(
                os.path.join(item, f) for f in sorted(os.listdir(item))
                if f.lower().endswith(".xlsx") and not f.lower().endswith(".xlsm")
            )
        elif os.path.isfile(item):
            files.append(item)
        else:
            print(f"错误: 找不到路径 {item}", file=sys.stderr)
            return 2

    if not files:
        print("错误: 没有找到任何 .xlsx 文件", file=sys.stderr)
        return 2

    exit_code = 0
    for path in files:
        if args.check:
            report = fix_workbook(path, "", size_mode=args.size, check_only=True)
            print(_format_report(report, check_only=True))
        else:
            out = args.output
            if not out:
                base, ext = os.path.splitext(path)
                out = f"{base}_fixed{ext}"
            report = fix_workbook(path, out, size_mode=args.size,
                                  keep_wps_part=args.keep_wps_part)
            print(_format_report(report, check_only=False))
        if report.errors:
            exit_code = 1
        print()
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
