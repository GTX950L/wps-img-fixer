"""生成一个可公开的演示用测试样本 (不含任何真实业务数据)。

用法:
    python scripts/make_test_sample.py [输出路径]

生成的样本模拟"WPS 单元格内嵌图片"的文件结构:
- 单元格里写 =DISPIMG("ID_xxx", 1) 公式
- 图片存在 WPS 私有的 xl/cellimages.xml
- 用本工具修复后, 图片转为标准 Excel 图片
"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from tests.sample_factory import make_wps_xlsx  # noqa: E402


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else "demo_wps_dispimg.xlsx"
    make_wps_xlsx(
        out,
        cells=[("H2", "ID_001"), ("I2", "ID_002"), ("H3", "ID_003"), ("I3", "ID_004")],
        row_height=92.0,
    )
    print(f"已生成演示样本: {out}")
    print("在 Excel 中打开会看到一串 =DISPIMG(...) 代码, 图片不显示。")
    print(f"修复: python -m wps_dispimg_fixer {out}")


if __name__ == "__main__":
    main()
