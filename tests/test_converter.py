"""wps-dispimg-fixer 单元测试 (纯标准库, 无需第三方依赖)。

运行: python -m unittest discover -s tests -v
"""

from __future__ import annotations

import os
import re
import sys
import tempfile
import unittest
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from wps_dispimg_fixer.converter import (  # noqa: E402
    _normalize_media,
    cell_ref_to_rc,
    fix_workbook,
    has_wps_cellimage,
    scan_sheet_dispimg,
)
from wps_dispimg_fixer.image_utils import get_image_size  # noqa: E402
from tests.sample_factory import make_png, make_wps_xlsx  # noqa: E402


class TestCellRef(unittest.TestCase):
    def test_convert(self):
        self.assertEqual(cell_ref_to_rc("H2"), (7, 1))
        self.assertEqual(cell_ref_to_rc("A1"), (0, 0))
        self.assertEqual(cell_ref_to_rc("Z26"), (25, 25))
        self.assertEqual(cell_ref_to_rc("AA1"), (26, 0))

    def test_invalid(self):
        with self.assertRaises(ValueError):
            cell_ref_to_rc("!@#")


class TestImageUtils(unittest.TestCase):
    def test_png_size(self):
        png = make_png(320, 240, (10, 20, 30))
        self.assertEqual(get_image_size(png), (320, 240))

    def test_unknown(self):
        self.assertIsNone(get_image_size(b"not an image"))
        self.assertIsNone(get_image_size(b""))


class TestNormalizeMedia(unittest.TestCase):
    def test_variants(self):
        self.assertEqual(_normalize_media("media/image1.jpeg"), "image1.jpeg")
        self.assertEqual(_normalize_media("image1.jpeg"), "image1.jpeg")
        self.assertEqual(_normalize_media("../media/image1.jpeg"), "image1.jpeg")
        self.assertEqual(_normalize_media("/xl/media/image1.jpeg"), "image1.jpeg")
        self.assertEqual(_normalize_media("xl/media/image1.jpeg"), "image1.jpeg")


class TestDetect(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.wps_file = os.path.join(self.tmp, "wps_sample.xlsx")
        make_wps_xlsx(self.wps_file, cells=[("H2", "ID_001"), ("I2", "ID_002")])

    def test_detect(self):
        with zipfile.ZipFile(self.wps_file) as zf:
            part = has_wps_cellimage(zf)
            self.assertEqual(part, "xl/cellimages.xml")

    def test_scan(self):
        with zipfile.ZipFile(self.wps_file) as zf:
            cells = scan_sheet_dispimg(zf, "xl/worksheets/sheet1.xml")
        self.assertEqual(len(cells), 2)
        self.assertEqual((cells[0].col, cells[0].row, cells[0].img_id), (7, 1, "ID_001"))
        self.assertEqual((cells[1].col, cells[1].row, cells[1].img_id), (8, 1, "ID_002"))


class TestFix(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.wps_file = os.path.join(self.tmp, "wps_sample.xlsx")
        self.out_file = os.path.join(self.tmp, "fixed.xlsx")
        make_wps_xlsx(
            self.wps_file,
            cells=[("H2", "ID_001"), ("I2", "ID_002"), ("H3", "ID_003")],
            row_height=92.0,
            cellimage_ext=(3950335, 2962910),
        )

    def test_fix_structure(self):
        report = fix_workbook(self.wps_file, self.out_file)
        self.assertFalse(report.errors, report.errors)
        self.assertEqual(report.fixed, 3)
        self.assertEqual(report.missing, 0)
        self.assertTrue(os.path.exists(self.out_file))

        with zipfile.ZipFile(self.out_file) as zf:
            names = set(zf.namelist())
            # cellimages 私有部分已删除
            self.assertFalse(any("cellimages" in n for n in names))
            # 生成了标准 drawing
            self.assertIn("xl/drawings/drawing1.xml", names)
            self.assertIn("xl/drawings/_rels/drawing1.xml.rels", names)
            self.assertIn("xl/worksheets/_rels/sheet1.xml.rels", names)
            # sheet rels 指向 drawing
            srel = zf.read("xl/worksheets/_rels/sheet1.xml.rels").decode("utf-8")
            self.assertIn("relationships/drawing", srel)
            self.assertIn("../drawings/drawing1.xml", srel)
            # sheet 中已无 DISPIMG 与 WPS 命名空间
            sheet = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
            self.assertNotIn("DISPIMG", sheet)
            self.assertNotIn("wps.cn", sheet)
            self.assertIn('<drawing r:id="', sheet)
            # drawing 中 3 个锚点
            drawing = zf.read("xl/drawings/drawing1.xml").decode("utf-8")
            self.assertEqual(drawing.count("<xdr:oneCellAnchor>"), 3)
            # anchor 位置正确
            self.assertIn("<xdr:col>7</xdr:col>", drawing)
            self.assertIn("<xdr:row>1</xdr:row>", drawing)
            self.assertIn("<xdr:col>8</xdr:col>", drawing)
            # keep 模式: 保留原始尺寸
            self.assertIn('cx="3950335" cy="2962910"', drawing)
            # Content_Types 已更新
            ct = zf.read("[Content_Types].xml").decode("utf-8")
            self.assertNotIn("cellimage", ct)
            self.assertIn('PartName="/xl/drawings/drawing1.xml"', ct)
            # workbook rels 已移除 WPS 关系
            wb_rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
            self.assertNotIn("wps.cn", wb_rels)

    def test_fix_cell_size_mode(self):
        out2 = os.path.join(self.tmp, "fixed_cell.xlsx")
        report = fix_workbook(self.wps_file, out2, size_mode="cell")
        self.assertFalse(report.errors, report.errors)
        with zipfile.ZipFile(out2) as zf:
            drawing = zf.read("xl/drawings/drawing1.xml").decode("utf-8")
            # 行高 92pt -> 122.7px -> cy=1,168,4xx; 图片 200x150 -> aspect=1.333
            m = re.search(r'<xdr:ext cx="(\d+)" cy="(\d+)"/>', drawing)
            cx, cy = int(m.group(1)), int(m.group(2))
            self.assertAlmostEqual(cy / 9525, 92 * 96 / 72, delta=2)
            self.assertAlmostEqual(cx / cy, 200 / 150, delta=0.02)

    def test_fix_keeps_cell_style(self):
        report = fix_workbook(self.wps_file, self.out_file)
        self.assertFalse(report.errors, report.errors)
        with zipfile.ZipFile(self.out_file) as zf:
            sheet = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
            m = re.search(r'<c r="H2"([^>]*)/>', sheet)
            self.assertIsNotNone(m, "H2 单元格应保留但清空公式")
            self.assertIn('s="1"', m.group(1), "应保留原单元格样式")

    def test_media_roundtrip(self):
        report = fix_workbook(self.wps_file, self.out_file)
        self.assertFalse(report.errors, report.errors)
        with zipfile.ZipFile(self.out_file) as zf:
            rels = zf.read("xl/drawings/_rels/drawing1.xml.rels").decode("utf-8")
            targets = re.findall(r'Target="\.\./media/([^"]+)"', rels)
            self.assertEqual(len(targets), 3)
            for t in targets:
                self.assertIn(f"xl/media/{t}", zf.namelist())


class TestEdgeCases(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def test_missing_image(self):
        """DISPIMG 引用的 ID 在 cellimages.xml 中不存在 -> 报告缺失且不破坏文件。"""
        f = os.path.join(self.tmp, "missing.xlsx")
        make_wps_xlsx(
            f,
            cells=[("H2", "ID_001"), ("I2", "ID_NOT_FOUND")],
            images=[{"id": "ID_001", "descr": "a", "w": 200, "h": 150,
                     "rgb": (1, 2, 3), "file": "image1.png"}],
        )
        out = os.path.join(self.tmp, "missing_fixed.xlsx")
        report = fix_workbook(f, out)
        self.assertEqual(report.fixed, 1)
        self.assertEqual(report.missing, 1)
        with zipfile.ZipFile(out) as zf:
            sheet = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
            # 缺失图片的单元格应原样保留
            self.assertIn("ID_NOT_FOUND", sheet)

    def test_cellimages_without_dispimg(self):
        """有 cellimages 部件但没有 DISPIMG 单元格 -> 报告错误且不写输出。"""
        f = os.path.join(self.tmp, "orphan.xlsx")
        make_wps_xlsx(f, cells=[])  # 无 DISPIMG 单元格, 但仍带 cellimages 部件
        out = os.path.join(self.tmp, "orphan_fixed.xlsx")
        report = fix_workbook(f, out)
        self.assertTrue(report.errors)
        self.assertFalse(os.path.exists(out))

    def test_plain_xlsx(self):
        """普通 xlsx (完全无 WPS cellimages 部件) -> 报告无需修复。"""
        f = os.path.join(self.tmp, "plain.xlsx")
        make_wps_xlsx(f, cells=[("H2", "ID_001")])
        # 移除 cellimages 部件与关联, 模拟普通文件
        import zipfile
        with zipfile.ZipFile(f) as zin:
            items = {n: zin.read(n) for n in zin.namelist()}
        items.pop("xl/cellimages.xml")
        items.pop("xl/_rels/cellimages.xml.rels")
        wb_rels = items["xl/_rels/workbook.xml.rels"].decode("utf-8")
        wb_rels = re.sub(r'<Relationship [^>]*cellImage[^>]*/>', "", wb_rels)
        items["xl/_rels/workbook.xml.rels"] = wb_rels.encode("utf-8")
        ct = items["[Content_Types].xml"].decode("utf-8")
        ct = re.sub(r'<Override [^>]*cellimages[^>]*/>', "", ct)
        items["[Content_Types].xml"] = ct.encode("utf-8")
        with zipfile.ZipFile(f, "w") as zout:
            for n, d in items.items():
                zout.writestr(n, d)
        out = os.path.join(self.tmp, "plain_fixed.xlsx")
        report = fix_workbook(f, out)
        self.assertTrue(report.errors)
        self.assertFalse(os.path.exists(out))

    def test_check_only(self):
        """--check 模式不应生成输出文件。"""
        f = os.path.join(self.tmp, "check.xlsx")
        make_wps_xlsx(f, cells=[("H2", "ID_001")])
        report = fix_workbook(f, "", check_only=True)
        self.assertEqual(report.fixed, 1)
        self.assertFalse(report.errors)


if __name__ == "__main__":
    unittest.main()
