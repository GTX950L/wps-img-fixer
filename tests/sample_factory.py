"""测试样本工厂: 用纯标准库构造带 WPS 单元格内嵌图片 (cellimages.xml) 的 xlsx。

生成的样本结构与真实 WPS 文件一致, 但不含任何真实业务数据,
可以安全地用于单元测试和随仓库公开。

用法:
    from sample_factory import make_wps_xlsx
    make_wps_xlsx("sample.xlsx", cells=[("H2", "ID_001"), ("I3", "ID_002")])
"""

from __future__ import annotations

import struct
import zipfile
import zlib
from typing import Dict, List, Tuple


# ---------------------------------------------------------------------------
# 最小合法 OOXML 部件
# ---------------------------------------------------------------------------

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/cellimages.xml" ContentType="application/vnd.wps-officedocument.cellimage+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""

APP_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>WPS Office</Application></Properties>"""

CORE_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>test</dc:title></cp:coreProperties>"""

WORKBOOK_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"""

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf></cellXfs>
</styleSheet>"""

THEME_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
<a:themeElements>
<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
<a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2>
<a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4>
<a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6>
<a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>
<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/></a:majorFont>
<a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
</a:themeElements></a:theme>"""


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------

def make_png(width: int, height: int, rgb: Tuple[int, int, int]) -> bytes:
    """生成纯色 PNG (RGB, 8bit)。"""
    def chunk(tag: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    raw = b"".join(b"\x00" + bytes(rgb) * width for _ in range(height))
    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", zlib.compress(raw, 6))
            + chunk(b"IEND", b""))


def _cell_ref_to_rc(ref: str) -> Tuple[int, int]:
    """'H2' -> (col, row) 0-based。"""
    letters = "".join(ch for ch in ref if ch.isalpha())
    digits = "".join(ch for ch in ref if ch.isdigit())
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch.upper()) - ord("A") + 1)
    return col - 1, int(digits) - 1


# ---------------------------------------------------------------------------
# 生成 WPS 风格 xlsx
# ---------------------------------------------------------------------------

def make_wps_xlsx(path: str,
                  cells: List[Tuple[str, str]],
                  images: List[Dict] = None,
                  row_height: float = 92.0,
                  col_width: float = 18.0,
                  cellimage_ext: Tuple[int, int] = (3950335, 2962910)) -> str:
    """构造一个带 WPS 单元格内嵌图片的 xlsx。

    cells: [(单元格引用, 图片ID)], 例如 [("H2", "ID_001"), ("I2", "ID_002")]
    images: 图片定义列表, 每项 {id, descr, w, h, rgb, file}
            缺省时为 cells 中每个 ID 生成一张 200x150 纯色图。
    返回文件路径。
    """
    if images is None:
        images = []
        palette = [(200, 60, 60), (60, 120, 200), (60, 160, 80), (200, 160, 60)]
        for i, (ref, img_id) in enumerate(cells):
            rgb = palette[i % len(palette)]
            images.append({"id": img_id, "descr": f"sample-{img_id}",
                           "w": 200, "h": 150, "rgb": rgb,
                           "file": f"image{i + 1}.png"})

    id_to_img = {img["id"]: img for img in images}

    # 1. 构造 sheet1.xml
    max_col, max_row = 0, 0
    for ref, _ in cells:
        col, row = _cell_ref_to_rc(ref)
        max_col, max_row = max(max_col, col), max(max_row, row)

    cols_xml = (f'<cols><col min="1" max="{max_col + 1}" width="{col_width}" customWidth="1"/></cols>'
                if max_col >= 0 else "")
    # 表头 + 数据行
    rows_xml = ['<row r="1" ht="38" customHeight="1" spans="1:%d"><c r="A1" t="s"><v>0</v></c></row>'
                % (max_col + 1)]
    for r in range(1, max_row + 2):
        row_attrs = f'<row r="{r}" ht="{row_height}" customHeight="1" spans="1:{max_col + 1}">'
        cells_xml = ""
        for ref, img_id in cells:
            col, row0 = _cell_ref_to_rc(ref)
            if row0 + 1 == r:
                cells_xml += (f'<c r="{ref}" s="1" t="str">'
                              f'<f>_xlfn.DISPIMG(&quot;{img_id}&quot;,1)</f>'
                              f'<v>=DISPIMG(&quot;{img_id}&quot;,1)</v></c>')
        # 留一个普通文本单元格
        if not cells_xml:
            cells_xml = f'<c r="A{r}" t="inlineStr"><is><t>row {r}</t></is></c>'
        rows_xml.append(row_attrs + cells_xml + "</row>")

    sheet1_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData">
<sheetPr/><dimension ref="A1:{_rc_to_ref(max_col, max_row)}"/>
<sheetViews><sheetView workbookViewId="0"/></sheetViews>
<sheetFormatPr defaultColWidth="9" defaultRowHeight="49" customHeight="1"/>
{cols_xml}
<sheetData>{"".join(rows_xml)}</sheetData>
<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>"""

    # 2. 构造 cellimages.xml (WPS 私有)
    cellimage_xml = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                     f'<etc:cellImages xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
                     f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                     f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                     f'xmlns:etc="{_NS_ET}">']
    cellimage_rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    cx, cy = cellimage_ext
    for i, img in enumerate(images, start=1):
        rid = f"rId{i}"
        cellimage_xml.append(
            f'<etc:cellImage><xdr:pic><xdr:nvPicPr>'
            f'<xdr:cNvPr id="{i + 1}" name="{img["id"]}" descr="{img["descr"]}"/>'
            f'<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
            f'<xdr:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
            f'<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic></etc:cellImage>')
        cellimage_rels.append(
            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            f'Target="media/{img["file"]}"/>')
    cellimage_xml.append("</etc:cellImages>")
    cellimage_rels.append("</Relationships>")

    # 3. workbook.xml.rels (带 WPS cellImage 关系)
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
        '<Relationship Id="rId4" Type="http://www.wps.cn/officeDocument/2020/cellImage" Target="cellimages.xml"/>'
        '</Relationships>')

    # 4. 图片
    media = {}
    for img in images:
        media[f"xl/media/{img['file']}"] = make_png(img.get("w", 200), img.get("h", 150),
                                                    tuple(img.get("rgb", (128, 128, 128))))

    # 5. 打包
    parts = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": ROOT_RELS,
        "docProps/app.xml": APP_XML,
        "docProps/core.xml": CORE_XML,
        "xl/workbook.xml": WORKBOOK_XML,
        "xl/_rels/workbook.xml.rels": workbook_rels,
        "xl/worksheets/sheet1.xml": sheet1_xml,
        "xl/styles.xml": STYLES_XML,
        "xl/theme/theme1.xml": THEME_XML,
        "xl/cellimages.xml": "\n".join(cellimage_xml),
        "xl/_rels/cellimages.xml.rels": "\n".join(cellimage_rels),
    }
    parts.update(media)

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
    return path


_NS_ET = "http://www.wps.cn/officeDocument/2017/etCustomData"


def _rc_to_ref(col: int, row: int) -> str:
    """0-based (col, row) -> 'H2'"""
    letters = ""
    c = col + 1
    while c:
        c, rem = divmod(c - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return f"{letters}{row + 1}"


if __name__ == "__main__":
    out = "test_sample_wps.xlsx"
    make_wps_xlsx(
        out,
        cells=[("H2", "ID_001"), ("I2", "ID_002"), ("H3", "ID_003")],
    )
    print(f"已生成测试样本: {out}")
