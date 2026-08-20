"""核心转换逻辑: 把 WPS 单元格内嵌图片 (DISPIMG) 转为标准 OOXML 图片。

转换步骤:
1. 解析 xl/cellimages.xml (WPS 私有) 得到每张图片的 ID/描述/嵌入关系/尺寸
2. 扫描每个工作表, 找出所有 =DISPIMG("ID_xxx",1) 公式单元格
3. 为每个这样的单元格生成标准 xdr:oneCellAnchor (xl/drawings/drawingN.xml)
4. 清除单元格里的 DISPIMG 公式 (保留样式), 写入 <drawing> 引用
5. 更新 sheet rels / workbook rels / [Content_Types].xml, 删除 WPS 私有部分
6. 重新打包输出新的 .xlsx
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from typing import Dict, List, Optional
from xml.etree import ElementTree as ET

from .image_utils import get_image_size

# ---------------------------------------------------------------------------
# 常量
# ---------------------------------------------------------------------------

NS_ET = "http://www.wps.cn/officeDocument/2017/etCustomData"
REL_WPS_CELLIMAGE = "http://www.wps.cn/officeDocument/2020/cellImage"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
CT_CELLIMAGE = "application/vnd.wps-officedocument.cellimage+xml"
CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
CT_RELS = "application/vnd.openxmlformats-package.relationships+xml"
CT_SHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"

EMU_PER_PX = 9525  # 1 像素 = 9525 EMU

ET.register_namespace("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")
ET.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

# 匹配 <c r="H2" s="6" t="str"><f>...</f><v>...</v></c>
_CELL_RE = re.compile(r'<c r="([A-Z]+\d+)"([^>]*)>(.*?)</c>', re.S)
# 匹配 DISPIMG 公式里的图片 ID, 兼容 &quot; 与直接引号两种写法
_DISPIMG_RE = re.compile(r'(?:DISPIMG|_xlfn\.DISPIMG)\s*\(\s*&quot;([^&"]+)&quot;|(?:DISPIMG|_xlfn\.DISPIMG)\s*\(\s*"([^"]+)"', re.I)
# 单元格引用 "H2" -> (col=7, row=1), 0-based
_CELL_REF_RE = re.compile(r'^([A-Z]+)(\d+)$')


# ---------------------------------------------------------------------------
# 数据结构
# ---------------------------------------------------------------------------

@dataclass
class CellImage:
    """cellimages.xml 中定义的一张 WPS 单元格图片。"""
    img_id: str          # cNvPr@name, 例如 ID_62B52A37...
    descr: str           # cNvPr@descr, 通常是文件名/来源信息
    cnvpr_id: int        # cNvPr@id
    embed_r_id: str      # blip@r:embed, 例如 rId3
    off_x: int           # xfrm/off@x (EMU)
    off_y: int
    ext_cx: int          # xfrm/ext@cx (EMU)
    ext_cy: int
    media_file: str = ""  # 对应的 xl/media/ 图片文件名


@dataclass
class DispImgCell:
    """工作表中引用 DISPIMG 的一个单元格。"""
    sheet_file: str      # xl/worksheets/sheetN.xml
    cell_ref: str        # 例如 H2
    col: int             # 0-based 列号
    row: int             # 0-based 行号
    img_id: str


@dataclass
class SheetFixResult:
    sheet_file: str
    drawing_part: str          # 生成的 drawing part, 如 xl/drawings/drawing1.xml
    anchors: List[dict] = field(default_factory=list)   # 生成的图片锚点信息
    cleared_cells: List[str] = field(default_factory=list)  # 被清除公式的单元格
    missing_ids: List[str] = field(default_factory=list)    # 找不到图片的 DISPIMG ID


@dataclass
class FixReport:
    input_path: str
    output_path: str = ""
    has_wps_cellimage: bool = False
    cellimage_part: str = ""
    total_images: int = 0                 # cellimages.xml 中的图片数
    total_dispimg: int = 0                # 所有工作表的 DISPIMG 单元格数
    fixed: int = 0                        # 成功转换的图片数
    missing: int = 0                      # 找不到对应图片的引用数
    sheets: List[SheetFixResult] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# 单元格引用与行列转换
# ---------------------------------------------------------------------------

def cell_ref_to_rc(ref: str):
    """'H2' -> (col=7, row=1), 0-based。"""
    m = _CELL_REF_RE.match(ref)
    if not m:
        raise ValueError(f"无效的单元格引用: {ref}")
    letters, digits = m.groups()
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch.upper()) - ord("A") + 1)
    return col - 1, int(digits) - 1


# ---------------------------------------------------------------------------
# WPS 私有部分检测
# ---------------------------------------------------------------------------

def has_wps_cellimage(zf: zipfile.ZipFile) -> Optional[str]:
    """检查工作簿是否带 WPS 单元格内嵌图片, 返回 cellimages part 路径或 None。"""
    names = set(zf.namelist())
    # 方式1: workbook.xml.rels 中存在 WPS cellImage 类型关系
    if "xl/_rels/workbook.xml.rels" in names:
        rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
        if REL_WPS_CELLIMAGE in rels:
            m = re.search(r'Target="([^"]*cellimages[^"]*\.xml)"', rels)
            if m:
                return _normalize_part(m.group(1))
    # 方式2: Content_Types 中声明了 WPS cellimage
    if "[Content_Types].xml" in names:
        ct = zf.read("[Content_Types].xml").decode("utf-8", "replace")
        if CT_CELLIMAGE in ct:
            m = re.search(r'PartName="(/xl/[^"]*cellimages[^"]*\.xml)"', ct)
            if m:
                return _normalize_part(m.group(1))
    # 方式3: 直接找 cellimages part
    for n in names:
        if re.match(r"xl/cellimages\d*\.xml$", n):
            return n
    return None


def _normalize_part(part: str) -> str:
    """把关系 Target 中的相对/绝对路径规范化为 zip 内部路径, 如 xl/cellimages.xml。"""
    part = part.replace("\\", "/")
    if part.startswith("/"):
        part = part.lstrip("/")
    elif part.startswith("../"):
        # 从 xl/ 目录看过去的相对路径
        part = part[3:]
        if not part.startswith("xl/"):
            part = "xl/" + part
    elif not part.startswith("xl/"):
        part = "xl/" + part
    return part


# ---------------------------------------------------------------------------
# 解析 cellimages.xml
# ---------------------------------------------------------------------------

def parse_cellimages(zf: zipfile.ZipFile, part: str) -> List[CellImage]:
    """解析 WPS cellimages.xml, 返回图片列表。"""
    ns = {
        "etc": NS_ET,
        "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }
    root = ET.fromstring(zf.read(part))
    rels_file = _rels_path(part)
    r_id_to_media: Dict[str, str] = {}
    try:
        rels_root = ET.fromstring(zf.read(rels_file))
        rns = "http://schemas.openxmlformats.org/package/2006/relationships"
        for rel in rels_root.findall(f"{{{rns}}}Relationship"):
            if rel.get("Type") == REL_IMAGE:
                r_id_to_media[rel.get("Id")] = rel.get("Target", "")
    except KeyError:
        pass  # 没有 rels 文件时按 rId 顺序猜测

    images: List[CellImage] = []
    for ci in root.findall(f"{{{NS_ET}}}cellImage"):
        pic = ci.find("xdr:pic", ns)
        if pic is None:
            continue
        nv = pic.find("xdr:nvPicPr/xdr:cNvPr", ns)
        blip = pic.find("xdr:blipFill/a:blip", ns)
        off = pic.find("xdr:spPr/a:xfrm/a:off", ns)
        ext = pic.find("xdr:spPr/a:xfrm/a:ext", ns)
        img_id = nv.get("name", "") if nv is not None else ""
        if not img_id:
            continue
        media_file = ""
        if blip is not None:
            rid = blip.get(f"{{{ns['r']}}}embed") or blip.get("r:embed") or ""
            media_file = r_id_to_media.get(rid, "")
            if not media_file:
                # 回退: 按 cellimage 顺序猜测 media/imageN
                idx = len(images) + 1
                media_file = f"media/image{idx}.jpeg"
        media_file = _normalize_media(media_file)
        images.append(CellImage(
            img_id=img_id,
            descr=nv.get("descr", "") if nv is not None else "",
            cnvpr_id=int(nv.get("id", "0")) if nv is not None else 0,
            embed_r_id=blip.get(f"{{{ns['r']}}}embed", "") if blip is not None else "",
            off_x=int(off.get("x", "0")) if off is not None else 0,
            off_y=int(off.get("y", "0")) if off is not None else 0,
            ext_cx=int(ext.get("cx", "0")) if ext is not None else 0,
            ext_cy=int(ext.get("cy", "0")) if ext is not None else 0,
            media_file=media_file,
        ))
    return images


def _rels_path(part: str) -> str:
    """'xl/cellimages.xml' -> 'xl/_rels/cellimages.xml.rels'"""
    dir_, name = part.rsplit("/", 1)
    return f"{dir_}/_rels/{name}.rels"


def _normalize_media(media: str) -> str:
    """把关系 Target 中的图片路径规范化为纯文件名 (如 image1.jpeg)。"""
    media = media.replace("\\", "/").strip()
    if media.startswith("/"):
        media = media.lstrip("/")
    # 去掉 media/ 前缀与 ../ 前缀
    if media.startswith("media/"):
        media = media[len("media/"):]
    while media.startswith("../"):
        media = media[3:]
    # 若还带着目录, 只取文件名
    if "/" in media:
        media = media.rsplit("/", 1)[-1]
    return media


# ---------------------------------------------------------------------------
# 扫描工作表中的 DISPIMG 单元格
# ---------------------------------------------------------------------------

def scan_sheet_dispimg(zf: zipfile.ZipFile, sheet_file: str) -> List[DispImgCell]:
    """扫描单个工作表, 返回所有 DISPIMG 单元格。"""
    xml = zf.read(sheet_file).decode("utf-8", "replace")
    cells: List[DispImgCell] = []
    for m in _CELL_RE.finditer(xml):
        ref, attrs, inner = m.group(1), m.group(2), m.group(3)
        hit = _DISPIMG_RE.search(inner)
        if not hit:
            continue
        img_id = hit.group(1) or hit.group(2)
        try:
            col, row = cell_ref_to_rc(ref)
        except ValueError:
            continue
        cells.append(DispImgCell(sheet_file=sheet_file, cell_ref=ref,
                                 col=col, row=row, img_id=img_id))
    return cells


# ---------------------------------------------------------------------------
# 生成标准 drawing 相关 XML
# ---------------------------------------------------------------------------

def _build_drawing_xml(anchors: List[dict]) -> str:
    """生成标准 xdr:wsDr (oneCellAnchor 列表) XML。"""
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    ]
    for a in anchors:
        parts.append("  <xdr:oneCellAnchor>")
        parts.append(f"    <xdr:from><xdr:col>{a['col']}</xdr:col><xdr:colOff>0</xdr:colOff>"
                     f"<xdr:row>{a['row']}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>")
        parts.append(f'    <xdr:ext cx="{a["cx"]}" cy="{a["cy"]}"/>')
        parts.append("    <xdr:pic>")
        parts.append("      <xdr:nvPicPr>")
        parts.append(f'        <xdr:cNvPr id="{a["cnvpr_id"]}" name="{_xml_escape(a["name"])}" descr="{_xml_escape(a["descr"])}"/>')
        parts.append('        <xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>')
        parts.append("      </xdr:nvPicPr>")
        parts.append(f'      <xdr:blipFill><a:blip r:embed="{a["r_id"]}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>')
        parts.append(f'      <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{a["cx"]}" cy="{a["cy"]}"/></a:xfrm>'
                     '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>')
        parts.append("    </xdr:pic>")
        parts.append("    <xdr:clientData/>")
        parts.append("  </xdr:oneCellAnchor>")
    parts.append("</xdr:wsDr>")
    return "\n".join(parts)


def _build_rels_xml(rels: List[tuple]) -> str:
    """生成 .rels XML。rels: [(Id, Type, Target), ...]"""
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ]
    for rid, rtype, target in rels:
        parts.append(f'  <Relationship Id="{rid}" Type="{rtype}" Target="{target}"/>')
    parts.append("</Relationships>")
    return "\n".join(parts)


def _xml_escape(text: str) -> str:
    """XML 属性转义。"""
    return (text.replace("&", "&amp;").replace('"', "&quot;")
                .replace("<", "&lt;").replace(">", "&gt;"))


# ---------------------------------------------------------------------------
# 工作表 XML 清洗
# ---------------------------------------------------------------------------

def clean_sheet_xml(xml: str, clear_refs: List[str]) -> str:
    """清除 DISPIMG 单元格的公式内容 (保留位置与样式), 添加 drawing 引用。

    clear_refs: 需要清空的单元格引用列表, 例如 ["H2", "I2"]。
    返回 (新xml, drawing 引用是否已添加由调用方决定)。
    """
    clear_set = set(clear_refs)

    def repl(m: re.Match) -> str:
        ref, attrs, inner = m.group(1), m.group(2), m.group(3)
        if ref not in clear_set:
            return m.group(0)
        # 保留样式属性 s="N"
        s_m = re.search(r'\bs="\d+"', attrs)
        if s_m:
            return f'<c r="{ref}" {s_m.group(0)}/>'
        return f'<c r="{ref}"/>'

    xml = _CELL_RE.sub(repl, xml)
    # 若文档中不再有 WPS 私有命名空间引用, 移除其声明
    if "etc:" not in xml:
        xml = re.sub(r'\s+xmlns:etc="[^"]*"', "", xml, count=1)
    return xml


# ---------------------------------------------------------------------------
# 尺寸计算
# ---------------------------------------------------------------------------

def _parse_row_heights(sheet_xml: str) -> Dict[int, float]:
    """解析行高(磅)映射: row_index(0-based) -> height(pt)。"""
    heights: Dict[int, float] = {}
    for m in re.finditer(r"<row\b[^>]*>", sheet_xml):
        tag = m.group(0)
        r_m = re.search(r'\br="(\d+)"', tag)
        ht_m = re.search(r'\bht="([\d.]+)"', tag)
        if r_m and ht_m:
            heights[int(r_m.group(1)) - 1] = float(ht_m.group(1))
    return heights


def _parse_col_widths(sheet_xml: str, default_width: float = 9.0) -> Dict[int, float]:
    """解析列宽(字符单位)映射: col_index(0-based) -> width。"""
    widths: Dict[int, float] = {}
    for m in re.finditer(r"<col\b[^>]*/>", sheet_xml):
        tag = m.group(0)
        min_m = re.search(r'\bmin="(\d+)"', tag)
        max_m = re.search(r'\bmax="(\d+)"', tag)
        w_m = re.search(r'\bwidth="([\d.]+)"', tag)
        if not (min_m and max_m):
            continue
        cmin, cmax = int(min_m.group(1)) - 1, int(max_m.group(1)) - 1
        w = float(w_m.group(1)) if w_m else default_width
        for c in range(cmin, cmax + 1):
            widths[c] = w
    return widths


def _get_default_row_height(sheet_xml: str) -> float:
    m = re.search(r'<sheetFormatPr[^>]*?defaultRowHeight="([\d.]+)"', sheet_xml)
    return float(m.group(1)) if m else 13.5


def _pt_to_px(pt: float) -> int:
    """磅 -> 像素 (96dpi)。"""
    return int(round(pt * 96 / 72))


def _col_width_to_px(width: float) -> int:
    """Excel 列宽(字符单位) -> 像素的近似换算。"""
    return int(round(width * 7 + 5))


def calc_anchor_size(img: CellImage, media_data: bytes, sheet_xml: str,
                     cell: DispImgCell, size_mode: str) -> tuple:
    """计算 oneCellAnchor 的 ext 尺寸 (cx, cy) 单位 EMU。

    size_mode:
      - "keep": 保留 cellimages.xml 中的原始尺寸
      - "cell": 高度按所在行高, 宽度按图片真实比例 (贴合单元格, 类似手工修复的效果)
    """
    if size_mode == "keep" and img.ext_cx > 0 and img.ext_cy > 0:
        return img.ext_cx, img.ext_cy

    # 图片真实宽高比 (像素)
    aspect: Optional[float] = None
    real = get_image_size(media_data) if media_data else None
    if real and real[0] > 0 and real[1] > 0:
        aspect = real[0] / real[1]
    elif img.ext_cx > 0 and img.ext_cy > 0:
        aspect = img.ext_cx / img.ext_cy
    if not aspect or aspect <= 0:
        aspect = 1.0

    # 目标行高 (px)
    heights = _parse_row_heights(sheet_xml)
    default_ht = _get_default_row_height(sheet_xml)
    row_pt = heights.get(cell.row, default_ht)
    row_px = _pt_to_px(row_pt)
    if row_px <= 0:
        row_px = 20  # 兜底

    cy = row_px * EMU_PER_PX
    cx = int(round(cy * aspect))
    return cx, cy


# ---------------------------------------------------------------------------
# 主转换入口
# ---------------------------------------------------------------------------

def _next_drawing_number(zf: zipfile.ZipFile) -> int:
    """当前 zip 中已有的 drawing 最大编号 + 1。"""
    max_n = 0
    for n in zf.namelist():
        m = re.match(r"xl/drawings/drawing(\d+)\.xml$", n)
        if m:
            max_n = max(max_n, int(m.group(1)))
    return max_n + 1


def _existing_drawing_rels(zf: zipfile.ZipFile, sheet_file: str):
    """返回 (sheet_rels_xml 或 None, 已有 rIds, 已有 drawing rel Id)。"""
    rels_part = _rels_path(sheet_file)
    names = set(zf.namelist())
    if rels_part not in names:
        return None, set(), None
    xml = zf.read(rels_part).decode("utf-8", "replace")
    rids = set(re.findall(r'Id="(rId\d+)"', xml))
    drawing_rid = None
    m = re.search(r'<Relationship Id="(rId\w+)" Type="[^"]*?/drawing"', xml)
    if m:
        drawing_rid = m.group(1)
    return xml, rids, drawing_rid


def fix_workbook(input_path: str, output_path: str = "", size_mode: str = "cell",
                 keep_wps_part: bool = False, check_only: bool = False) -> FixReport:
    """修复一个 xlsx 文件。返回修复报告。

    size_mode:
      - "cell" (默认): 图片高度按所在行高, 宽度按原图比例。保证图片不遮挡相邻行,
                       与手工修复的效果一致。
      - "keep": 保留 WPS cellimages.xml 中记录的原始尺寸。若表格行高小于图片高度,
                图片会溢出并遮挡其他内容 (忠实但可能不可用)。
    keep_wps_part: 是否保留原 cellimages.xml (默认删除, 与手工修复一致)。
    check_only:    只检测分析, 不写输出文件。
    """
    report = FixReport(input_path=input_path, output_path=output_path)

    try:
        zf = zipfile.ZipFile(input_path)
    except zipfile.BadZipFile:
        report.errors.append(f"不是有效的 xlsx (zip) 文件: {input_path}")
        return report

    with zf:
        names = set(zf.namelist())

        # 1. 检测 WPS cellimage
        ci_part = has_wps_cellimage(zf)
        if not ci_part:
            report.errors.append("未检测到 WPS 单元格内嵌图片 (cellimages.xml), 无需修复。")
            return report
        report.has_wps_cellimage = True
        report.cellimage_part = ci_part

        # 2. 解析 cellimages
        images = parse_cellimages(zf, ci_part)
        report.total_images = len(images)
        img_by_id = {img.img_id: img for img in images}

        # 3. 扫描所有工作表
        sheet_files = sorted(
            n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n)
        )
        all_dispimg: List[DispImgCell] = []
        for sf in sheet_files:
            all_dispimg.extend(scan_sheet_dispimg(zf, sf))
        report.total_dispimg = len(all_dispimg)

        if not all_dispimg:
            report.errors.append(
                f"检测到 cellimages.xml 但未找到 DISPIMG 公式单元格 "
                f"({len(images)} 张图片成为孤儿, 无法确定位置)。"
            )
            return report

        # 4. 按工作表分组
        by_sheet: Dict[str, List[DispImgCell]] = {}
        for c in all_dispimg:
            by_sheet.setdefault(c.sheet_file, []).append(c)

        # 5. 生成 drawing 并清洗 sheet
        drawing_no = _next_drawing_number(zf)
        new_drawings: List[str] = []  # 需要写入 zip 的新 drawing part
        media_refs: Dict[str, str] = {}   # 新 rId -> media 文件 (仅新建 drawing 场景)
        sheet_drawing_rel: Dict[str, str] = {}  # sheet_file -> (rels xml, drawing part)
        sheet_new_rels: Dict[str, str] = {}

        for sheet_file in sorted(by_sheet):
            cells = by_sheet[sheet_file]
            sheet_xml = zf.read(sheet_file).decode("utf-8", "replace")

            anchors = []
            cleared = []
            missing = []
            rel_entries = []   # (rid, REL_IMAGE, target)
            rid_no = 1
            used_rids = set()
            cnvpr_no = 2

            # 判断是否已有 drawing 关系
            existing_rels_xml, existing_rids, drawing_rid = _existing_drawing_rels(zf, sheet_file)
            append_to_existing = drawing_rid is not None

            for c in cells:
                img = img_by_id.get(c.img_id)
                if img is None or not img.media_file:
                    missing.append(c.img_id)
                    continue
                media_part = f"xl/media/{img.media_file}"
                if media_part not in names:
                    media_part = f"xl/{img.media_file}"  # 兼容不同前缀
                if media_part not in names:
                    missing.append(c.img_id)
                    continue

                # 计算尺寸
                media_bytes = zf.read(media_part)
                cx, cy = calc_anchor_size(img, media_bytes, sheet_xml, c, size_mode)

                if append_to_existing:
                    # rId 需要避开已有 drawing rels 中的 id
                    while f"rId{rid_no}" in existing_rids:
                        rid_no += 1
                else:
                    while f"rId{rid_no}" in used_rids:
                        rid_no += 1

                rid = f"rId{rid_no}"
                used_rids.add(rid)
                rid_no += 1

                anchors.append({
                    "col": c.col, "row": c.row,
                    "cx": cx, "cy": cy,
                    "cnvpr_id": cnvpr_no,
                    "name": img.img_id,
                    "descr": img.descr or img.img_id,
                    "r_id": rid,
                })
                cnvpr_no += 1
                rel_entries.append((rid, REL_IMAGE, f"../media/{img.media_file}"))
                cleared.append(c.cell_ref)

            if not anchors:
                # 该 sheet 所有引用都缺图
                report.sheets.append(SheetFixResult(
                    sheet_file=sheet_file, drawing_part="",
                    cleared_cells=[], missing_ids=sorted(set(missing))))
                continue

            drawing_part = f"xl/drawings/drawing{drawing_no}.xml"
            drawing_no += 1

            if append_to_existing:
                # 追加到已有 drawing XML
                existing_drawing = zf.read(drawing_rid_target(zf, sheet_file, drawing_rid))
                new_drawing_xml = _append_anchors(existing_drawing, anchors, rel_entries)
                new_drawings.append((drawing_rid_target(zf, sheet_file, drawing_rid), new_drawing_xml))
                # 合并 rels
                new_rels_xml = _merge_rels(existing_rels_xml, rel_entries)
                sheet_new_rels[sheet_file] = new_rels_xml
            else:
                new_drawing_xml = _build_drawing_xml(anchors)
                new_drawings.append((drawing_part, new_drawing_xml))
                drawing_rels_part = f"xl/drawings/_rels/drawing{drawing_no - 1}.xml.rels"
                new_drawings.append((drawing_rels_part, _build_rels_xml(rel_entries)))
                # sheet rels: 新建或合并
                if existing_rels_xml is None:
                    sheet_new_rels[sheet_file] = _build_rels_xml(
                        [("rIdDrawing", REL_DRAWING, f"../drawings/drawing{drawing_no - 1}.xml")])
                else:
                    sheet_new_rels[sheet_file] = _merge_rels(
                        existing_rels_xml,
                        [("rIdDrawing", REL_DRAWING, f"../drawings/drawing{drawing_no - 1}.xml")])

            # 清洗 sheet XML: 清空 DISPIMG 单元格 + 添加 drawing 引用
            sheet_xml = clean_sheet_xml(sheet_xml, cleared)
            if not append_to_existing:
                # 在 </worksheet> 前插入 <drawing r:id="rIdDrawing"/>
                rid_ref = "rIdDrawing"
                if existing_rels_xml is not None and "rIdDrawing" in existing_rids:
                    # 避免冲突
                    n = 1
                    while f"rIdDrawing{n}" in existing_rids:
                        n += 1
                    rid_ref = f"rIdDrawing{n}"
                    # 修正 sheet rels 中用的 id
                    if sheet_new_rels.get(sheet_file):
                        sheet_new_rels[sheet_file] = sheet_new_rels[sheet_file].replace(
                            'Id="rIdDrawing"', f'Id="{rid_ref}"')
                drawing_tag = f'<drawing r:id="{rid_ref}"/>'
                sheet_xml = re.sub(r"</worksheet>", drawing_tag + "</worksheet>", sheet_xml, count=1)
            new_drawings.append((sheet_file, sheet_xml))
            # sheet 的 rels 也要写回 (新建或更新)
            if sheet_file in sheet_new_rels:
                new_drawings.append((_rels_path(sheet_file), sheet_new_rels[sheet_file]))

            report.sheets.append(SheetFixResult(
                sheet_file=sheet_file, drawing_part=drawing_part,
                anchors=anchors, cleared_cells=cleared, missing_ids=sorted(set(missing))))
            report.fixed += len(anchors)

        report.missing = len({m for s in report.sheets for m in s.missing_ids})

        # 6. 更新 workbook.xml.rels (移除 cellImage 关系)
        new_workbook_rels = None
        wb_rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
        if REL_WPS_CELLIMAGE in wb_rels:
            new_workbook_rels = re.sub(
                r'<Relationship [^>]*Type="[^"]*wps\.cn[^"]*cellImage[^"]*"[^>]*/>', "", wb_rels)
            new_workbook_rels = re.sub(r"\n{3,}", "\n", new_workbook_rels)

        # 7. 更新 [Content_Types].xml
        ct = zf.read("[Content_Types].xml").decode("utf-8", "replace")
        new_ct = ct
        # 移除 cellimages override
        ci_override = re.search(r'<Override PartName="[^"]*cellimages\d*\.xml"[^>]*/>', new_ct)
        if ci_override:
            new_ct = new_ct.replace(ci_override.group(0), "")
        # 添加 drawing override
        drawing_parts = {d for d, _ in new_drawings if re.match(r"xl/drawings/drawing\d+\.xml$", d)}
        for dp in sorted(drawing_parts):
            if f'PartName="/{dp}"' not in new_ct:
                new_ct = new_ct.replace(
                    "</Types>",
                    f'<Override PartName="/{dp}" ContentType="{CT_DRAWING}"/></Types>')

        # 8. 写新 zip
        if not check_only:
            _write_output(zf, output_path, new_drawings, new_workbook_rels, new_ct,
                          remove_parts={ci_part, _rels_path(ci_part)},
                          keep_wps_part=keep_wps_part)
            report.output_path = output_path

        return report


def drawing_rid_target(zf: zipfile.ZipFile, sheet_file: str, rid: str) -> str:
    """根据 sheet rels 中的 drawing 关系 rId 找到 drawing part 路径。"""
    rels_xml = zf.read(_rels_path(sheet_file)).decode("utf-8", "replace")
    m = re.search(
        rf'<Relationship Id="{re.escape(rid)}" Type="[^"]*?/drawing" Target="([^"]+)"', rels_xml)
    if not m:
        raise ValueError(f"找不到 drawing 关系 {rid} in {sheet_file}")
    target = m.group(1)
    # target 相对于 xl/worksheets/ 目录, 规范化 (处理 ../ 与 / 前缀)
    if target.startswith("/"):
        target = target.lstrip("/")
    else:
        target = f"xl/worksheets/{target}"
    # 规范化 xl/worksheets/../drawings/drawing1.xml -> xl/drawings/drawing1.xml
    parts = []
    for seg in target.split("/"):
        if seg == "..":
            if parts:
                parts.pop()
        elif seg in ("", "."):
            continue
        else:
            parts.append(seg)
    return "/".join(parts)


def _append_anchors(drawing_xml: str, anchors: List[dict], rel_entries: List[tuple]) -> str:
    """把新 anchor 追加到已有 drawing XML, 并保证命名空间完整。"""
    root = ET.fromstring(drawing_xml)
    wsdr_ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    new_anchors = ET.fromstring(_build_drawing_xml(anchors))
    for child in list(new_anchors):
        root.append(child)
    # 序列化 (保留命名空间前缀 xdr/a/r)
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + ET.tostring(root, encoding="unicode")


def _merge_rels(existing_rels_xml: Optional[str], new_rels: List[tuple]) -> str:
    """合并新关系到已有 rels XML。"""
    rns = "http://schemas.openxmlformats.org/package/2006/relationships"
    if existing_rels_xml is None:
        return _build_rels_xml(new_rels)
    root = ET.fromstring(existing_rels_xml)
    existing_ids = {rel.get("Id") for rel in root.findall(f"{{{rns}}}Relationship")}
    for rid, rtype, target in new_rels:
        if rid in existing_ids:
            continue
        rel = ET.SubElement(root, f"{{{rns}}}Relationship")
        rel.set("Id", rid)
        rel.set("Type", rtype)
        rel.set("Target", target)
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + ET.tostring(root, encoding="unicode")


def _write_output(zf: zipfile.ZipFile, output_path: str,
                  modified: List[tuple],
                  new_workbook_rels: Optional[str],
                  new_ct: str,
                  remove_parts: set,
                  keep_wps_part: bool = False) -> None:
    """把修改后的内容写回新的 zip 文件。"""
    modified_map = dict(modified)
    if new_workbook_rels is not None:
        modified_map["xl/_rels/workbook.xml.rels"] = new_workbook_rels
    modified_map["[Content_Types].xml"] = new_ct

    with zipfile.ZipFile(output_path, "w", allowZip64=True) as zout:
        for item in zf.infolist():
            name = item.filename
            if name in remove_parts and not keep_wps_part:
                continue
            if name in modified_map:
                data = modified_map[name].encode("utf-8")
                zout.writestr(item, data)
            else:
                zout.writestr(item, zf.read(name))
        # 新条目 (如新 drawing part)
        new_names = set(modified_map) - set(zf.namelist())
        for name in sorted(new_names):
            info = zipfile.ZipInfo(name)
            info.compress_type = zipfile.ZIP_DEFLATED
            zout.writestr(info, modified_map[name].encode("utf-8"))
