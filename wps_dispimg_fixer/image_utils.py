"""图片尺寸解析工具: 从图片二进制中读取宽高(像素)。

支持 JPEG / PNG / GIF / BMP, 全部使用标准库实现。
用于"按单元格缩放"模式时按图片真实比例计算显示尺寸。
"""

from __future__ import annotations

import struct
from typing import Optional, Tuple


def get_image_size(data: bytes) -> Optional[Tuple[int, int]]:
    """读取图片像素尺寸 (width, height)。

    支持 JPEG/PNG/GIF/BMP, 无法识别时返回 None。
    """
    if not data:
        return None
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return _png_size(data)
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return _gif_size(data)
    if data[:2] == b"BM":
        return _bmp_size(data)
    if data[:2] == b"\xff\xd8":
        return _jpeg_size(data)
    return None


def _png_size(data: bytes) -> Optional[Tuple[int, int]]:
    # PNG: 签名8字节后是 IHDR 块: 长度(4) + 'IHDR'(4) + 宽(4BE) + 高(4BE)
    if len(data) < 24:
        return None
    w, h = struct.unpack(">II", data[16:24])
    return (w, h) if w and h else None


def _gif_size(data: bytes) -> Optional[Tuple[int, int]]:
    # GIF: 头6字节 + 逻辑屏幕宽度(2LE) + 高度(2LE)
    if len(data) < 10:
        return None
    w, h = struct.unpack("<HH", data[6:10])
    return (w, h) if w and h else None


def _bmp_size(data: bytes) -> Optional[Tuple[int, int]]:
    # BMP: 偏移18处为宽度(4LE), 偏移22处为高度(4LE, 可为负)
    if len(data) < 26:
        return None
    w = struct.unpack("<i", data[18:22])[0]
    h = struct.unpack("<i", data[22:26])[0]
    if w > 0 and h != 0:
        return (w, abs(h))
    return None


def _jpeg_size(data: bytes) -> Optional[Tuple[int, int]]:
    """遍历 JPEG marker 找到 SOF (Start Of Frame) 读取尺寸。

    SOF marker: 0xC0-0xC3, 0xC5-0xC7, 0xC9-0xCB, 0xCD-0xCF
    (0xC4/0xC8/0xCC 是 DHT/JPG/DAC, 不含尺寸信息)
    """
    i = 2  # 跳过 \xff\xd8
    n = len(data)
    while i + 9 < n:
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        # 跳过填充字节
        while marker == 0xFF and i + 2 < n:
            i += 1
            marker = data[i + 1]
        if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
            # 无长度字段的 marker
            i += 2
            continue
        if i + 4 > n:
            return None
        seg_len = struct.unpack(">H", data[i + 2:i + 4])[0]
        if seg_len < 2:
            return None
        if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
            if i + 9 > n:
                return None
            h = struct.unpack(">H", data[i + 5:i + 7])[0]
            w = struct.unpack(">H", data[i + 7:i + 9])[0]
            return (w, h) if w and h else None
        i += 2 + seg_len
    return None
