"""wps-dispimg-fixer: 修复 WPS 单元格内嵌图片 (DISPIMG) 在 Excel 中不显示的问题。

WPS 的"单元格内嵌图片"是其私有扩展:
- 单元格写入公式 =DISPIMG("ID_xxx", 1)
- 图片本体存在 xl/cellimages.xml (WPS 私有命名空间)
微软 Excel 不认识这两样, 因此打开后图片消失, 只剩一串代码。

本工具把 WPS 私有机制转换为标准 OOXML drawing 机制,
让图片在 Excel / WPS 中都能正常显示。
"""

__version__ = "0.1.0"
