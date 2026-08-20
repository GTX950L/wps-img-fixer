# wps-img-fixer

> 🟢 **直接使用**：在线打开 👉 **[wps_img_fixer.html（点击运行）](https://gtx950l.github.io/wps-img-fixer/wps_img_fixer.html)** —— 拖入文件即可修复，**无需安装、无需联网、数据不出本机**。
>
> 也有 [Python 命令行版](#python-命令行版) 可批量处理。
>
> 🔗 项目仓库：[github.com/GTX950L/wps-img-fixer](https://github.com/GTX950L/wps-img-fixer)

修复 **WPS 单元格内嵌图片在 Excel 中不显示、只剩一串代码** 的问题。

## 问题是什么

WPS 表格里把图片"嵌入单元格"（图片随单元格移动/排序）时，保存的文件用的是 **WPS 私有扩展**：

- 单元格里写入公式 `=DISPIMG("ID_xxx", 1)`
- 图片本体存放在 WPS 私有的 `xl/cellimages.xml`（微软 Excel 不认识这个文件）

于是用微软 Excel（或部分其他软件）打开这个文件时：

| 在 WPS 中 | 在 Excel 中 |
| --- | --- |
| 图片正常显示 | 图片消失，单元格里只剩一串代码 |
| | `=_xlfn.DISPIMG("ID_62B52A37D6BA44EC970C96E11457D488",1)` |

> 注：WPS 的普通"浮动图片"（插入 → 图片）不受影响，问题只出现在**嵌入单元格**的图片。

## 本工具做什么

把 WPS 私有机制**转换**为标准的 OOXML 图片（`xl/drawings/drawingN.xml`），转换后：

- ✅ 图片在 Excel / WPS / 手机版中都能正常显示
- ✅ 图片仍锚定在原单元格位置（跟随行列移动）
- ✅ 单元格里的 `=DISPIMG(...)` 公式被清除（保留原单元格样式）
- ✅ 表格数据、公式、格式完全不变

转换前（Excel 打开，图片全丢）：

```
H2: =_xlfn.DISPIMG("ID_62B52A37D6BA44EC970C96E11457D488",1)   ← 一串代码
```

转换后（Excel 打开，图片正常显示，公式已清空）：

```
H2: (空, 图片显示在 H 列单元格位置)
```

## 两种使用方式

### 1. 在线 HTML 版（推荐）

直接打开 **[wps_img_fixer.html](https://gtx950l.github.io/wps-img-fixer/wps_img_fixer.html)**：

- **纯单文件 HTML**，JSZip 内嵌，**完全离线可用**（断网也行）
- 拖入或选择 .xlsx → 自动检测 WPS DISPIMG → 一键修复 → 下载结果
- 全部在浏览器内完成，**文件不上传任何服务器**
- 内置演示文件可一键体验
- 也可下载 `wps_img_fixer.html` 到本地双击用浏览器打开

页面支持两个选项：

| 选项 | 说明 |
| --- | --- |
| 图片尺寸模式 · `keep` | **默认**。保留 WPS 中图片的原始显示尺寸（最忠实） |
| 图片尺寸模式 · `cell` | 按所在单元格行高缩放，更贴合表格（适合行高已设置好的表） |
| 保留 cellimages.xml | 一般不勾——保留 WPS 私有部件可能让 Excel 误读 |

### 2. Python 命令行版

环境要求：Python 3.8+，**零第三方依赖**（只用标准库）。

```bash
# 修复单个文件
python -m wps_dispimg_fixer 报告.xlsx

# 批量修复文件夹下所有 xlsx
python -m wps_dispimg_fixer 数据文件夹/

# 只检测, 不生成文件
python -m wps_dispimg_fixer 报告.xlsx --check
```

输出示例：

```
文件: C:/Users/1/Desktop/FAI 116_0819D.xlsx
  检测: 包含 WPS 单元格内嵌图片 (cellimages.xml), 共 23 张
  工作表 sheet1.xml:
    转换图片: 23 个
    清除公式: 23 个单元格
  完成: 转换 23 张图片, 清除 23 个公式, 缺失 0 个
  输出: C:/Users/1/Desktop/FAI 116_0819D_fixed.xlsx
```

### 参数

| 参数 | 说明 |
| --- | --- |
| `--check` | 只检测并报告，不写文件 |
| `--size keep` | **默认**。保留 WPS 中图片的原始显示尺寸（最忠实） |
| `--size cell` | 按所在单元格行高缩放图片，更贴合表格（适合行高已设置好的表） |
| `--keep-wps-part` | 保留原 `cellimages.xml` 文件（默认删除，避免 Excel 误读） |

### 验证一下

修复后建议用 Excel 打开确认图片正常。也可以用 openpyxl 快速检查：

```python
from openpyxl import load_workbook
wb = load_workbook("修复后.xlsx")
ws = wb["Sheet1"]
print(len(ws._images), "张图片")   # 修复前是 0, 修复后是图片数
```

## 原理（简版）

OOXML (xlsx) 本质是一个 zip 包，WPS 的"单元格内嵌图片"和标准图片分别长这样：

```
标准 Excel 图片                           WPS 单元格内嵌图片
─────────────────────                    ─────────────────────
xl/worksheets/sheet1.xml                 xl/worksheets/sheet1.xml
  └─ <drawing r:id="rIdDrawing"/>          └─ 单元格: =DISPIMG("ID_xxx",1)
xl/drawings/drawing1.xml                 xl/cellimages.xml   ← WPS 私有
xl/media/image1.png                      xl/media/image1.png
```

微软 Excel 不识别 `cellimages.xml` 和 `DISPIMG` 函数，所以图片丢失。

本工具的转换流程：

1. 解析 `xl/cellimages.xml`，得到每张图片的 ID、描述、尺寸、对应 media 文件
2. 扫描每个工作表，找出所有 `=DISPIMG("ID_xxx",1)` 公式单元格（同时拿到目标行列）
3. 为每个图片生成标准 `xl/drawings/drawingN.xml`（`oneCellAnchor` 锚定在原单元格）
4. 清空单元格里的 DISPIMG 公式（保留位置与样式），写入 `<drawing>` 引用
5. 更新 `[Content_Types].xml`、workbook/sheet 关系文件，删除 WPS 私有部件
6. 重新打包输出

## 项目结构

```
wps-img-fixer/
├── wps_img_fixer.html           # ⭐ 浏览器版（单文件，JSZip 内嵌）
├── .nojekyll                    # GitHub Pages 部署 (避免 Jekyll 过滤)
├── wps_dispimg_fixer/           # Python 版主包
│   ├── cli.py                   # 命令行入口
│   ├── converter.py             # 核心转换逻辑
│   └── image_utils.py           # 图片尺寸解析 (JPEG/PNG/GIF/BMP)
├── tests/
│   ├── sample_factory.py        # 测试样本生成 (纯标准库, 可公开)
│   └── test_converter.py        # 单元测试 (15 个用例)
├── scripts/
│   └── make_test_sample.py      # 生成演示用测试样本
└── README.md
```

## 开发与测试

```bash
python -m unittest discover -s tests -v    # 跑单元测试
python scripts/make_test_sample.py          # 生成演示样本 (demo_wps_dispimg.xlsx)
```

测试样本由代码生成，不含任何真实业务数据，可放心用于演示与二次开发。

## 已知限制

- 只处理"WPS 单元格内嵌图片"。普通浮动图片、图表等不受影响。
- 若某个 `=DISPIMG(...)` 引用的图片 ID 在 `cellimages.xml` 中不存在（图片数据缺失），该单元格会**原样保留**并在报告中提示，避免误删信息。
- 转换后的图片是"锚定单元格"的标准浮动图片，默认不随行高列宽变化（与 WPS 嵌入图片行为略有差异，可在 Excel 中自行调整大小）。

## 许可证

[MIT](LICENSE)
