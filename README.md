# wps-img-fixer

> 🟢 **直接使用**：在线打开 👉 **[wps_img_fixer.html（点击运行）](https://gtx950l.github.io/wps-img-fixer/wps_img_fixer.html)** —— 拖入文件即可修复，**无需安装、无需联网、数据不出本机**。
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

## 怎么使用

### 在线使用（推荐）

打开 **[wps_img_fixer.html](https://gtx950l.github.io/wps-img-fixer/wps_img_fixer.html)**，然后：

1. **拖入 .xlsx 文件**（或点击左侧选择区域，支持一次拖多个文件批量修复）
2. （可选）确认右侧选项：**图片尺寸模式**默认 `cell`，一般不用改
3. 点 **「开始修复」** —— 右侧会显示每个工作表的转换明细
4. 点 **「下载修复后的文件」** —— 得到 `xxx_fixed.xlsx`，用 Excel 打开即可看到图片

> 💡 想先看看效果？点左侧 **「载入内置演示文件」**，再点「开始修复」，即可体验 4 张示例图片的完整修复流程（演示文件是代码生成的纯色图，不含任何真实数据）。

### 离线使用

1. 下载 [wps_img_fixer.html](https://github.com/GTX950L/wps-img-fixer/raw/main/wps_img_fixer.html)（右键另存为）
2. 双击用浏览器打开，操作同上 —— 单文件自包含（JSZip 内嵌），**断网也能用**

### 选项说明

| 选项 | 说明 |
| --- | --- |
| 图片尺寸模式 · `cell` | **默认**。图片高度按所在行高、宽度按原图比例——贴合单元格、不遮挡相邻行 |
| 图片尺寸模式 · `keep` | 保留 WPS 记录的原始尺寸（最忠实，但表格行高偏小时图片会溢出遮挡） |
| 保留 cellimages.xml | 一般不勾。默认删除 WPS 私有部件，避免 Excel 误读 |

### 验证一下

修复后用 Excel 打开确认图片正常显示、位置与原来一致即可。整个检测与转换都在浏览器本地完成，**文件不会上传到任何服务器**。

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
├── wps_img_fixer.html     # ⭐ 唯一入口：单文件 HTML 工具（JSZip 内嵌，完全离线）
├── .nojekyll              # GitHub Pages 部署（避免 Jekyll 过滤）
├── LICENSE                # MIT 许可
└── README.md
```

## 已知限制

- 只处理"WPS 单元格内嵌图片"。普通浮动图片、图表等不受影响。
- 若某个 `=DISPIMG(...)` 引用的图片 ID 在 `cellimages.xml` 中不存在（图片数据缺失），该单元格会**原样保留**并在报告中提示，避免误删信息。
- 转换后的图片是"锚定单元格"的标准浮动图片，默认不随行高列宽变化（与 WPS 嵌入图片行为略有差异，可在 Excel 中自行调整大小）。

## 许可证

[MIT](LICENSE)
