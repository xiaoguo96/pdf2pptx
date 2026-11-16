# PDF2PPTX — 图片型 PDF 转可编辑 PPTX

将**图片型 / 扫描型 PDF**（每页是一张整页位图、无文字层）转换为**可编辑 PPTX**：
文字还原为原生文本框（可选中、可改写），背景图片作为独立图片对象（可移动、可缩放），
幻灯片尺寸与 PDF 页面一致，视觉布局尽量保持原样。

## 转换原理

```
PDF 页面 ──PyMuPDF 渲染──> 高分辨率位图 (150 DPI)
        ──RapidOCR──> 每行文字 + 边界框 + 置信度
        ──元素分类──> PROTECT(保持原样) / REBUILD(重建为文本框)
        ──像素分析──> 字号 / 颜色 / 加粗（由笔画宽度估算）
        ──OpenCV 平滑填充──> 仅擦除 REBUILD 行文字（逐行线性插值，无涂抹阴影）
        ──python-pptx──> 幻灯片 = 处理后的整页背景图 + 原生文本框叠加
```

1. **元素保护（PROTECT）**：以下 OCR 行**不擦除、不重渲染，保持原始像素**——
   页眉/页脚带（logo、公司名、页码）、纯数字/百分比/千分位/年份标签（图表数值、
   目录页码、徽章编号）、极短小标签、紧邻高饱和彩色图形（柱状图、图标）或背景
   不平滑（照片/渐变/纹理）的文字。OCR 合并行（如 `01关于柯幻03`）按字宽比例
   拆分：数字段保持原样，仅中间文字段重建为可编辑文本框。
2. **文字重建（REBUILD）**：平滑背景上的正文按边界框生成独立原生文本框；字号取
   “高度反推”与“宽度反推”（Σ字宽系数）的**较小值**，保证重排文字不会溢出原
   矩形框/卡片；颜色取字形核心像素中位数，加粗由距离变换估算笔画宽度比判定；
   关闭自动换行、零内边距、垂直居中，位置与原始像素坐标一一对应。
3. **背景修复**：仅对 REBUILD 行做逐行线性插值填充（从文字左右最近的背景像素
   插值），替代 TELEA inpaint，消除渐变/纯色背景上的涂抹阴影。
4. **图片**：每页背景图作为单个 Picture 对象铺满幻灯片，可整体选中、移动、缩放、替换。
5. **图形/线条/表格/图表**：见下方“还原情况与限制”。

## 依赖环境

- Python 3.9+（开发验证于 3.10, Windows；Linux/macOS 同样可用）
- 依赖清单见 [requirements.txt](requirements.txt)：
  `flask` `pymupdf` `opencv-python-headless` `numpy` `pillow` `python-pptx` `rapidocr-onnxruntime`
- OCR 模型（PaddleOCR 中英文检测+识别 ONNX 模型）随 `rapidocr-onnxruntime` 自动打包，
  首次运行自动加载，**无需联网、无需 GPU**。

## 安装与启动

```bash
git clone <本仓库地址>
cd pdf2pptx
python -m venv .venv
.venv\Scripts\activate            # Linux/macOS: source .venv/bin/activate
pip install -r requirements.txt
```

### 方式一：命令行转换

```bash
python scripts/convert.py 输入.pdf 输出.pptx --dpi 150
# 同时生成元素清单 输出.manifest.json（每页文字/坐标/字号/颜色/加粗）
```

### 方式二：Web Demo（上传 PDF → 下载 PPTX）

```bash
python app.py                     # 默认 http://127.0.0.1:7860
# 自定义端口: PORT=8000 python app.py
```

打开页面 → 选择/拖拽 PDF → 点击“开始转换” → 自动下载 `xxx_editable.pptx`，
状态栏显示 `slides=N;texts=M` 统计。

### 方式三：Docker（云端部署即得在线 Demo）

```bash
docker build -t pdf2pptx .
docker run -p 7860:7860 pdf2pptx
```

可一键部署到 Render / Railway / Hugging Face Spaces / 任意支持 Docker 的云平台
（Dockerfile 已按 `PORT` 环境变量监听）。

## 元素还原情况与无法还原的内容

| 元素类型 | 还原结果 | 可编辑性 |
|---|---|---|
| 正文文字（平滑背景） | ✅ 原生文本框，位置/字号/颜色/加粗自动估算，字号受宽度约束不溢出 | 可选中、改写、改字体字号颜色 |
| 页眉/页脚（logo、公司名、页码） | ✅ **原样保留**（PROTECT，不擦除不重渲染） | 保持背景图像素级原样 |
| 图表数值/百分比/年份标签、目录页码、徽章编号 | ✅ **原样保留**（PROTECT） | 保持背景图像素级原样 |
| 贴彩色图形的文字、照片/渐变上的文字 | ✅ **原样保留**（PROTECT，避免破坏与阴影） | 保持背景图像素级原样 |
| 整页背景（含插画、照片、装饰图形、图表本体） | ✅ 独立 Picture 对象（仅重建行被平滑擦除） | 可整体选中、移动、缩放、替换 |
| 装饰线条、图标、艺术字 | ⚠️ 保留在背景图中 | 不可单独编辑 |
| 表格 | ⚠️ 本样例无表格；通用管线未做表格线框重建 | 单元格文字可编辑，框线不可 |
| 竖排/弯曲/极小文字 | ❌ OCR 可能漏检，原文字保留在背景中 | 不可编辑 |

**无法完全还原的根因**：源 PDF 每页是单一位图（`get_drawings()` 为 0、无文字层），
矢量信息在生成时已丢失；程序化重建任意图形/图表/表格的语义结构不可靠，
因此采取“背景保真 + 文字原生可编辑”策略，并把可编辑文字精确叠加回原位。

## 可编辑性验证材料（verify/ + samples/）

- [verify/verify_pptx.py](verify/verify_pptx.py)：打开 PPTX 逐页统计对象，并对每个文本框做
  **“改写探针”**（写入再恢复，证明可编辑）。
- 样例输出（本仓库 `samples/`）：
  - `verify_result.verify.md` / `.json` — 10 页、70 个可编辑文本框，探针通过 70/70（100%）；
    另有 91 个受保护行（页码/logo/图表数值等）保持原样
  - `quality_check.json` — 原页 vs 转换渲染逐页像素对比：非文字区域 98.1% 像素偏差 < 12
  - `output_manifest.json` — 每页元素清单，含 `texts`（重建）与 `protected`（原样保留+原因）
  - `pptx_render/slide01-10.png` — PowerPoint COM 实际导出（证明 PowerPoint 可正常打开）
  - `compare/cmp01-10.jpg` — 左原页 / 右转换结果 并排对比图
  - `demo_screenshots/` — 浏览器实测 Demo 上传→转换→下载全流程截图
- 兼容性：输出文件由 PowerPoint COM 成功打开并导出；python-pptx 可正常读写；
  WPS Office 可直接打开（OOXML 标准格式）。

## 实测结果（测试3.pdf，10 页图片型 PDF）

- 10 页 → 10 张幻灯片（852×1212 pt，与 PDF 页面一致）
- 70 个可编辑文本框 + 91 个受保护元素行（页眉/页脚/页码/图表数值/徽章等原样保留）
- 文本改写探针 70/70 通过
- 背景保真度（非文字区域像素匹配率）平均 98.1%；受保护区域逐页偏差 ≤2（压缩噪声级）
- 重渲染文字宽度逐行核验：与原行宽偏差 ≤1px，无矩形框/卡片溢出
- Web Demo 端到端实测：上传 → 约 60s → 自动下载 PPTX

## 目录结构

```
pdf2pptx/
├── app.py                  # Flask Web Demo
├── templates/index.html    # 上传页面
├── converter/core.py       # 转换管线（渲染/OCR/排版估算/背景修复/构建 PPTX）
├── scripts/convert.py      # 命令行入口
├── verify/verify_pptx.py   # 可编辑性验证脚本
├── samples/                # 实测产物与验证材料
├── requirements.txt
└── Dockerfile
```

## 已知限制

- 仅适用于图片型 PDF；带文字层的 PDF 建议直接用 PowerPoint 导入或 pdf2pptx 类工具。
- 字号/加粗为启发式估算，艺术字体可能与原字体观感有差异（默认 Microsoft YaHei）。
- 受保护行（页码、图表数值、贴图形文字等）保持原样但**不可编辑**——这是刻意的
  保真取舍，避免擦除/重渲染破坏 logo、页码、图表与矩形卡片。
- 复杂纹理背景上的文字不再擦除（改判保护），因此不会残留涂抹阴影。
- 在线服务为同步转换，大文件（>20 页）建议改用任务队列。
