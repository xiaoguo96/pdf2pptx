# PDF2PPTX — 图片型 PDF 转可编辑 PPTX

将**图片型 / 扫描型 PDF**（每页是一张整页位图、无文字层）转换为**可编辑 PPTX**：
页面上的文字（正文、标题、页眉页脚、页码、目录编号、图表数值标签等）**几乎全部**
还原为原生文本框（可选中、可改写、可改字体字号颜色），背景图片作为独立图片对象
（可移动、可缩放、可替换），幻灯片尺寸与 PDF 页面一致，视觉布局尽量保持原样。

> **在线 Demo（可直接操作：上传 PDF → 下载可编辑 PPTX）**
> https://server-edmonton-encourages-translated.trycloudflare.com
>
> 公网实例运行**离线 OCR 模式**（零 API 消耗，实测 165 个可编辑文本框）；
> 视觉大模型（LLM）模式需本地配置 `LLM_API_KEY` 后按下方"方式二"运行。
> 该地址为 Cloudflare 临时隧道，演示机开机时有效；长期部署请用仓库内
> Dockerfile 一键部署（见"方式三"）。

## 转换原理

```
PDF 页面 ──PyMuPDF 渲染──> 高分辨率位图 (150 DPI)
        ──提取(双模式)──> A. 视觉大模型(OpenAI 兼容 API)：整页图 → 文字块+bbox+颜色+语义类别
        │                B. 本地 RapidOCR：逐行文字 + 边界框（无 API key 时自动降级）
        ──文字校验──>   LLM 读文 × 本地 OCR 交叉仲裁；争议行放大裁剪复读；
        │               页码一律放大复读取真值（防 01/1 误读、防补零幻觉）
        ──像素分析──>   字号/字距用真实字体度量拟合（PIL + Microsoft YaHei）；
        │               颜色取字形核心像素中位数；加粗由距离变换笔画宽度比判定
        ──文字擦除──>   平滑背景 → 逐行线性插值填充；复杂/彩色背景 → 仅字形像素
        │               紧致掩码 TELEA 修复（不伤折线/柱形/徽章图形）
        └─python-pptx──> 幻灯片 = 修复后的整页背景图 + 逐行原生文本框叠加
```

### 双提取模式（`--mode auto|llm|ocr`）

| 模式 | 提取方式 | 说明 |
|---|---|---|
| `llm` | 视觉大模型读整页图，输出文字块、bbox、颜色、加粗及语义类别 | 需 `LLM_API_KEY`；本地 OCR 逐行交叉校验，三方仲裁防误读 |
| `ocr` | RapidOCR 逐行识别 | 完全离线，无 API 消耗 |
| `auto` | 有 key 走 llm，否则 ocr；单页 API 失败自动回退 ocr | 默认 |

### 重建与保护策略

1. **文字重建（默认）**：页面上几乎所有文字都重建为独立原生文本框——包括页码、
   页眉/页脚公司名与标语、目录条目编号、封面标题、图表数值标签。
   - **字号/字距拟合**：字号由字形高度反推（紧框 ≈ 0.86 em），再用真实字体度量
     校验渲染宽度：偏窄则加正字距撑满原宽度，偏宽则等比缩小字号，保证文字
     **不溢出原矩形、不遮挡相邻内容**；
   - **防幻觉保护**：声称区域内实测无墨迹或墨迹过稀疏的元素不重建（保持原样），
     避免模型幻觉文字覆盖真实内容；
   - **图表图形保护**：模型判为 chart_label/badge 且紧贴复杂图形的元素保持原样，
     擦除只作用于字形像素本身（TELEA 紧致掩码），折线、柱形、徽章图形无损。
2. **背景擦除双策略**：平滑背景逐行线性插值（无涂抹阴影）；彩色/渐变/贴图形
   背景仅擦字形像素并 TELEA 修复，修复范围限制在掩码膨胀 2px 内。
3. **文字正确性三方仲裁**：LLM 整页读文 × 本地 OCR 宽松裁剪识别交叉比对，
   长度不一致时用墨迹连通域字形数裁决，仍有争议时对该行放大裁剪复读；
   页码类元素一律放大复读。
4. **图片**：每页背景图作为单个 Picture 对象铺满幻灯片，可整体选中、移动、缩放、替换。

## 依赖环境

- Python 3.9+（开发验证于 3.10, Windows；Linux/macOS 同样可用）
- 依赖清单见 [requirements.txt](requirements.txt)：
  `flask` `pymupdf` `opencv-python-headless` `numpy` `pillow` `python-pptx` `rapidocr-onnxruntime`
- OCR 模型（PaddleOCR 中英文检测+识别 ONNX 模型）随 `rapidocr-onnxruntime` 自动打包，
  首次运行自动加载，**无需联网、无需 GPU**。
- LLM 模式仅用 Python 标准库（`urllib`）调用 OpenAI 兼容接口，无额外依赖。

## 主要依赖及 AI / 第三方 API 使用情况说明

按题目统一说明要求，列明本项目使用的外部资源：

| 类别 | 名称 | 用途 |
|---|---|---|
| AI 编程工具 | Trae（AI IDE） | 辅助编写本项目源代码、调试与测试脚本 |
| 开源项目 | PyMuPDF | PDF 页面渲染 |
| 开源项目 | RapidOCR (PaddleOCR ONNX 模型) | 本地 OCR：离线提取 + LLM 读文交叉校验 |
| 开源项目 | OpenCV | 像素分析、文字擦除（插值填充 / TELEA 修复） |
| 开源项目 | Pillow | 真实字体度量（字号/字距拟合） |
| 开源项目 | python-pptx | 生成 PPTX 原生对象 |
| 开源项目 | Flask | Web Demo |
| 第三方 API | OpenAI 兼容中转接口（面试官提供） | LLM 模式：视觉大模型整页布局提取与争议行复读 |

**敏感信息处理**：API Key 仅通过环境变量 `LLM_API_KEY` 注入，
**未提交到本仓库任何文件**（仓库内代码只读取 `os.environ`）。

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
python scripts/convert.py 输入.pdf 输出.pptx --dpi 150 --mode auto
# 同时生成元素清单 输出.manifest.json（每页文字/坐标/字号/字距/颜色/加粗/擦除方式/保护原因）
# LLM 模式（需面试官提供的 key，切勿写死在代码里）:
#   export LLM_API_KEY=sk-xxxx   # Windows: set LLM_API_KEY=sk-xxxx
#   python scripts/convert.py 输入.pdf 输出.pptx --mode llm
```

### 方式二：Web Demo（上传 PDF → 下载 PPTX）

```bash
export LLM_API_KEY=sk-xxxx       # 可选；不设置则 Demo 自动使用本地 OCR 模式
python app.py                    # 默认 http://127.0.0.1:7860
# 自定义端口: PORT=8000 python app.py
```

打开页面 → 选择/拖拽 PDF → 点击"开始转换" → 自动下载 `xxx_editable.pptx`，
状态栏显示 `slides=N;texts=M` 统计。

### 方式三：Docker（云端部署即得在线 Demo）

```bash
docker build -t pdf2pptx .
docker run -p 7860:7860 pdf2pptx
```

可一键部署到 Render / Railway / Hugging Face Spaces / 任意支持 Docker 的云平台
（Dockerfile 已按 `PORT` 环境变量监听）。

## 元素还原情况（类型 / 还原结果 / 可编辑性）

以 `测试3.pdf`（10 页图片型 PDF）LLM 模式实测为准：

| 元素类型 | 还原结果 | 可编辑性 |
|---|---|---|
| 正文 / 标题 / 多行段落 | ✅ 原生文本框，位置/字号/颜色/加粗自动还原，字距撑满原宽度 | 可选中、改写、改字体字号颜色 |
| 页码（01、02、3、4、5、6…） | ✅ 原生文本框（放大复读取数字真值） | 完全可编辑 |
| 页眉 / 页脚（公司名、标语） | ✅ 原生文本框 | 完全可编辑 |
| 目录条目编号与页码 | ✅ 原生文本框 | 完全可编辑 |
| 封面主标题 / 副标题 | ✅ 原生文本框（字形紧致擦除，封面插画无损） | 完全可编辑 |
| 图表数值标签（数据点、轴刻度、百分比） | ✅ 原生文本框；擦除仅限字形像素，折线/柱形无损（实测彩色图形像素损伤 0.08%） | 完全可编辑 |
| 图表本体（折线、柱形、坐标轴、图例图形）、图标、装饰图形 | ⚠️ 原样保留在背景图中 | 不可单独编辑 |
| 无墨迹 / 极小文字（防幻觉保护） | 原样保留在背景图中 | 不可编辑 |
| 整页背景（含插画、照片、渐变） | ✅ 独立 Picture 对象（仅文字行被擦除修复） | 可整体选中、移动、缩放、替换 |

**无法完全还原的根因**：源 PDF 每页是单一位图（`get_drawings()` 为 0、无文字层），
矢量信息在生成时已丢失；程序化重建任意图形/图表/表格的语义结构不可靠，
因此采取"背景保真 + 文字全部原生可编辑"策略，把可编辑文字精确叠加回原位。

## 实测结果（测试3.pdf，10 页图片型 PDF）

**LLM 模式（主交付，视觉大模型 gpt-5.6-luna）**

- 10 页 → 10 张幻灯片（852×1212 pt，与 PDF 页面一致），全部页面由 API 提取
- **169 个原生可编辑文本框**（含页码 8 个、页眉页脚、目录编号、封面标题、
  图表数值标签），0 个保护元素
- 文本改写探针 **169/169 通过**（逐框写入再恢复，实证可编辑）
- 黑块检测 **0 处**；残留重影检测 **0 处**（原墨迹像素未改变比例 >35% 判重影）
- 页码真值核验：第 2–9 页 01 / 02 / 01 / 02 / 3 / 4 / 5 / 6 **全部正确**
- 折线图完整性：p8 图表绘图区彩色图形像素损伤 **0.08%**（折线无损）
- 背景保真度：非文字区域像素偏差 <12 的比例平均 **96.1%**
- PowerPoint COM 成功打开并逐页导出（`samples/pptx_render/`）

**OCR 模式（离线降级，无 API 消耗）**

- 165 个可编辑文本框 + 1 个保护元素；改写探针 165/165 通过

## 可编辑性验证材料（verify/ + samples/）

- [verify/verify_pptx.py](verify/verify_pptx.py)：打开 PPTX 逐页统计对象，并对每个文本框做
  **"改写探针"**（写入再恢复，证明可编辑）。
- 样例输出（本仓库 `samples/`，均为上述最终实测产物）：
  - `结果说明.md` — 元素还原情况 / 验证结果 / 无法还原内容的完整说明
  - `output_测试3_editable.pptx` — LLM 模式转换结果（169 文本框）
  - `output_manifest.json` — 每页元素清单：文字/坐标/字号/字距/颜色/加粗/擦除方式
  - `verify_result.verify.md` / `.json` — 探针报告 169/169
  - `quality_check.json` — 原页 vs PowerPoint 渲染逐页像素对比（背景匹配率 96.1%）
  - `pptx_render/slide01-10.jpg` — PowerPoint COM 实际导出（证明可正常打开）
  - `compare/cmp01-10.jpg` — 左原页 / 右转换结果 并排对比图
  - `demo_screenshots/` — 浏览器实测 Demo 上传→转换→下载全流程截图
- 兼容性：输出文件由 PowerPoint COM 成功打开并导出；python-pptx 可正常读写；
  WPS Office 可直接打开（OOXML 标准格式）。

## 目录结构

```
pdf2pptx/
├── app.py                  # Flask Web Demo
├── templates/index.html    # 上传页面
├── converter/core.py       # 转换管线（渲染/提取/交叉校验/度量拟合/背景修复/构建 PPTX）
├── converter/llm.py        # OpenAI 兼容视觉 API 布局提取器（key 仅从环境变量读取）
├── scripts/convert.py      # 命令行入口
├── verify/verify_pptx.py   # 可编辑性验证脚本
├── samples/                # 实测产物与验证材料
├── requirements.txt
└── Dockerfile
```

## 已知限制

- 仅适用于图片型 PDF；带文字层的 PDF 建议直接用 PowerPoint 导入或 pdf2pptx 类工具。
- 字体默认 Microsoft YaHei，艺术字体观感可能与原稿有差异；字号/加粗为度量拟合估算。
- 图表**数值标签**可编辑，但折线/柱形等**图形本体**保留在背景图中，改动数字后
  图形不会联动（源文件为位图，矢量结构已丢失）。
- 无墨迹/极小文字受防幻觉保护不重建；竖排、弯曲文字识别率有限。
- 在线服务为同步转换，大文件（>20 页）建议改用任务队列。
