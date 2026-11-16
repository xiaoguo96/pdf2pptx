---
title: 图片型PDF转可编辑PPTX
emoji: 📄
colorFrom: blue
colorTo: indigo
sdk: docker
app_port: 8080
pinned: false
license: apache-2.0
short_description: 上传扫描/图片型PDF，下载文字全部可编辑的PPTX
---

# PDF2PPTX — 图片型 PDF 转可编辑 PPTX

上传**图片型 / 扫描型 PDF**（每页为整页位图、无文字层），下载**可编辑 PPTX**：
页面文字（正文、标题、页眉页脚、页码、目录编号、图表数值标签等）几乎全部还原为
原生文本框（可选中、改写、改字体字号颜色），背景图作为独立图片对象（可移动/缩放），
幻灯片尺寸与 PDF 页面一致。

## 使用方法

1. 打开本页面（创空间应用）；
2. 选择或拖拽上传 PDF（≤32MB）；
3. 点击"开始转换"，等待转换完成（页数越多耗时越长，约每页 5–20 秒）；
4. 浏览器自动下载 `xxx_editable.pptx`，用 PowerPoint / WPS 打开即可编辑。

## 运行模式说明

- 本创空间默认运行**离线 OCR 模式**（RapidOCR，完全本地推理，不依赖外部 API）；
- 若配置了环境变量 `LLM_API_KEY`（创空间设置页-环境变量），则自动切换为
  **视觉大模型模式**（OpenAI 兼容接口，文字识别与布局还原效果更好）；
- API Key 仅通过环境变量注入，未写入任何代码或文件。

## 技术栈

PyMuPDF（渲染）→ 视觉大模型 / RapidOCR（文字与布局提取，双模式自动仲裁）→
OpenCV（背景擦除修复：平滑背景逐行插值 / 复杂背景字形紧致 TELEA）→
PIL 真实字体度量（字号/字距拟合）→ python-pptx（原生对象重建）。

## 源代码仓库

完整源码、README（依赖/配置/启动方式）、实测结果与可编辑性验证材料：
https://github.com/xiaoguo96/pdf2pptx

## 已知限制

- 仅适用于图片型 PDF；带文字层 PDF 建议直接用 PowerPoint 导入。
- 图表数值标签可编辑，但折线/柱形等图形本体保留在背景图中（源为位图，矢量结构已丢失）。
- 字体默认 Microsoft YaHei；无墨迹/极小文字受防幻觉保护不重建。
