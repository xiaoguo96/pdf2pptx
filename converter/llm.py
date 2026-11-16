# -*- coding: utf-8 -*-
"""OpenAI 兼容视觉大模型布局提取器（第三方 API，key 仅从环境变量注入）。

环境变量：
  LLM_API_KEY   必填，未设置则 LLM 模式自动关闭，回退本地 OCR 管线
  LLM_API_BASE  可选，默认 https://www.cnxmai.com/v1
  LLM_MODEL     可选，默认 gpt-5.6-luna
"""
import base64
import json
import os
import ssl
import urllib.request
from dataclasses import dataclass

import cv2

DEFAULT_BASE = "https://www.cnxmai.com/v1"
DEFAULT_MODEL = "gpt-5.6-luna"

PROMPT = """这是演示文稿的一页幻灯片图片（尺寸 %dx%d 像素）。
请识别页面上所有【文字元素】，输出 JSON 数组，每个元素字段：
- text: 文字内容（同一文本块的多行用 \\n 分隔）
- bbox: [x0, y0, x1, y1] 像素坐标（左上角为原点）
- kind: 元素类别，取值 "body"(正文/标题) | "page_number"(页码) | "logo"(品牌名/logo文字) | "chart_label"(图表数值/百分比/年份/轴标签) | "badge"(编号徽章) | "header"(页眉带文字) | "footer"(页脚带文字)
- color: 文字颜色近似值，如 "#333333"
- bold: 是否加粗 true/false
只输出 JSON，不要其他文字。"""


@dataclass
class Element:
    text: str
    x: float      # px, in full-render space
    y: float
    w: float
    h: float
    kind: str
    color_hex: str
    bold: bool


class LlmLayoutExtractor:
    def __init__(self, api_key=None, base=None, model=None):
        self.api_key = api_key or os.environ.get("LLM_API_KEY", "")
        self.base = base or os.environ.get("LLM_API_BASE", DEFAULT_BASE)
        self.model = model or os.environ.get("LLM_MODEL", DEFAULT_MODEL)
        self._ctx = ssl._create_unverified_context()

    @classmethod
    def available(cls):
        return bool(os.environ.get("LLM_API_KEY", ""))

    def _post(self, payload):
        req = urllib.request.Request(
            self.base.rstrip("/") + "/chat/completions",
            data=json.dumps(payload).encode("utf-8"),
            headers={"Authorization": f"Bearer {self.api_key}",
                     "Content-Type": "application/json"})
        last = None
        for _ in range(2):
            try:
                with urllib.request.urlopen(req, timeout=180,
                                            context=self._ctx) as resp:
                    return json.loads(resp.read().decode("utf-8"))
            except Exception as e:  # 429/5xx/网络抖动重试一次
                last = e
        raise last

    def extract(self, img_bgr, max_w=1302):
        """Return list[Element] for one page raster."""
        h, w = img_bgr.shape[:2]
        scale = max_w / float(w) if w > max_w else 1.0
        small = (cv2.resize(img_bgr, (int(w * scale), int(h * scale)))
                 if scale != 1.0 else img_bgr)
        sh, sw = small.shape[:2]
        ok, buf = cv2.imencode(".jpg", small, [cv2.IMWRITE_JPEG_QUALITY, 85])
        if not ok:
            raise RuntimeError("jpeg encode failed")
        b64 = base64.b64encode(buf.tobytes()).decode()
        payload = {
            "model": self.model,
            "max_tokens": 6000,
            "messages": [{
                "role": "user",
                "content": [
                    {"type": "text", "text": PROMPT % (sw, sh)},
                    {"type": "image_url",
                     "image_url": {"url": f"data:image/jpeg;base64,{b64}"}},
                ]}]
        }
        data = self._post(payload)
        content = data["choices"][0]["message"]["content"]
        return self._parse(content, scale)

    @staticmethod
    def _parse(content, scale):
        s, e = content.find("["), content.rfind("]")
        if s < 0 or e <= s:
            raise ValueError("no JSON array in model output")
        out = []
        for el in json.loads(content[s:e + 1]):
            try:
                x0, y0, x1, y1 = [float(v) for v in el["bbox"]]
            except Exception:
                continue
            text = str(el.get("text", "")).strip()
            if not text or x1 - x0 < 2 or y1 - y0 < 2:
                continue
            out.append(Element(
                text=text,
                x=x0 / scale, y=y0 / scale,
                w=(x1 - x0) / scale, h=(y1 - y0) / scale,
                kind=str(el.get("kind", "body")).lower(),
                color_hex=str(el.get("color", "") or ""),
                bold=bool(el.get("bold", False)),
            ))
        return out
