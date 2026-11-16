FROM python:3.10-slim
WORKDIR /app
RUN apt-get update && apt-get install -y --no-install-recommends \
    libglib2.0-0 libgl1 libgomp1 \
    && rm -rf /var/lib/apt/lists/*
COPY requirements.txt .
RUN pip install --no-cache-dir \
    -i https://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com \
    -r requirements.txt \
    && python -c "import pymupdf, cv2, numpy, pptx, flask, rapidocr_onnxruntime; print('deps ok')"
COPY . .
ENV PORT=8080
EXPOSE 8080
CMD ["python", "app.py"]
