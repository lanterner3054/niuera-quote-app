FROM python:3.12-slim

# 腾讯云内网 apt 源(仅在腾讯云服务器上构建时有效,同机房高速)
RUN for f in /etc/apt/sources.list.d/debian.sources /etc/apt/sources.list; do \
        [ -f "$f" ] && sed -i -E 's|https?://deb.debian.org|http://mirrors.tencentyun.com|g' "$f"; \
    done; true

# libreoffice-calc: 让「导出 PDF」按钮在云端可用; noto-cjk: 单据内中文字体
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-calc fonts-noto-cjk tzdata \
    && rm -rf /var/lib/apt/lists/*

ENV TZ=Asia/Shanghai
WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt \
        -i https://mirrors.cloud.tencent.com/pypi/simple/

COPY app.py ./
COPY engine ./engine
COPY static ./static
COPY tools ./tools

# data/ output/ legacy/ 由 volume 挂载,不打进镜像
CMD ["python", "-m", "uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
