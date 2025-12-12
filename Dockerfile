FROM python:3.10-slim

WORKDIR /app

# 安装系统依赖
RUN apt-get update && apt-get install -y \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# 复制依赖文件
COPY requirements.txt .

# 安装 Python 依赖
RUN pip install --no-cache-dir -r requirements.txt
RUN pip install --no-cache-dir gunicorn

# 复制应用代码
COPY . .

# 创建输出目录
RUN mkdir -p /app/web/output

# 暴露端口 (Hugging Face Spaces 使用 7860)
EXPOSE 7860

# 启动命令 - 增加超时时间，使用 gevent worker 支持长连接
CMD ["gunicorn", "--bind", "0.0.0.0:7860", "--workers", "2", "--timeout", "300", "--keep-alive", "75", "web.app:app"]
