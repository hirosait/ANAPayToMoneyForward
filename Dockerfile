# Dockerfile
# ベースイメージとしてPythonを使用
FROM python:3.11-slim

# 作業ディレクトリを設定
WORKDIR /app

# 必要なパッケージをインストール

# 必要なパッケージをインストール
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    gcc \
    libc-dev \
    libffi-dev \
    libssl-dev \
    wget \
    unzip \
    fonts-liberation \
    fonts-ipafont-gothic \
    libappindicator3-1 \
    libasound2 \
    libatk-bridge2.0-0 \
    libatk1.0-0 \
    libcups2 \
    libdbus-1-3 \
    libdrm2 \
    libgbm1 \
    libgtk-3-0 \
    libnspr4 \
    libnss3 \
    libxss1 \
    libxtst6 \
    lsb-release \
    xdg-utils \
    chromium \
    chromium-driver \
    dbus-x11

# スクリーンショットを保存するディレクトリを作成
RUN mkdir -p /app/screenshots

# pythonの依存関係をインストール
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY requirements.txt /app/requirements.txt
COPY anapay2mf.py /app/anapay2mf.py
COPY quickstart.py /app/quickstart.py
COPY credentials.json /app/credentials.json
COPY token.json /app/token.json
COPY .env /app/.env

# 環境変数の読み込み
RUN pip install python-dotenv

# Heliumが必要とする追加ライブラリをインストール
RUN pip install selenium helium

# エントリーポイントとしてスクリプトを指定
CMD ["python", "anapay2mf.py"]
