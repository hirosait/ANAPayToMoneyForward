FROM arm64v8/python:3.11-slim

# 必要なパッケージのインストール
RUN apt-get update && apt-get install -y \
    chromium \
    chromium-driver \
    xvfb \
    xauth \
    fonts-liberation \
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
    libxcomposite1 \
    libxdamage1 \
    libxrandr2 \
    xdg-utils \
    --no-install-recommends

# pipでseleniumとHeliumをインストール
RUN pip install selenium helium

# 環境変数の設定
ENV DISPLAY=:99

WORKDIR /app

# スクリーンショットを保存するディレクトリを作成
RUN mkdir -p /app/screenshots

# pythonの依存関係をインストール
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY requirements.txt /app/requirements.txt
COPY anapay2mf.py /app/anapay2mf.py
COPY quickstart.py /app/quickstart.py
COPY service-account.json /app/service-account.json
COPY .env /app/.env

# 環境変数の読み込み
RUN pip install python-dotenv

# エントリーポイントの設定
ENTRYPOINT ["python", "/app/anapay2mf.py"]