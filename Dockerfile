FROM python:3.11-slim
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr \
    libglib2.0-0 \
    libsm6 \
    libxrender1 \
    libxext6 \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt
COPY . /app
ENV STREAMLIT_SERVER_PORT=8501
EXPOSE 8501
CMD ["bash", "-lc", "streamlit run app.py --server.port ${PORT:-8501} --server.address 0.0.0.0"]
