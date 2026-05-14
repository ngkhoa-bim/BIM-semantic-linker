# =============================================================
# BIM Semantic Linker — Backend Dockerfile
# Base: python:3.11-slim
# Compatible: Render, Cloud Run, local Docker
# =============================================================

FROM python:3.11-slim

# ── System deps cần cho ifcopenshell ────────────────────────
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    libglu1-mesa \
    libxi6 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# ── Cài Python dependencies trước (tận dụng Docker layer cache) ──
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# ── Copy source code ─────────────────────────────────────────
COPY app.py .

# ── Thư mục upload tạm ───────────────────────────────────────
RUN mkdir -p /tmp/ifc_uploads

# ── Biến môi trường mặc định (override bằng -e hoặc .env) ────
ENV FLASK_ENV=production
ENV UPLOAD_FOLDER=/tmp/ifc_uploads
ENV PORT=7860

EXPOSE 7860

# ── Dùng gunicorn cho production ──────────────────────────────
# Cài gunicorn riêng (không cần thêm vào requirements.txt)
RUN pip install --no-cache-dir gunicorn

CMD ["gunicorn", "--bind", "0.0.0.0:7860", "--workers", "2", "--timeout", "300", "app:app"]

# ── Hướng dẫn build & run local ──────────────────────────────
# docker build -t bim-backend .
# docker run -p 7860:7860 \
#   -e NEO4J_URI=neo4j+s://xxxx.databases.neo4j.io \
#   -e NEO4J_USER=neo4j \
#   -e NEO4J_PASSWORD=yourpassword \
#   -e GEMINI_MODEL="gemini-1.5-pro" \
#  bim-semantic-backend