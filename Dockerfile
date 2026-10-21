FROM python:3.11-slim

ENV PYTHONUNBUFFERED=1 \
    DEBIAN_FRONTEND=noninteractive

WORKDIR /app

# Instalar utilitarios del sistema
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# Instalar dependencias de Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Instalar Chromium de Playwright y todas sus dependencias de Linux
RUN playwright install chromium --with-deps

# Copiar el código del proyecto
COPY . .

EXPOSE 8000

# Iniciar servidor FastAPI
CMD ["uvicorn", "servidor_api:app", "--host", "0.0.0.0", "--port", "8000"]

