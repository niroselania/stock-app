FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY index.html processor.py server.py ./

ENV HOST=0.0.0.0
ENV PORT=8765

EXPOSE 8765

CMD ["python", "server.py"]
