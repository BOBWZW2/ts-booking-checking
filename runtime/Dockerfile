FROM mcr.microsoft.com/playwright/python:v1.62.0-noble
RUN pip install --no-cache-dir playwright==1.62.0 starlette==1.6.0 uvicorn==0.52.1 openpyxl==3.1.5
WORKDIR /opt/service
COPY boot.py ./
ENV PYTHONUNBUFFERED=1 PYTHONDONTWRITEBYTECODE=1 PORT=10000
USER pwuser
EXPOSE 10000
CMD ["python", "boot.py"]
