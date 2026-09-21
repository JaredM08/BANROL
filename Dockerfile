FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY banrol_motor.py .
COPY templates/ ./templates/
EXPOSE 8000
CMD ["python3", "banrol_motor.py", "--serve", "--port", "8000"]
