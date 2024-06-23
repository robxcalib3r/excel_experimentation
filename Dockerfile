FROM python:3.9

RUN apt-get update && apt-get install -y iputils-ping

WORKDIR /app

COPY requirements.txt .

RUN pip install -r requirements.txt

COPY . .

CMD ["python", "./auto_ping_from_excel.py"]