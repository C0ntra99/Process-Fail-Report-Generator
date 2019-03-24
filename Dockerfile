FROM python:3

copy . PF/
WORKDIR PF/
RUN pip install -r requirements.txt