FROM python:3.11-alpine3.18
RUN python -m venv /opt/venv
ENV PATH="/opt/venv/bin:$PATH"
COPY requirements.txt requirements.txt
RUN pip install --upgrade pip && \
    pip install --upgrade setuptools && \
    pip install --upgrade wheel && \
    pip install -r requirements.txt

WORKDIR /usr/src/spo_action
COPY src .
CMD ["python", "/usr/src/spo_action/main.py"]
