FROM python:3.14.0-slim
LABEL maintainer="Xiaoli.Zhang@rr-research.no"
ENV http_proxy=http://proxy.ihelse.net:3128
ENV https_proxy=http://proxy.ihelse.net:3128
# install dependencies
COPY requirements.txt /
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        build-essential=12.12 \
        libjpeg-dev=1:2.1.5-4 \
        libxml2=2.12.7+dfsg+really2.9.14-2.1+deb13u2 \
        libxslt1-dev=1.1.35-1.2+deb13u3 \
        poppler-utils=25.03.0-5+deb13u2 \
        zlib1g-dev=1:1.3.dfsg+really1.3.1-1+b1 \
    && rm -rf /var/lib/apt/lists/* \
    && apt-get clean \
    && pip install --no-cache-dir -r requirements.txt \
    && rm /requirements.txt
# copy over config, template and script
COPY Config /pronto/Config
COPY In /pronto/In
COPY Script /pronto/Script
COPY pronto /pronto/Script/pronto
