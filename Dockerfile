FROM alpine:latest

RUN apk update && apk add openjdk11
RUN apk update && apk add libreoffice
RUN apk update && apk add imagemagick

RUN apk update && apk add python3 py3-pip

WORKDIR /word2md
COPY requirements.txt /word2md/requirements.txt

RUN /usr/bin/pip3 install -r requirements.txt --break-system-packages

COPY . /word2md/

CMD sh /word2md/process_all_docx.sh
