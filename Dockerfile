FROM alpine:latest

RUN apk update
RUN apk add python3 py3-pip

WORKDIR /word2md
COPY requirements.txt /word2md/requirements.txt

# RUN /usr/bin/pip3 install python-docx chevron markdown pyyaml --break-system-packages
RUN /usr/bin/pip3 install -r requirements.txt --break-system-packages

COPY . /word2md/

CMD sh /word2md/process_all_docx.sh
