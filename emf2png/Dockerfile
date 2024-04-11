FROM alpine:latest

RUN apk update && apk add libreoffice
RUN apk update && apk add imagemagick

WORKDIR /app

COPY . /app

CMD sh ./entrypoint.sh