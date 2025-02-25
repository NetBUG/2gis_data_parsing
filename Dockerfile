FROM ubuntu:18.04

ENV TZ=Etc/UTC
ENV DEBIAN_FRONTEND=noninteractive 

RUN apt-get update && \
    apt-get -y install curl git software-properties-common && \
    add-apt-repository ppa:ondrej/php && \
    apt-get update
RUN apt-get -y install php7.0 php7.2-curl php7.2-xml php7.2-mbstring
# php7.0-xml php7.0-mbstring php7.0-composer

WORKDIR /app

# RUN git clone https://github.com/mbry/DgdatToXlsx
RUN https://github.com/eugenemarenin/2gis_data_parsing .
RUN mv download Download

RUN php -r "copy('https://getcomposer.org/installer', 'composer-setup.php');" && \
    php composer-setup.php && \
    php -r "unlink('composer-setup.php');"

RUN php composer.phar update

CMD php --version
