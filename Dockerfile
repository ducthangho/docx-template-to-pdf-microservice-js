FROM node:12.13

RUN export DEBIAN_FRONTEND=noninteractive
RUN apt-get update
RUN apt-get install -y --force-yes fontconfig && apt-get install -y --force-yes \
    libreoffice \
    unzip \
    zip \
    cabextract \
    xfonts-utils \
    && rm -rf /var/lib/apt/lists/* 

#RUN wget http://ftp.de.debian.org/debian/pool/contrib/m/msttcorefonts/ttf-mscorefonts-installer_3.6_all.deb
#RUN dpkg -i ttf-mscorefonts-installer_3.6_all.deb


RUN mkdir /app && chown node:node /app
USER node
WORKDIR /app/
COPY package.json package-lock.json /app/
COPY Fonts/ /usr/local/share/fonts
RUN fc-cache -f -v
#RUN npm install
ENV NODE_PATH=/app/node_modules
COPY . /app
#RUN chmod 0755 /app/VietnameseTextNormalizer
#RUN npm config set unsafe-perm true
#RUN (cd /app && rm package-lock.json && npm install)
RUN rm package-lock.json
RUN npm install --save node-addon-api
RUN npm install

EXPOSE 8000
CMD [ "node", "server.js" ]
