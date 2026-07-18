FROM node:20

WORKDIR /usr/src/app
# OS 보안 패치
RUN apt-get update && apt-get upgrade -y && rm -rf /var/lib/apt/lists/*

COPY package*.json ./
RUN npm config set registry https://nexus.hedej.lge.com/repository/npm-group/ --global && \
    npm install -f

COPY . .
EXPOSE 8080
RUN ln -sf /dev/stdout /usr/src/app/access.log && ln -sf /dev/stderr /usr/src/app/error.log
CMD [ "npm", "start" ]
