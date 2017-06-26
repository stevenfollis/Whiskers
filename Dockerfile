FROM node:8.1.2

EXPOSE 3978
EXPOSE 443
EXPOSE 56791

COPY . /app

WORKDIR /app
RUN npm install

CMD ["npm", "start"]