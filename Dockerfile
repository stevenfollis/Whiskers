# Uses a lightweight Alpine image
FROM node:8.1.3-alpine

# Open Port 443 for https traffic
EXPOSE 443

# Copy application files into container
COPY . /app

# Install dependencies via npm
WORKDIR /app
RUN npm install

# Start application
CMD ["npm", "start"]