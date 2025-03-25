
FROM node:18-alpine as build

WORKDIR /app

# Copy package files
COPY package.json package-lock.json ./

# Install dependencies
RUN npm ci

# Copy project files
COPY . .

# Build the app
RUN npm run build

# Production stage
FROM nginx:stable-alpine as production

# Copy built assets from build stage
COPY --from=build /app/dist /usr/share/nginx/html

# Copy nginx config
COPY --from=build /app/nginx.conf /etc/nginx/conf.d/default.conf

# Expose port 80
EXPOSE 80

# Start nginx
CMD ["nginx", "-g", "daemon off;"]
