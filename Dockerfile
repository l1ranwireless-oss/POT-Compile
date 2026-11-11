FROM nginx:alpine
ENV API_HOST=excel-api \
    API_PORT=8000
COPY index.html /usr/share/nginx/html/index.html
COPY excel-compiler.html /usr/share/nginx/html/excel-compiler.html
COPY nginx.conf.template /etc/nginx/templates/default.conf.template
EXPOSE 80