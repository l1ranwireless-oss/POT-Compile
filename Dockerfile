FROM nginx:alpine
COPY index.html /usr/share/nginx/html/index.html
COPY excel-compiler.html /usr/share/nginx/html/excel-compiler.html
COPY nginx.conf /etc/nginx/conf.d/default.conf
EXPOSE 80