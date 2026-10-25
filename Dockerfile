FROM php:8.3-apache-bookworm

RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 python3-pip python3-venv \
    && rm -rf /var/lib/apt/lists/*

# Apache: document root = /var/www/html/public
ENV APACHE_DOCUMENT_ROOT=/var/www/html/public
RUN sed -ri -e 's!/var/www/html!${APACHE_DOCUMENT_ROOT}!g' /etc/apache2/sites-available/*.conf \
 && sed -ri -e 's!/var/www/!${APACHE_DOCUMENT_ROOT}!g' /etc/apache2/apache2.conf /etc/apache2/conf-available/*.conf \
 && a2enmod rewrite headers \
 && echo "upload_max_filesize = 40M" > /usr/local/etc/php/conf.d/uploads.ini \
 && echo "post_max_size = 45M" >> /usr/local/etc/php/conf.d/uploads.ini \
 && echo "max_execution_time = 300" >> /usr/local/etc/php/conf.d/uploads.ini \
 && echo "memory_limit = 512M" >> /usr/local/etc/php/conf.d/uploads.ini

WORKDIR /var/www/html

COPY worker/requirements.txt /var/www/html/worker/requirements.txt
RUN pip3 install --break-system-packages --no-cache-dir -r /var/www/html/worker/requirements.txt

COPY public /var/www/html/public
COPY worker /var/www/html/worker

RUN mkdir -p /var/www/html/storage/uploads /var/www/html/storage/output \
 && chown -R www-data:www-data /var/www/html/storage \
 && chmod +x /var/www/html/worker/build_fiel.py

ENV PYTHON_BIN=python3
ENV STIRLING_URL=https://pdf.klowshop.com

EXPOSE 80
CMD ["apache2-foreground"]
