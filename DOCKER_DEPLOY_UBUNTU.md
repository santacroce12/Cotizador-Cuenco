# Deploy Docker en Ubuntu Server

## 1. Requisitos

- Ubuntu Server con Docker y Docker Compose plugin instalados
- Puerto `8000` libre en el servidor, o un reverse proxy delante
- Dominio configurado para apuntar al servidor si vas a usar `https://cuencotech.com`

## 2. Archivos persistentes

Este proyecto ya quedó preparado para persistir:

- Base SQLite: `./data/database.db`
- Imágenes de productos: `./storage/uploads/productos`

No van dentro de la imagen. Quedan fuera del contenedor.

## 3. Configuración

Crear el archivo `.env` a partir del ejemplo:

```bash
cp .env.example .env
```

Completar como mínimo:

- `APP_SECRET_KEY`
- `ADMIN_SETUP_TOKEN`
- `APP_BASE_URL`
- `SMTP_PASSWORD`

## 4. Build y arranque

```bash
docker compose build
docker compose up -d
```

Ver logs:

```bash
docker compose logs -f
```

Estado:

```bash
docker compose ps
```

## 5. URL local del servicio

El contenedor expone la app en:

- `http://IP_DEL_SERVIDOR:8000`

Si vas a publicar con dominio, lo correcto es poner Nginx o Apache delante y dejar Docker escuchando en `8000`.

## 6. Reverse proxy recomendado

Ejemplo mínimo con Nginx:

```nginx
server {
    listen 80;
    server_name cuencotech.com;

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_set_header X-Forwarded-Host $host;
    }
}
```

Con HTTPS, agregá Certbot o el certificado que uses en tu infraestructura.

## 7. Actualización de versión

```bash
git pull
docker compose build
docker compose up -d
```

## 8. Backup

Respaldar al menos:

- `./data/database.db`
- `./storage/uploads/productos/`
- `.env`

## 9. Primer acceso

Si la base está vacía, entrá a:

- `/setup-admin`

Ahí se crea el primer administrador usando `ADMIN_SETUP_TOKEN`.
