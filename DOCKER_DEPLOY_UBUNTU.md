# Deploy Docker en Ubuntu Server

## 1. Requisitos

- Ubuntu Server con Docker y Docker Compose plugin instalados
- Puerto `9000` libre en el servidor, o un reverse proxy delante
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

- `APP_SECRET_KEY`: **obligatoria**. La app no arranca sin una clave propia de 16 caracteres o más. Generala con:

  ```bash
  python3 -c "import secrets; print(secrets.token_urlsafe(48))"
  ```

  No la cambies después: si cambia, se cierran todas las sesiones.
- `APP_BASE_URL`
- `SMTP_PASSWORD`
- `AUTH_COOKIE_SECURE=true` **solo si** se entra por HTTPS (ver punto 6). Con `http://` dejala en `false` o no se va a poder iniciar sesión.

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

- `http://IP_DEL_SERVIDOR:9000`

Si vas a publicar con dominio, lo correcto es poner Nginx o Apache delante y dejar Docker escuchando en `9000`.

## 6. Reverse proxy recomendado

Ejemplo mínimo con Nginx:

```nginx
server {
    listen 80;
    server_name cuencotech.com;

    location / {
        proxy_pass http://127.0.0.1:9000;
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

## 8. Backups

El script `scripts/backup.py` genera un `.zip` con la base (copiada de forma segura aunque la app esté en uso, y verificada) y las fotos de productos. Guarda los zips en `./data/backups/` y borra los de más de 30 días.

Probarlo a mano:

```bash
docker compose exec -T cotizador python scripts/backup.py
```

Programarlo todos los días a las 2 AM (`crontab -e` en el servidor):

```cron
0 2 * * * cd /ruta/al/cotizador && docker compose exec -T cotizador python scripts/backup.py >> /var/log/cotizador-backup.log 2>&1
```

**Importante:** un backup en el mismo disco no sirve si se rompe el disco o el servidor. Copiá `./data/backups/` a otro lugar todos los días (Google Drive o Dropbox con `rclone`, un NAS, otra PC). Guardá también una copia del `.env` en un lugar seguro: sin la `APP_SECRET_KEY` y la contraseña SMTP no se puede reconstruir el servicio.

Restaurar un backup:

```bash
docker compose down
unzip cotizador-backup-AAAAMMDD-HHMMSS.zip -d /tmp/restaurar
cp /tmp/restaurar/database.db ./data/database.db
rm -f ./data/database.db-wal ./data/database.db-shm
cp -r /tmp/restaurar/uploads/productos/. ./storage/uploads/productos/
docker compose up -d
```

Probá restaurar un backup en otra máquina al menos una vez: un backup que nunca se probó no es un backup.

## 9. Primer acceso

Crear el primer administrador (pide la contraseña sin mostrarla):

```bash
docker compose exec cotizador flask --app app crear-admin admin
```

El mismo comando sirve para recuperar el acceso si se olvida la contraseña del administrador: la cambia y reactiva el usuario.

Después entrar a `/login`. Desde **Usuarios** el administrador crea las cuentas, carga los datos de firma de cada uno, cambia contraseñas y da de baja a quien ya no trabaja en la empresa (no se borran: sus cotizaciones y la auditoría quedan). Cada usuario cambia su propia contraseña desde **Mi cuenta** (su nombre arriba a la derecha).

Después de 5 intentos fallidos de login, ese usuario queda bloqueado 15 minutos desde esa IP.
