# Deploy en DonWeb

## Archivos que tienen que estar subidos

- `app.py`
- `templates/`
- `static/`
- `requirements.txt`
- `wsgi.py`
- `passenger_wsgi.py`
- `local_settings.json`

No subas:

- `database.db` si queres arrancar limpio
- `__pycache__/`
- `static/uploads/productos/` si no queres arrastrar archivos viejos

## 1. Verificar el tipo de hosting

Tu plan tiene que soportar aplicaciones Python con Passenger o equivalente.

Si el plan no soporta Python, DonWeb va a seguir devolviendo `403 Forbidden` o no va a ofrecer la opcion de crear la app.

## 2. Crear la aplicacion Python en DonWeb

En el panel de DonWeb:

1. Crear una aplicacion Python.
2. Elegir como directorio de la app la carpeta donde subiste este proyecto.
3. Elegir como archivo de entrada `passenger_wsgi.py`.
4. Elegir una version de Python 3.11 o 3.12 si esta disponible.

## 3. Instalar dependencias

Dentro del entorno virtual que genere DonWeb, correr:

```bash
pip install -r requirements.txt
```

## 4. Crear `local_settings.json`

Usa `local_settings.example.json` como base y cargale los valores reales de produccion.

Campos minimos recomendados:

- `APP_SECRET_KEY`
- `APP_BASE_URL`
- `SMTP_HOST`
- `SMTP_PORT`
- `SMTP_USE_SSL`
- `SMTP_USE_TLS`
- `SMTP_USERNAME`
- `SMTP_PASSWORD`
- `SMTP_FROM`
- `FOLLOWUP_DEFAULT_TO_EMAIL`

Ejemplo:

```json
{
  "APP_SECRET_KEY": "una-clave-larga-y-aleatoria",
  "SMTP_HOST": "smtp.example.com",
  "SMTP_PORT": 465,
  "SMTP_USE_SSL": true,
  "SMTP_USE_TLS": false,
  "SMTP_USERNAME": "usuario@example.com",
  "SMTP_PASSWORD": "TU_PASSWORD_REAL",
  "SMTP_FROM": "usuario@example.com",
  "APP_BASE_URL": "https://cotizador.example.com",
  "FOLLOWUP_DEFAULT_TO_EMAIL": "equipo@example.com"
}
```

## 5. Permisos de escritura

Estas rutas tienen que ser escribibles por el proceso web:

- `database.db`
- `static/uploads/productos/`

Permisos sugeridos:

- carpetas: `755`
- archivos: `644`

Si el proceso web no puede escribir, vas a tener errores al:

- crear cotizaciones
- guardar usuarios
- subir imagenes
- registrar auditoria

## 6. Primer inicio

1. Reiniciar la aplicacion desde el panel.
2. Entrar a `/login`.
3. Usar una cuenta administradora existente.

## 7. Si sigue apareciendo `403 Forbidden`

Eso normalmente significa una de estas cosas:

1. El dominio apunta a una carpeta estatica y no a la app Python.
2. La app Python no fue creada en DonWeb.
3. `passenger_wsgi.py` no esta en el directorio correcto.
4. El plan no soporta Python/Passenger.

## 8. Verificacion minima

Una vez publicada, estas URLs deberian responder:

- `/login`
- `/`
- `/historial`
- `/dashboard`
