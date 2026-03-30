import os
import json
import re
import smtplib
import threading
import time
from collections import defaultdict
from datetime import datetime, timedelta
from email.message import EmailMessage
from functools import wraps
from io import BytesIO
from pathlib import Path

import jwt
from flask import Flask, Response, flash, g, jsonify, redirect, render_template, request, url_for
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image, ImageOps, UnidentifiedImageError
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func, or_
from werkzeug.exceptions import RequestEntityTooLarge
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename


app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)


def resolver_ruta_configurada(valor, default_path):
    valor = (valor or "").strip()
    if not valor:
        return default_path
    ruta = Path(valor)
    if not ruta.is_absolute():
        ruta = (Path(app.root_path) / ruta).resolve()
    return ruta


DATA_DIR = resolver_ruta_configurada(os.getenv("DATA_DIR"), Path(app.root_path))
DATA_DIR.mkdir(parents=True, exist_ok=True)
db_path = resolver_ruta_configurada(os.getenv("DATABASE_PATH"), DATA_DIR / "database.db")
db_path.parent.mkdir(parents=True, exist_ok=True)
app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{db_path.as_posix()}"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["MAX_CONTENT_LENGTH"] = 40 * 1024 * 1024
db = SQLAlchemy(app)

NOMBRE_FANTASIA = "Cuenco Tech"
RAZON_SOCIAL = "Cuenco Tech S.A."
CUIT = "30-71831614-2"
DOMICILIO = "Rafael Cubillos 2056, M5500 Godoy Cruz, Mendoza"
DEFAULT_FOLLOWUP_EMAIL = "cotizador@cuencotech.com"
DEFAULT_SMTP_HOST = "a0021139.ferozo.com"
DEFAULT_SMTP_PORT = 465
DEFAULT_SMTP_USERNAME = "cotizador@cuencotech.com"
DEFAULT_SMTP_FROM = "cotizador@cuencotech.com"
DEFAULT_APP_BASE_URL = "https://cuencotech.com"
LOCAL_SETTINGS_PATH = resolver_ruta_configurada(os.getenv("LOCAL_SETTINGS_PATH"), Path(app.root_path) / "local_settings.json")


def cargar_local_settings():
    if not LOCAL_SETTINGS_PATH.exists():
        return {}
    try:
        data = json.loads(LOCAL_SETTINGS_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


LOCAL_SETTINGS = cargar_local_settings()
app.config["SECRET_KEY"] = (
    os.getenv("APP_SECRET_KEY")
    or os.getenv("SECRET_KEY")
    or str(LOCAL_SETTINGS.get("APP_SECRET_KEY") or LOCAL_SETTINGS.get("SECRET_KEY") or "").strip()
    or "change-this-secret-in-local-settings"
)
AUTH_COOKIE_SECURE = str(
    os.getenv("AUTH_COOKIE_SECURE") or LOCAL_SETTINGS.get("AUTH_COOKIE_SECURE") or ""
).strip().lower() in ("1", "true", "yes")

PLACEHOLDER_PRODUCTO = "placeholder_product.png"
UPLOADS_PRODUCTOS_DIR = resolver_ruta_configurada(
    os.getenv("UPLOADS_PRODUCTOS_DIR"), Path(app.static_folder) / "uploads" / "productos"
)
UPLOADS_PRODUCTOS_DIR.mkdir(parents=True, exist_ok=True)
ALLOWED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".gif"}
MAX_IMAGE_UPLOAD_BYTES = 6 * 1024 * 1024
PRODUCT_IMAGE_MAX_DIMENSION = 1400
PRODUCT_IMAGE_TARGET_BYTES = 450 * 1024
PRODUCT_IMAGE_QUALITY_STEPS = (82, 76, 70, 64, 58)
HISTORIAL_PER_PAGE = 20
DASHBOARD_OPERATIVO_PER_PAGE = 15
AUDITORIA_PER_PAGE = 25
ESTADOS_COTIZACION = ("En progreso", "Aceptada", "Rechazada")
FAMILIAS_COTIZACION = (
    "SEGURIDAD URBANA",
    "PARKING",
    "TRANSPORTE INTELIGENTE",
    "CONECTIVIDAD SATELITAL",
    "SALAS DE CONTROL",
    "SMART CITIES",
)
SECTORES_CLIENTE = {
    "Publico": ("Municipal", "Provincial", "Nacional", "Otro"),
    "Privado": ("Energía", "Agroindustria", "Hotelería", "Educación", "Turismo", "Retail", "Otro"),
}
REMINDER_POLL_SECONDS = 60
_reminder_worker_lock = threading.Lock()
_reminder_worker_started = False
NUMERO_COTIZACION_PREFIX = "CT"


class Cotizacion(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    numero_cotizacion = db.Column(db.String(20))
    estado = db.Column(db.String(20), default="En progreso")
    seguimiento_activo = db.Column(db.Boolean, default=False)
    seguimiento_email = db.Column(db.String(150))
    seguimiento_cada_dias = db.Column(db.Integer)
    seguimiento_proximo_envio = db.Column(db.DateTime)
    seguimiento_ultimo_envio = db.Column(db.DateTime)
    nombre_fantasia = db.Column(db.String(100))
    razon_social = db.Column(db.String(100))
    cuit = db.Column(db.String(50))
    cliente = db.Column(db.String(100), nullable=False)
    cliente_id = db.Column(db.Integer, db.ForeignKey("cliente.id"))
    cliente_razon_social = db.Column(db.String(100))
    cliente_cuit = db.Column(db.String(50))
    familia = db.Column(db.String(50))
    fecha = db.Column(db.DateTime, default=datetime.utcnow)
    moneda = db.Column(db.String(10), default="ARS")
    condicion_iva = db.Column(db.String(50))
    total_neto = db.Column(db.Float)
    total_iva = db.Column(db.Float)
    total_final = db.Column(db.Float)
    items = db.relationship("ItemCotizacion", backref="parent", cascade="all, delete-orphan")


class ItemCotizacion(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    cotizacion_id = db.Column(db.Integer, db.ForeignKey("cotizacion.id"))
    descripcion = db.Column(db.String(200))
    cantidad = db.Column(db.Float)
    costo_unitario = db.Column(db.Float)
    costo_extra = db.Column(db.Float, default=5.0)
    margen = db.Column(db.Float)
    iva_item = db.Column(db.Float, default=21.0)
    precio_venta = db.Column(db.Float)
    subtotal = db.Column(db.Float)
    imagen_url = db.Column(db.String(500))

    @property
    def iva_porcentaje(self):
        # Alias para compatibilidad con plantillas previas.
        return self.iva_item or 0.0


class Cliente(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    razon_social = db.Column(db.String(100))
    cuit = db.Column(db.String(20), unique=True)
    domicilio = db.Column(db.String(200))
    sector = db.Column(db.String(50))
    subsector = db.Column(db.String(50))
    email = db.Column(db.String(100))
    telefono = db.Column(db.String(50))
    condicion_iva = db.Column(db.String(50), default="Consumidor Final")
    cotizaciones = db.relationship("Cotizacion", backref="cliente_ref", lazy=True)


class Usuario(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    is_admin = db.Column(db.Boolean, default=False, nullable=False)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)


class Auditoria(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    usuario_id = db.Column(db.Integer, db.ForeignKey("usuario.id"))
    username = db.Column(db.String(50), nullable=False)
    accion = db.Column(db.String(80), nullable=False)
    tipo_entidad = db.Column(db.String(50), nullable=False)
    entidad_id = db.Column(db.Integer)
    entidad_ref = db.Column(db.String(80))
    detalle = db.Column(db.Text)
    fecha = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    usuario = db.relationship("Usuario", backref="auditorias", lazy=True)


def formatear_numero_cotizacion(anio, secuencia):
    return f"{NUMERO_COTIZACION_PREFIX}-{int(anio):04d}-{int(secuencia):04d}"


def parsear_numero_cotizacion(valor):
    numero = (valor or "").strip().upper()
    match = re.fullmatch(rf"{NUMERO_COTIZACION_PREFIX}-(\d{{4}})-(\d+)", numero)
    if not match:
        return None
    return int(match.group(1)), int(match.group(2))


def generar_token_usuario(usuario):
    token = jwt.encode(
        {
            "user_id": usuario.id,
            "exp": datetime.utcnow() + timedelta(hours=24),
        },
        app.config["SECRET_KEY"],
        algorithm="HS256",
    )
    return token if isinstance(token, str) else token.decode("utf-8")


def obtener_usuario_desde_token(token):
    if not token:
        return None
    try:
        data = jwt.decode(token, app.config["SECRET_KEY"], algorithms=["HS256"])
    except jwt.PyJWTError:
        return None
    return db.session.get(Usuario, data.get("user_id"))


def token_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        token = request.cookies.get("x-access-token")
        current_user = obtener_usuario_desde_token(token)
        if not current_user:
            if request.endpoint in {"agregar_cliente", "actualizar_estado_cotizacion", "filtrar_historial", "eliminar_cotizacion"}:
                return jsonify({"error": "auth_required"}), 401
            return redirect(url_for("login", next=request.path))
        g.current_user = current_user
        return f(*args, **kwargs)

    return decorated


def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        current_user = getattr(g, "current_user", None)
        if not current_user:
            token = request.cookies.get("x-access-token")
            current_user = obtener_usuario_desde_token(token)
            if current_user:
                g.current_user = current_user
        if not current_user:
            return redirect(url_for("login", next=request.path))
        if not current_user.is_admin:
            if request.endpoint in {"eliminar_cotizacion"}:
                return jsonify({"error": "admin_required"}), 403
            return redirect(url_for("index"))
        return f(*args, **kwargs)

    return decorated


@app.context_processor
def inject_current_user():
    return {
        "current_user": getattr(g, "current_user", None),
    }


with app.app_context():
    db.create_all()
    columnas_item = [col[1] for col in db.session.execute(db.text("PRAGMA table_info(item_cotizacion)")).fetchall()]
    if "iva_item" not in columnas_item:
        db.session.execute(db.text("ALTER TABLE item_cotizacion ADD COLUMN iva_item FLOAT DEFAULT 21.0"))
        db.session.commit()
    if "costo_extra" not in columnas_item:
        db.session.execute(db.text("ALTER TABLE item_cotizacion ADD COLUMN costo_extra FLOAT DEFAULT 5.0"))
        db.session.commit()
    columnas_cot = [col[1] for col in db.session.execute(db.text("PRAGMA table_info(cotizacion)")).fetchall()]
    if "moneda" not in columnas_cot:
        db.session.execute(db.text("ALTER TABLE cotizacion ADD COLUMN moneda VARCHAR(10) DEFAULT 'ARS'"))
        db.session.commit()
    if "condicion_iva" not in columnas_cot:
        db.session.execute(db.text("ALTER TABLE cotizacion ADD COLUMN condicion_iva VARCHAR(50)"))
        db.session.commit()
    if "numero_cotizacion" not in columnas_cot:
        db.session.execute(db.text("ALTER TABLE cotizacion ADD COLUMN numero_cotizacion VARCHAR(20)"))
        db.session.commit()
    if "estado" not in columnas_cot:
        db.session.execute(db.text("ALTER TABLE cotizacion ADD COLUMN estado VARCHAR(20) DEFAULT 'En progreso'"))
        db.session.commit()
    if "seguimiento_activo" not in columnas_cot:
        db.session.execute(db.text("ALTER TABLE cotizacion ADD COLUMN seguimiento_activo BOOLEAN DEFAULT 0"))
        db.session.commit()
    if "seguimiento_email" not in columnas_cot:
        db.session.execute(db.text("ALTER TABLE cotizacion ADD COLUMN seguimiento_email VARCHAR(150)"))
        db.session.commit()
    if "seguimiento_cada_dias" not in columnas_cot:
        db.session.execute(db.text("ALTER TABLE cotizacion ADD COLUMN seguimiento_cada_dias INTEGER"))
        db.session.commit()
    if "seguimiento_proximo_envio" not in columnas_cot:
        db.session.execute(db.text("ALTER TABLE cotizacion ADD COLUMN seguimiento_proximo_envio DATETIME"))
        db.session.commit()
    if "seguimiento_ultimo_envio" not in columnas_cot:
        db.session.execute(db.text("ALTER TABLE cotizacion ADD COLUMN seguimiento_ultimo_envio DATETIME"))
        db.session.commit()
    if "cliente_razon_social" not in columnas_cot:
        db.session.execute(db.text("ALTER TABLE cotizacion ADD COLUMN cliente_razon_social VARCHAR(100)"))
        db.session.commit()
    if "cliente_cuit" not in columnas_cot:
        db.session.execute(db.text("ALTER TABLE cotizacion ADD COLUMN cliente_cuit VARCHAR(50)"))
        db.session.commit()
    if "cliente_id" not in columnas_cot:
        db.session.execute(db.text("ALTER TABLE cotizacion ADD COLUMN cliente_id INTEGER"))
        db.session.commit()
    if "familia" not in columnas_cot:
        db.session.execute(db.text("ALTER TABLE cotizacion ADD COLUMN familia VARCHAR(50)"))
        db.session.commit()
    db.session.execute(
        db.text("UPDATE cotizacion SET estado = 'En progreso' WHERE estado IS NULL OR TRIM(estado) = ''")
    )
    db.session.commit()
    columnas_cliente = [col[1] for col in db.session.execute(db.text("PRAGMA table_info(cliente)")).fetchall()]
    if "domicilio" not in columnas_cliente:
        db.session.execute(db.text("ALTER TABLE cliente ADD COLUMN domicilio VARCHAR(200)"))
        db.session.commit()
    if "sector" not in columnas_cliente:
        db.session.execute(db.text("ALTER TABLE cliente ADD COLUMN sector VARCHAR(50)"))
        db.session.commit()
    if "subsector" not in columnas_cliente:
        db.session.execute(db.text("ALTER TABLE cliente ADD COLUMN subsector VARCHAR(50)"))
        db.session.commit()
    if "email" not in columnas_cliente:
        db.session.execute(db.text("ALTER TABLE cliente ADD COLUMN email VARCHAR(100)"))
        db.session.commit()
    if "telefono" not in columnas_cliente:
        db.session.execute(db.text("ALTER TABLE cliente ADD COLUMN telefono VARCHAR(50)"))
        db.session.commit()
    if "condicion_iva" not in columnas_cliente:
        db.session.execute(
            db.text("ALTER TABLE cliente ADD COLUMN condicion_iva VARCHAR(50) DEFAULT 'Consumidor Final'")
        )
        db.session.commit()
    columnas_usuario = [col[1] for col in db.session.execute(db.text("PRAGMA table_info(usuario)")).fetchall()]
    if "is_admin" not in columnas_usuario:
        db.session.execute(db.text("ALTER TABLE usuario ADD COLUMN is_admin BOOLEAN DEFAULT 0"))
        db.session.commit()
    db.session.execute(db.text("UPDATE usuario SET is_admin = 0 WHERE is_admin IS NULL"))
    db.session.commit()
    if not db.session.query(Usuario.id).filter(Usuario.is_admin.is_(True)).first():
        primer_usuario = Usuario.query.order_by(Usuario.id.asc()).first()
        if primer_usuario:
            primer_usuario.is_admin = True
            db.session.commit()
    cotizaciones_ordenadas = Cotizacion.query.order_by(Cotizacion.fecha.asc(), Cotizacion.id.asc()).all()
    secuencias_por_anio = {}
    hubo_cambios_numeracion = False
    for cot in cotizaciones_ordenadas:
        anio = (cot.fecha or datetime.utcnow()).year
        secuencias_por_anio[anio] = secuencias_por_anio.get(anio, 0) + 1
        numero_esperado = formatear_numero_cotizacion(anio, secuencias_por_anio[anio])
        if (cot.numero_cotizacion or "").strip() != numero_esperado:
            cot.numero_cotizacion = numero_esperado
            hubo_cambios_numeracion = True
    if hubo_cambios_numeracion:
        db.session.commit()
    for sql in (
        "CREATE INDEX IF NOT EXISTS ix_cotizacion_fecha_id ON cotizacion (fecha DESC, id DESC)",
        "CREATE INDEX IF NOT EXISTS ix_cotizacion_estado_fecha ON cotizacion (estado, fecha DESC)",
        "CREATE INDEX IF NOT EXISTS ix_cotizacion_moneda_fecha ON cotizacion (moneda, fecha DESC)",
        "CREATE INDEX IF NOT EXISTS ix_cotizacion_familia_fecha ON cotizacion (familia, fecha DESC)",
        "CREATE INDEX IF NOT EXISTS ix_cotizacion_cliente_fecha ON cotizacion (cliente_id, fecha DESC)",
        "CREATE INDEX IF NOT EXISTS ix_cliente_nombre ON cliente (nombre)",
        "CREATE INDEX IF NOT EXISTS ix_cliente_sector ON cliente (sector)",
        "CREATE INDEX IF NOT EXISTS ix_cliente_subsector ON cliente (subsector)",
        "CREATE INDEX IF NOT EXISTS ix_auditoria_fecha_id ON auditoria (fecha DESC, id DESC)",
        "CREATE INDEX IF NOT EXISTS ix_item_cotizacion_cotizacion_id ON item_cotizacion (cotizacion_id)",
    ):
        db.session.execute(db.text(sql))
    db.session.commit()


def archivo_imagen_permitido(filename):
    return Path(filename or "").suffix.lower() in ALLOWED_IMAGE_EXTENSIONS


def optimizar_bytes_imagen(raw_bytes):
    if not raw_bytes:
        raise ValueError("La imagen subida esta vacia.")
    if len(raw_bytes) > MAX_IMAGE_UPLOAD_BYTES:
        raise ValueError("Cada imagen debe pesar como maximo 6 MB antes de subirse.")

    try:
        image = Image.open(BytesIO(raw_bytes))
        image = ImageOps.exif_transpose(image)
        if getattr(image, "is_animated", False):
            image.seek(0)
    except (UnidentifiedImageError, OSError):
        raise ValueError("El archivo seleccionado no es una imagen valida.")

    if image.mode in ("RGBA", "LA", "P"):
        flattened = Image.new("RGB", image.size, (255, 255, 255))
        alpha = image.getchannel("A") if "A" in image.getbands() else None
        flattened.paste(image.convert("RGBA"), mask=alpha)
        image = flattened
    elif image.mode != "RGB":
        image = image.convert("RGB")

    image.thumbnail((PRODUCT_IMAGE_MAX_DIMENSION, PRODUCT_IMAGE_MAX_DIMENSION), Image.Resampling.LANCZOS)

    optimized_bytes = None
    for quality in PRODUCT_IMAGE_QUALITY_STEPS:
        buffer = BytesIO()
        image.save(buffer, format="JPEG", quality=quality, optimize=True, progressive=True)
        optimized_bytes = buffer.getvalue()
        if len(optimized_bytes) <= PRODUCT_IMAGE_TARGET_BYTES:
            break

    return optimized_bytes


def guardar_imagen_producto(file_storage, descripcion, row_id):
    if not file_storage or not file_storage.filename:
        return None
    if not archivo_imagen_permitido(file_storage.filename):
        raise ValueError("Formato de imagen no permitido. Usa PNG, JPG, WEBP o GIF.")

    base = secure_filename(descripcion) or "producto"
    nombre_archivo = f"{datetime.utcnow():%Y%m%d%H%M%S%f}_{secure_filename(str(row_id))}_{base}.jpg"
    destino = UPLOADS_PRODUCTOS_DIR / nombre_archivo
    file_storage.stream.seek(0)
    optimized_bytes = optimizar_bytes_imagen(file_storage.read())
    destino.write_bytes(optimized_bytes)
    return f"uploads/productos/{nombre_archivo}"


def es_ruta_imagen_local(ruta):
    return bool(ruta) and not str(ruta).startswith(("http://", "https://"))


def eliminar_imagen_local(ruta):
    if not es_ruta_imagen_local(ruta):
        return
    archivo = Path(app.static_folder) / str(ruta)
    if archivo.exists():
        archivo.unlink()


def generar_numero_cotizacion(fecha_referencia=None):
    fecha_base = fecha_referencia or datetime.utcnow()
    inicio_anio = datetime(fecha_base.year, 1, 1)
    inicio_anio_siguiente = datetime(fecha_base.year + 1, 1, 1)
    ultima_secuencia = 0
    for (numero_raw,) in db.session.query(Cotizacion.numero_cotizacion).filter(
        Cotizacion.fecha >= inicio_anio,
        Cotizacion.fecha < inicio_anio_siguiente,
    ).all():
        numero_parseado = parsear_numero_cotizacion(numero_raw)
        if not numero_parseado:
            continue
        anio, secuencia = numero_parseado
        if anio == fecha_base.year and secuencia > ultima_secuencia:
            ultima_secuencia = secuencia
    return formatear_numero_cotizacion(fecha_base.year, ultima_secuencia + 1)


def obtener_config_smtp():
    config_local = {}
    if LOCAL_SETTINGS_PATH.exists():
        try:
            config_local = json.loads(LOCAL_SETTINGS_PATH.read_text(encoding="utf-8"))
            if not isinstance(config_local, dict):
                config_local = {}
        except Exception as exc:
            print(f"[config] No se pudo leer {LOCAL_SETTINGS_PATH.name}: {exc}")
            config_local = {}

    def cfg(nombre, default=""):
        valor = os.getenv(nombre)
        if valor not in (None, ""):
            return valor
        valor_local = config_local.get(nombre)
        if valor_local not in (None, ""):
            return valor_local
        return default

    port_raw = str(cfg("SMTP_PORT", DEFAULT_SMTP_PORT)).strip()
    try:
        port = int(port_raw)
    except ValueError:
        port = DEFAULT_SMTP_PORT
    use_ssl = str(cfg("SMTP_USE_SSL", "")).strip().lower() in ("1", "true", "yes")
    if not use_ssl and port == 465:
        use_ssl = True
    use_tls = str(cfg("SMTP_USE_TLS", "")).strip().lower() in ("1", "true", "yes")
    if not use_tls and not use_ssl and port == 587:
        use_tls = True
    return {
        "host": str(cfg("SMTP_HOST", DEFAULT_SMTP_HOST)).strip(),
        "port": port,
        "username": str(cfg("SMTP_USERNAME", DEFAULT_SMTP_USERNAME)).strip(),
        "password": str(cfg("SMTP_PASSWORD", "") or ""),
        "from_email": str(cfg("SMTP_FROM", cfg("SMTP_USERNAME", DEFAULT_SMTP_FROM)) or DEFAULT_SMTP_FROM).strip(),
        "use_ssl": use_ssl,
        "use_tls": use_tls,
        "base_url": str(cfg("APP_BASE_URL", DEFAULT_APP_BASE_URL)).strip().rstrip("/"),
        "default_to": str(cfg("FOLLOWUP_DEFAULT_TO_EMAIL", DEFAULT_FOLLOWUP_EMAIL)).strip(),
    }


def smtp_esta_configurado():
    config = obtener_config_smtp()
    return bool(config["host"] and config["from_email"] and config["username"] and config["password"])


def auth_cookie_kwargs():
    return {
        "httponly": True,
        "samesite": "Lax",
        "secure": AUTH_COOKIE_SECURE or request.is_secure,
        "max_age": 60 * 60 * 24,
    }


def no_hay_usuarios():
    return Usuario.query.count() == 0


def obtener_admin_setup_token():
    return str(os.getenv("ADMIN_SETUP_TOKEN") or LOCAL_SETTINGS.get("ADMIN_SETUP_TOKEN") or "").strip()


def buscar_usuario_por_username(username):
    username = (username or "").strip()
    if not username:
        return None
    return Usuario.query.filter(func.lower(Usuario.username) == username.lower()).first()


def registrar_auditoria(accion, tipo_entidad, entidad_id=None, entidad_ref=None, detalle=None, usuario=None, username=None):
    usuario_actual = usuario or getattr(g, "current_user", None)
    username_final = username or (usuario_actual.username if usuario_actual else "sistema")
    usuario_id = usuario_actual.id if usuario_actual else None

    try:
        db.session.add(
            Auditoria(
                usuario_id=usuario_id,
                username=username_final,
                accion=(accion or "").strip(),
                tipo_entidad=(tipo_entidad or "").strip(),
                entidad_id=entidad_id,
                entidad_ref=(entidad_ref or "").strip() or None,
                detalle=(detalle or "").strip() or None,
            )
        )
        db.session.commit()
    except Exception as exc:
        db.session.rollback()
        print(f"[auditoria] No se pudo registrar la accion '{accion}': {exc}")


@app.errorhandler(RequestEntityTooLarge)
def handle_request_entity_too_large(_error):
    flash("La carga total de archivos excede el maximo permitido para una sola cotizacion.", "danger")
    destino = request.referrer or url_for("index")
    return redirect(destino)


def normalizar_estado_cotizacion(valor):
    valor = (valor or "").strip()
    if not valor:
        return "En progreso"
    for estado in ESTADOS_COTIZACION:
        if valor.lower() == estado.lower():
            return estado
    return None


def normalizar_familia_cotizacion(valor):
    valor = (valor or "").strip()
    if not valor:
        return None
    for familia in FAMILIAS_COTIZACION:
        if valor.lower() == familia.lower():
            return familia
    return None


def normalizar_subsector_dashboard(valor):
    valor = (valor or "").strip()
    if not valor:
        return None
    for subsectores in SECTORES_CLIENTE.values():
        for subsector in subsectores:
            if valor.lower() == subsector.lower():
                return subsector
    return None


def normalizar_sector_cliente(valor):
    valor = (valor or "").strip()
    for sector in SECTORES_CLIENTE:
        if valor.lower() == sector.lower():
            return sector
    return None


def normalizar_subsector_cliente(sector, valor):
    sector_normalizado = normalizar_sector_cliente(sector)
    if not sector_normalizado:
        return None
    valor = (valor or "").strip()
    for subsector in SECTORES_CLIENTE[sector_normalizado]:
        if valor.lower() == subsector.lower():
            return subsector
    return None


def validar_payload_cliente(data, cliente_actual=None):
    nombre = (data.get("nombre") or "").strip()
    cuit = (data.get("cuit") or "").strip()
    sector = normalizar_sector_cliente(data.get("sector"))
    subsector = normalizar_subsector_cliente(sector, data.get("subsector"))

    if not nombre:
        return None, "Nombre requerido"
    if not sector:
        return None, "Sector requerido"
    if not subsector:
        return None, "Subsector invalido para el sector seleccionado"

    if cuit:
        query = Cliente.query.filter(Cliente.cuit == cuit)
        if cliente_actual:
            query = query.filter(Cliente.id != cliente_actual.id)
        repetido = query.first()
        if repetido:
            return None, "Ya existe un cliente con ese CUIT"

    payload = {
        "nombre": nombre,
        "razon_social": (data.get("razon_social") or "").strip(),
        "cuit": cuit,
        "domicilio": (data.get("domicilio") or "").strip(),
        "sector": sector,
        "subsector": subsector,
        "email": (data.get("email") or "").strip(),
        "telefono": (data.get("telefono") or "").strip(),
        "condicion_iva": (data.get("condicion_iva") or "Consumidor Final").strip() or "Consumidor Final",
    }
    return payload, None


def parsear_entero_positivo(valor, default=None):
    if valor is None:
        valor = ""
    valor = str(valor).strip()
    if not valor:
        return default
    try:
        numero = int(valor)
    except ValueError:
        return default
    return numero if numero > 0 else default


def paginar_query(query, page=1, per_page=20):
    page = max(parsear_entero_positivo(page, default=1) or 1, 1)
    per_page = max(parsear_entero_positivo(per_page, default=20) or 20, 1)
    total = query.order_by(None).count()
    pages = max((total + per_page - 1) // per_page, 1)
    if page > pages:
        page = pages
    items = query.limit(per_page).offset((page - 1) * per_page).all()
    start = ((page - 1) * per_page) + 1 if total else 0
    end = min(page * per_page, total) if total else 0
    return {
        "items": items,
        "page": page,
        "per_page": per_page,
        "total": total,
        "pages": pages,
        "has_prev": page > 1,
        "has_next": page < pages,
        "prev_page": page - 1 if page > 1 else None,
        "next_page": page + 1 if page < pages else None,
        "start": start,
        "end": end,
    }


def construir_paginas_visibles(page, pages, radius=2):
    if pages <= 1:
        return [1]
    inicio = max(page - radius, 1)
    fin = min(page + radius, pages)
    paginas = list(range(inicio, fin + 1))
    if 1 not in paginas:
        paginas.insert(0, 1)
    if pages not in paginas:
        paginas.append(pages)
    resultado = []
    anterior = None
    for numero in paginas:
        if anterior and numero - anterior > 1:
            resultado.append(None)
        resultado.append(numero)
        anterior = numero
    return resultado


def normalizar_followup_email(valor, cliente_sel=None):
    config = obtener_config_smtp()
    return config["default_to"] or DEFAULT_FOLLOWUP_EMAIL


def construir_link_cotizacion(cotizacion, editar=False):
    config = obtener_config_smtp()
    sufijo = f"/cotizacion/{cotizacion.id}/editar" if editar else f"/cotizacion/{cotizacion.id}"
    return f"{config['base_url']}{sufijo}"


def preparar_seguimiento_cotizacion(cotizacion, cliente_sel=None):
    activo_anterior = bool(cotizacion.seguimiento_activo)
    email_anterior = (cotizacion.seguimiento_email or "").strip()
    dias_anterior = cotizacion.seguimiento_cada_dias
    activo = request.form.get("seguimiento_activo") == "1"
    email_raw = (request.form.get("seguimiento_email") or "").strip()
    dias = parsear_entero_positivo(request.form.get("seguimiento_cada_dias"), default=None)
    email = normalizar_followup_email(email_raw, cliente_sel=cliente_sel)

    if not activo:
        cotizacion.seguimiento_activo = False
        cotizacion.seguimiento_email = email or None
        cotizacion.seguimiento_cada_dias = dias
        cotizacion.seguimiento_proximo_envio = None
        return

    if not dias:
        dias = 7

    cotizacion.seguimiento_activo = True
    cotizacion.seguimiento_email = email or None
    cotizacion.seguimiento_cada_dias = dias
    if (
        not cotizacion.seguimiento_proximo_envio
        or not activo_anterior
        or email_anterior != (email or "")
        or dias_anterior != dias
    ):
        cotizacion.seguimiento_proximo_envio = datetime.utcnow() + timedelta(days=dias)


def enviar_mail_recordatorio(cotizacion):
    config = obtener_config_smtp()
    destinatario = (cotizacion.seguimiento_email or "").strip()
    if not smtp_esta_configurado() or not destinatario:
        return False

    asunto = f"Seguimiento cotizacion #{cotizacion.numero_cotizacion or cotizacion.id}"
    link_ver = construir_link_cotizacion(cotizacion, editar=False)
    link_editar = construir_link_cotizacion(cotizacion, editar=True)
    cuerpo = f"""Hola,

Hay alguna novedad con esta cotizacion?

Cotizacion: #{cotizacion.numero_cotizacion or cotizacion.id}
Cliente: {cotizacion.cliente}
Estado actual: {cotizacion.estado or 'En progreso'}
Total: {cotizacion.moneda or 'ARS'} {cotizacion.total_final or 0:.2f}

Ver cotizacion:
{link_ver}

Editar cotizacion / cambiar estado:
{link_editar}

Este recordatorio se envio automaticamente desde Cuenco Tech.
"""

    mensaje = EmailMessage()
    mensaje["Subject"] = asunto
    mensaje["From"] = config["from_email"]
    mensaje["To"] = destinatario
    mensaje.set_content(cuerpo)

    smtp_class = smtplib.SMTP_SSL if config["use_ssl"] else smtplib.SMTP
    with smtp_class(config["host"], config["port"], timeout=20) as smtp:
        if config["use_tls"] and not config["use_ssl"]:
            smtp.starttls()
        smtp.login(config["username"], config["password"])
        smtp.send_message(mensaje)
    return True


def procesar_recordatorios_vencidos():
    ahora = datetime.utcnow()
    cotizaciones = Cotizacion.query.filter(
        Cotizacion.seguimiento_activo.is_(True),
        Cotizacion.seguimiento_proximo_envio.isnot(None),
        Cotizacion.seguimiento_proximo_envio <= ahora,
    ).all()

    for cotizacion in cotizaciones:
        if normalizar_estado_cotizacion(cotizacion.estado) != "En progreso":
            continue
        if not cotizacion.seguimiento_email:
            continue
        try:
            enviado = enviar_mail_recordatorio(cotizacion)
        except Exception as exc:
            print(f"[seguimiento] Error enviando cotizacion {cotizacion.id}: {exc}")
            continue
        if not enviado:
            continue

        dias = cotizacion.seguimiento_cada_dias or 7
        ahora_envio = datetime.utcnow()
        cotizacion.seguimiento_ultimo_envio = ahora_envio
        cotizacion.seguimiento_proximo_envio = ahora_envio + timedelta(days=dias)
        db.session.add(cotizacion)

    if cotizaciones:
        db.session.commit()


def worker_recordatorios():
    while True:
        try:
            with app.app_context():
                procesar_recordatorios_vencidos()
        except Exception as exc:
            print(f"[seguimiento] Worker error: {exc}")
        time.sleep(REMINDER_POLL_SECONDS)


def iniciar_worker_recordatorios():
    global _reminder_worker_started
    if _reminder_worker_started:
        return
    with _reminder_worker_lock:
        if _reminder_worker_started:
            return
        thread = threading.Thread(target=worker_recordatorios, name="cotizador-followup-worker", daemon=True)
        thread.start()
        _reminder_worker_started = True


def construir_items_precargados(cotizacion):
    items_precargados = []
    if not cotizacion:
        return items_precargados

    for item in cotizacion.items:
        if item.imagen_url:
            preview_url = item.imagen_url if str(item.imagen_url).startswith(("http://", "https://")) else url_for(
                "static", filename=item.imagen_url
            )
        else:
            preview_url = url_for("static", filename=PLACEHOLDER_PRODUCTO)

        items_precargados.append(
            {
                "item_id": item.id,
                "descripcion": item.descripcion or "",
                "cantidad": item.cantidad or 0,
                "costo_unitario": item.costo_unitario or 0,
                "costo_extra": item.costo_extra or 0,
                "margen": item.margen or 0,
                "iva_item": item.iva_item or 0,
                "imagen_url": item.imagen_url or "",
                "preview_url": preview_url,
            }
        )
    return items_precargados


def renderizar_cotizador(cotizacion=None):
    clientes = Cliente.query.order_by(Cliente.nombre).all()
    return render_template(
        "cotizador.html",
        clientes=clientes,
        cotizacion=cotizacion,
        modo_edicion=bool(cotizacion),
        items_precargados=construir_items_precargados(cotizacion),
        familias_cotizacion=FAMILIAS_COTIZACION,
        smtp_configurado=smtp_esta_configurado(),
        followup_default_email=obtener_config_smtp()["default_to"],
    )


def generar_excel_cotizacion(cotizacion):
    wb = Workbook()
    ws = wb.active
    ws.title = "Cotizacion"

    dark_fill = PatternFill("solid", fgColor="0F172A")
    mid_fill = PatternFill("solid", fgColor="E2E8F0")
    total_fill = PatternFill("solid", fgColor="DBEAFE")
    white_font = Font(color="FFFFFF", bold=True)
    bold_font = Font(bold=True)
    title_font = Font(size=15, bold=True)
    thin_side = Side(style="thin", color="CBD5E1")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    def style_row(row_idx, fill=None, font=None, alignment=None):
        for cell in ws[row_idx]:
            if fill:
                cell.fill = fill
            if font:
                cell.font = font
            if alignment:
                cell.alignment = alignment
            cell.border = border

    def write_label_value(row_idx, label, value):
        ws.cell(row=row_idx, column=1, value=label)
        ws.cell(row=row_idx, column=2, value=value)
        ws.cell(row=row_idx, column=1).font = bold_font
        ws.cell(row=row_idx, column=1).fill = mid_fill
        ws.cell(row=row_idx, column=1).border = border
        ws.cell(row=row_idx, column=2).border = border

    ws.merge_cells("A1:H1")
    ws["A1"] = f"Cotizacion {cotizacion.numero_cotizacion or cotizacion.id}"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A1"].fill = dark_fill
    ws["A1"].font = Font(color="FFFFFF", bold=True, size=15)
    ws["A1"].border = border

    row = 3
    datos_generales = [
        ("Empresa", cotizacion.nombre_fantasia or NOMBRE_FANTASIA),
        ("Razon social empresa", cotizacion.razon_social or RAZON_SOCIAL),
        ("CUIT empresa", cotizacion.cuit or CUIT),
        ("Domicilio", DOMICILIO),
        ("Numero de cotizacion", cotizacion.numero_cotizacion or cotizacion.id),
        ("Fecha de creacion", cotizacion.fecha.strftime("%d/%m/%Y %H:%M") if cotizacion.fecha else ""),
        ("Estado", cotizacion.estado or "En progreso"),
        ("Familia", cotizacion.familia or ""),
        ("Cliente", cotizacion.cliente or ""),
        ("Razon social cliente", cotizacion.cliente_razon_social or ""),
        ("CUIT cliente", cotizacion.cliente_cuit or ""),
        ("Moneda", cotizacion.moneda or "ARS"),
        ("Condicion IVA", cotizacion.condicion_iva or ""),
    ]
    for label, value in datos_generales:
        write_label_value(row, label, value)
        row += 1

    row += 1
    headers = [
        "Descripcion",
        "Cantidad",
        "Costo unitario",
        "Costo extra %",
        "Margen",
        "IVA %",
        "Precio venta unitario",
        "Subtotal",
    ]
    for col_idx, header in enumerate(headers, 1):
        ws.cell(row=row, column=col_idx, value=header)
    style_row(row, fill=dark_fill, font=white_font, alignment=Alignment(horizontal="center"))

    money_columns = {3, 7, 8}
    percentage_columns = {4, 5, 6}
    for item in cotizacion.items:
        row += 1
        values = [
            item.descripcion or "",
            item.cantidad or 0,
            item.costo_unitario or 0,
            item.costo_extra or 0,
            item.margen or 0,
            item.iva_item or 0,
            item.precio_venta or 0,
            item.subtotal or 0,
        ]
        for col_idx, value in enumerate(values, 1):
            cell = ws.cell(row=row, column=col_idx, value=value)
            cell.border = border
            if col_idx in money_columns:
                cell.number_format = '#,##0.00'
            elif col_idx in percentage_columns:
                cell.number_format = '0.00'

    row += 2
    resumen_inicio = row
    resumen = [
        ("Subtotal neto", cotizacion.total_neto or 0),
        ("IVA total", cotizacion.total_iva or 0),
        ("Total final", cotizacion.total_final or 0),
    ]
    for idx, (label, value) in enumerate(resumen):
        r = resumen_inicio + idx
        ws.cell(row=r, column=6, value=label)
        ws.cell(row=r, column=7, value=value)
        ws.cell(row=r, column=6).font = bold_font
        ws.cell(row=r, column=6).fill = total_fill
        ws.cell(row=r, column=6).border = border
        ws.cell(row=r, column=7).border = border
        ws.cell(row=r, column=7).number_format = '#,##0.00'
        if label == "Total final":
            ws.cell(row=r, column=6).font = Font(bold=True, color="0F172A")
            ws.cell(row=r, column=7).font = Font(bold=True, color="0F172A")

    for column, width in {
        "A": 34,
        "B": 28,
        "C": 14,
        "D": 14,
        "E": 12,
        "F": 12,
        "G": 18,
        "H": 16,
    }.items():
        ws.column_dimensions[column].width = width

    for current_row in ws.iter_rows():
        for cell in current_row:
            if cell.row == 1:
                continue
            if not cell.alignment:
                cell.alignment = Alignment(vertical="center")

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def aplicar_filtro_fecha_cotizaciones(query, desde="", hasta=""):
    if desde:
        query = query.filter(func.date(Cotizacion.fecha) >= desde)
    if hasta:
        query = query.filter(func.date(Cotizacion.fecha) <= hasta)
    return query


def parsear_fecha_iso(valor):
    valor = (valor or "").strip()
    if not valor:
        return None
    try:
        return datetime.strptime(valor, "%Y-%m-%d").date()
    except ValueError:
        return None


def resolver_rango_dashboard(periodo, desde_raw="", hasta_raw=""):
    hoy = datetime.now().date()
    periodo = (periodo or "30").strip().lower()
    rangos_rapidos = {
        "7": 6,
        "30": 29,
        "60": 59,
        "90": 89,
        "365": 364,
    }

    if periodo in rangos_rapidos:
        hasta_date = hoy
        desde_date = hoy - timedelta(days=rangos_rapidos[periodo])
    else:
        periodo = "custom"
        desde_date = parsear_fecha_iso(desde_raw)
        hasta_date = parsear_fecha_iso(hasta_raw) or hoy
        if desde_date is None:
            desde_date = hasta_date - timedelta(days=29)

    if desde_date > hasta_date:
        desde_date, hasta_date = hasta_date, desde_date

    return periodo, desde_date.isoformat(), hasta_date.isoformat(), desde_date, hasta_date


def calcular_variacion_porcentual(actual, anterior):
    if not anterior:
        return None if not actual else 100.0
    return round(((actual - anterior) / anterior) * 100.0, 1)


def construir_series_dashboard(cotizaciones, desde_date, hasta_date):
    total_dias = max((hasta_date - desde_date).days + 1, 1)
    usar_semanas = total_dias > 62
    labels = []
    buckets = []

    cursor = desde_date
    while cursor <= hasta_date:
        if usar_semanas:
            bucket_start = cursor
            bucket_end = min(cursor + timedelta(days=6), hasta_date)
            labels.append(f"{bucket_start.strftime('%d/%m')} - {bucket_end.strftime('%d/%m')}")
            buckets.append({"start": bucket_start, "end": bucket_end, "creadas": 0, "cerradas": 0, "aceptadas": 0})
            cursor = bucket_end + timedelta(days=1)
        else:
            labels.append(cursor.strftime("%d/%m"))
            buckets.append({"start": cursor, "end": cursor, "creadas": 0, "cerradas": 0, "aceptadas": 0})
            cursor += timedelta(days=1)

    for cotizacion in cotizaciones:
        fecha_cot = (cotizacion.fecha or datetime.utcnow()).date()
        if fecha_cot < desde_date or fecha_cot > hasta_date:
            continue
        if usar_semanas:
            index = (fecha_cot - desde_date).days // 7
        else:
            index = (fecha_cot - desde_date).days
        if index < 0 or index >= len(buckets):
            continue

        bucket = buckets[index]
        estado = normalizar_estado_cotizacion(cotizacion.estado) or "En progreso"
        bucket["creadas"] += 1
        if estado in ("Aceptada", "Rechazada"):
            bucket["cerradas"] += 1
        if estado == "Aceptada":
            bucket["aceptadas"] += 1

    return {
        "labels": labels,
        "creadas": [bucket["creadas"] for bucket in buckets],
        "cerradas": [bucket["cerradas"] for bucket in buckets],
        "aceptadas": [bucket["aceptadas"] for bucket in buckets],
        "granularidad": "Semanal" if usar_semanas else "Diaria",
    }


def construir_top_clientes_dashboard(cotizaciones, limite=5):
    acumulado = defaultdict(lambda: {"cantidad": 0, "total": 0.0, "aceptadas": 0, "total_ars": 0.0, "total_usd": 0.0})
    for cotizacion in cotizaciones:
        nombre = (cotizacion.cliente_ref.nombre if cotizacion.cliente_ref else cotizacion.cliente) or "Sin cliente"
        estado = normalizar_estado_cotizacion(cotizacion.estado) or "En progreso"
        moneda = (cotizacion.moneda or "ARS").upper()
        total_final = cotizacion.total_final or 0.0
        acumulado[nombre]["cantidad"] += 1
        acumulado[nombre]["total"] += total_final
        if moneda == "USD":
            acumulado[nombre]["total_usd"] += total_final
        else:
            acumulado[nombre]["total_ars"] += total_final
        if estado == "Aceptada":
            acumulado[nombre]["aceptadas"] += 1

    ranking = []
    for nombre, data in acumulado.items():
        ranking.append(
            {
                "nombre": nombre,
                "cantidad": data["cantidad"],
                "aceptadas": data["aceptadas"],
                "total": round(data["total"], 2),
                "total_ars": round(data["total_ars"], 2),
                "total_usd": round(data["total_usd"], 2),
            }
        )

    ranking.sort(key=lambda item: (-item["cantidad"], -item["total"], item["nombre"].lower()))
    return ranking[:limite]


def construir_desglose_dashboard(cotizaciones, key_fn, default_label="Sin definir", limite=None):
    acumulado = defaultdict(
        lambda: {"cantidad": 0, "aceptadas": 0, "total": 0.0, "total_ars": 0.0, "total_usd": 0.0}
    )
    total_cotizaciones = len(cotizaciones)

    for cotizacion in cotizaciones:
        nombre = (key_fn(cotizacion) or "").strip() or default_label
        estado = normalizar_estado_cotizacion(cotizacion.estado) or "En progreso"
        moneda = (cotizacion.moneda or "ARS").upper()
        total_final = cotizacion.total_final or 0.0
        acumulado[nombre]["cantidad"] += 1
        acumulado[nombre]["total"] += total_final
        if moneda == "USD":
            acumulado[nombre]["total_usd"] += total_final
        else:
            acumulado[nombre]["total_ars"] += total_final
        if estado == "Aceptada":
            acumulado[nombre]["aceptadas"] += 1

    desglose = []
    for nombre, data in acumulado.items():
        cantidad = data["cantidad"]
        desglose.append(
            {
                "nombre": nombre,
                "cantidad": cantidad,
                "aceptadas": data["aceptadas"],
                "total": round(data["total"], 2),
                "total_ars": round(data["total_ars"], 2),
                "total_usd": round(data["total_usd"], 2),
                "porcentaje": round((cantidad / total_cotizaciones) * 100.0, 1) if total_cotizaciones else 0.0,
            }
        )

    desglose.sort(key=lambda item: (-item["cantidad"], -item["total"], item["nombre"].lower()))
    if limite:
        return desglose[:limite]
    return desglose


def aplicar_filtro_estado_dashboard(query, estado):
    estado = (estado or "todos").strip().lower()
    if estado == "cerradas":
        estado = "aceptadas"
    if estado == "aceptadas":
        return query.filter(Cotizacion.estado == "Aceptada")
    if estado == "rechazadas":
        return query.filter(Cotizacion.estado == "Rechazada")
    if estado == "en_progreso":
        return query.filter(Cotizacion.estado == "En progreso")
    return query


def aplicar_filtros_segmentacion_dashboard(query, familia="", sector="", subsector="", moneda=""):
    if familia:
        query = query.filter(Cotizacion.familia == familia)
    if sector:
        query = query.filter(Cliente.sector == sector)
    if subsector:
        query = query.filter(Cliente.subsector == subsector)
    if moneda in ("ARS", "USD"):
        query = query.filter(Cotizacion.moneda == moneda)
    return query


def aplicar_filtro_cliente_dashboard(query, cliente=""):
    cliente = (cliente or "").strip()
    if not cliente:
        return query
    patron = f"%{cliente}%"
    return query.filter(
        or_(
            Cotizacion.cliente.ilike(patron),
            Cotizacion.cliente_razon_social.ilike(patron),
            Cotizacion.cliente_cuit.ilike(patron),
            Cliente.nombre.ilike(patron),
            Cliente.razon_social.ilike(patron),
            Cliente.cuit.ilike(patron),
        )
    )


def resolver_estado_dashboard(valor):
    estado = (valor or "todos").strip().lower()
    if estado == "cerradas":
        estado = "aceptadas"
    if estado not in ("todos", "en_progreso", "aceptadas", "rechazadas"):
        estado = "todos"
    return estado


def resolver_filtros_dashboard_request():
    estado = resolver_estado_dashboard(request.args.get("estado"))
    familia = normalizar_familia_cotizacion(request.args.get("familia"))
    sector = normalizar_sector_cliente(request.args.get("sector"))
    subsector = (
        normalizar_subsector_cliente(sector, request.args.get("subsector"))
        if sector
        else normalizar_subsector_dashboard(request.args.get("subsector"))
    )
    moneda = (request.args.get("moneda") or "").strip().upper()
    if moneda not in ("ARS", "USD"):
        moneda = ""
    cliente = (request.args.get("cliente") or "").strip()

    periodo, desde, hasta, desde_date, hasta_date = resolver_rango_dashboard(
        request.args.get("periodo"),
        request.args.get("desde"),
        request.args.get("hasta"),
    )

    return {
        "estado": estado,
        "familia": familia,
        "sector": sector,
        "subsector": subsector,
        "moneda": moneda,
        "cliente": cliente,
        "periodo": periodo,
        "desde": desde,
        "hasta": hasta,
        "desde_date": desde_date,
        "hasta_date": hasta_date,
        "op_page": parsear_entero_positivo(request.args.get("op_page"), default=1) or 1,
    }


def construir_query_dashboard_periodo(filtros):
    query_periodo = Cotizacion.query.outerjoin(Cliente, Cotizacion.cliente_id == Cliente.id)
    query_periodo = aplicar_filtro_fecha_cotizaciones(query_periodo, filtros["desde"], filtros["hasta"])
    query_periodo = aplicar_filtro_cliente_dashboard(query_periodo, filtros["cliente"])
    query_periodo = aplicar_filtros_segmentacion_dashboard(
        query_periodo,
        familia=filtros["familia"],
        sector=filtros["sector"],
        subsector=filtros["subsector"],
        moneda=filtros["moneda"],
    )
    return query_periodo


def construir_contexto_dashboard_operativo(query_periodo, filtros):
    query_tabla = aplicar_filtro_estado_dashboard(query_periodo, filtros["estado"])
    paginacion = paginar_query(
        query_tabla.order_by(Cotizacion.fecha.desc(), Cotizacion.id.desc()),
        page=filtros.get("op_page", 1),
        per_page=DASHBOARD_OPERATIVO_PER_PAGE,
    )
    return {
        "cotizaciones": paginacion["items"],
        "paginacion_operativa": {
            **paginacion,
            "pages_visible": construir_paginas_visibles(paginacion["page"], paginacion["pages"]),
        },
        "filtro_estado": filtros["estado"],
        "filtro_periodo": filtros["periodo"],
        "filtro_desde": filtros["desde"],
        "filtro_hasta": filtros["hasta"],
        "filtro_familia": filtros["familia"] or "",
        "filtro_sector": filtros["sector"] or "",
        "filtro_subsector": filtros["subsector"] or "",
        "filtro_moneda": filtros["moneda"] or "",
        "filtro_cliente": filtros["cliente"] or "",
        "filtro_op_page": paginacion["page"],
    }


def persistir_cotizacion_desde_form(cotizacion=None):
    cliente_id_raw = (request.form.get("cliente_id") or "").strip()
    cliente_sel = None
    if cliente_id_raw.isdigit():
        cliente_sel = db.session.get(Cliente, int(cliente_id_raw))

    moneda = (request.form.get("moneda") or "ARS").upper()
    if moneda not in ("ARS", "USD"):
        moneda = "ARS"

    condicion_iva = request.form.get("condicion_iva") or "Consumidor Final"
    if condicion_iva not in ("Exento", "Consumidor Final", "Responsable Inscrito"):
        condicion_iva = "Consumidor Final"
    if cliente_sel and cliente_sel.condicion_iva in ("Exento", "Consumidor Final", "Responsable Inscrito"):
        condicion_iva = cliente_sel.condicion_iva

    es_edicion = cotizacion is not None
    estado_anterior = None
    familia_anterior = normalizar_familia_cotizacion(cotizacion.familia if cotizacion else None)
    if not cotizacion:
        fecha_creacion = datetime.utcnow()
        cotizacion = Cotizacion(
            fecha=fecha_creacion,
            numero_cotizacion=generar_numero_cotizacion(fecha_creacion),
            estado="En progreso",
            nombre_fantasia=NOMBRE_FANTASIA,
            razon_social=RAZON_SOCIAL,
            cuit=CUIT,
            total_neto=0.0,
            total_iva=0.0,
            total_final=0.0,
        )
    else:
        estado_anterior = normalizar_estado_cotizacion(cotizacion.estado) or "En progreso"
        cotizacion.numero_cotizacion = cotizacion.numero_cotizacion or generar_numero_cotizacion(cotizacion.fecha)
        cotizacion.estado = normalizar_estado_cotizacion(cotizacion.estado) or "En progreso"
        estado_form = normalizar_estado_cotizacion(request.form.get("estado"))
        if estado_form:
            cotizacion.estado = estado_form

    cotizacion.nombre_fantasia = NOMBRE_FANTASIA
    cotizacion.razon_social = RAZON_SOCIAL
    cotizacion.cuit = CUIT
    cotizacion.cliente_id = cliente_sel.id if cliente_sel else None
    cotizacion.cliente = cliente_sel.nombre if cliente_sel else (request.form.get("cliente") or "").strip()
    cotizacion.cliente_razon_social = (
        cliente_sel.razon_social if cliente_sel else (request.form.get("cliente_razon_social") or "").strip()
    )
    cotizacion.cliente_cuit = cliente_sel.cuit if cliente_sel else (request.form.get("cliente_cuit") or "").strip()
    familia_form = normalizar_familia_cotizacion(request.form.get("familia"))
    cotizacion.familia = familia_form or familia_anterior
    cotizacion.moneda = moneda
    cotizacion.condicion_iva = condicion_iva
    preparar_seguimiento_cotizacion(cotizacion, cliente_sel=cliente_sel)

    descs = request.form.getlist("desc[]")
    row_ids = request.form.getlist("row_id[]")
    item_ids = request.form.getlist("item_id[]")
    imagenes_actuales = request.form.getlist("imagen_actual[]")
    cants = request.form.getlist("cant[]")
    costs = request.form.getlist("costo[]")
    extras = request.form.getlist("extra[]")
    margs = request.form.getlist("margen[]")
    iva_items = request.form.getlist("iva_item[]")

    items_existentes = {str(item.id): item for item in cotizacion.items if item.id}
    neto_total = 0.0
    iva_total = 0.0

    for i, desc in enumerate(descs):
        desc = (desc or "").strip()
        if not desc:
            continue

        try:
            cantidad = float(cants[i] if i < len(cants) else 0)
            costo = float(costs[i] if i < len(costs) else 0)
            extra_pct = float(extras[i] if i < len(extras) else 5.0)
            margen = float(margs[i] if i < len(margs) else 0)
        except ValueError:
            continue

        cantidad = max(0.0, cantidad)
        costo = max(0.0, costo)
        extra_pct = max(0.0, extra_pct)
        margen = max(0.0, margen)

        try:
            iva_pct = float(iva_items[i] if i < len(iva_items) else 21.0)
        except (TypeError, ValueError):
            iva_pct = 21.0
        if iva_pct not in (21.0, 10.5, 0.0):
            iva_pct = 21.0

        costo_con_extra = costo * (1 + (extra_pct / 100.0))
        p_venta_neto = costo_con_extra * (1 + margen)
        sub_neto = cantidad * p_venta_neto
        iva_pct_aplicado = iva_pct if cotizacion.condicion_iva != "Exento" else 0.0
        monto_iva_item = sub_neto * (iva_pct_aplicado / 100.0)

        row_id = row_ids[i] if i < len(row_ids) else str(i)
        item_id = (item_ids[i] if i < len(item_ids) else "").strip()
        imagen_actual = (imagenes_actuales[i] if i < len(imagenes_actuales) else "").strip()
        try:
            imagen_local = guardar_imagen_producto(request.files.get(f"foto_{row_id}"), desc, row_id)
        except ValueError as exc:
            flash(str(exc), "danger")
            destino = "editar_cotizacion" if es_edicion else "index"
            kwargs = {"id": cotizacion.id} if es_edicion else {}
            return None, redirect(url_for(destino, **kwargs))

        item = items_existentes.pop(item_id, None) if item_id else None
        if item and not imagen_actual:
            imagen_actual = item.imagen_url or ""

        imagen_final = imagen_local or imagen_actual or None
        if item and imagen_local and item.imagen_url and item.imagen_url != imagen_local:
            eliminar_imagen_local(item.imagen_url)

        if not item:
            item = ItemCotizacion()
            cotizacion.items.append(item)

        item.descripcion = desc
        item.cantidad = cantidad
        item.costo_unitario = costo
        item.costo_extra = extra_pct
        item.margen = margen
        item.iva_item = iva_pct_aplicado
        item.precio_venta = round(p_venta_neto, 2)
        item.subtotal = round(sub_neto + monto_iva_item, 2)
        item.imagen_url = imagen_final

        neto_total += sub_neto
        iva_total += monto_iva_item

    for item_sobrante in items_existentes.values():
        eliminar_imagen_local(item_sobrante.imagen_url)
        db.session.delete(item_sobrante)

    if not cotizacion.familia:
        flash("Debes definir la familia de la cotizacion antes de guardarla.", "danger")
        destino = "editar_cotizacion" if es_edicion else "index"
        kwargs = {"id": cotizacion.id} if es_edicion else {}
        return None, redirect(url_for(destino, **kwargs))

    if not cotizacion.cliente or not cotizacion.items:
        destino = "editar_cotizacion" if es_edicion else "index"
        kwargs = {"id": cotizacion.id} if es_edicion else {}
        return None, redirect(url_for(destino, **kwargs))

    cotizacion.total_neto = round(neto_total, 2)
    cotizacion.total_iva = round(iva_total, 2)
    cotizacion.total_final = round(sum(item.subtotal for item in cotizacion.items), 2)

    db.session.add(cotizacion)
    db.session.commit()
    numero_ref = cotizacion.numero_cotizacion or str(cotizacion.id)
    registrar_auditoria(
        "Creó cotización" if not es_edicion else "Modificó cotización",
        "Cotización",
        entidad_id=cotizacion.id,
        entidad_ref=numero_ref,
        detalle=(
            f"Cliente: {cotizacion.cliente}. Familia: {cotizacion.familia}. Estado: {cotizacion.estado}. "
            f"Total: {cotizacion.moneda or 'ARS'} {cotizacion.total_final:,.2f}."
        ),
    )
    if es_edicion and estado_anterior and estado_anterior != cotizacion.estado:
        registrar_auditoria(
            "Cambió estado de cotización",
            "Cotización",
            entidad_id=cotizacion.id,
            entidad_ref=numero_ref,
            detalle=f"Estado anterior: {estado_anterior}. Estado nuevo: {cotizacion.estado}.",
        )
    return cotizacion, None


@app.route("/setup-admin", methods=["GET", "POST"])
def setup_admin():
    if not no_hay_usuarios():
        return redirect(url_for("login"))

    error = None
    setup_token_configurado = bool(obtener_admin_setup_token())
    if request.method == "POST":
        setup_token = (request.form.get("setup_token") or "").strip()
        username = (request.form.get("username") or "").strip()
        password = request.form.get("password") or ""
        password_confirm = request.form.get("password_confirm") or ""
        admin_setup_token = obtener_admin_setup_token()

        if not admin_setup_token:
            error = "El alta inicial esta deshabilitada hasta configurar ADMIN_SETUP_TOKEN."
        elif setup_token != admin_setup_token:
            error = "La clave de instalacion es invalida."
        elif not username or not password:
            error = "Usuario y contraseña son obligatorios."
        elif buscar_usuario_por_username(username):
            error = "Ese usuario ya existe."
        elif password != password_confirm:
            error = "Las contraseñas no coinciden."
        elif len(password) < 8:
            error = "La contraseña debe tener al menos 8 caracteres."
        else:
            nuevo = Usuario(username=username, is_admin=True)
            nuevo.set_password(password)
            db.session.add(nuevo)
            db.session.commit()
            registrar_auditoria(
                "Creo usuario inicial",
                "Usuario",
                entidad_id=nuevo.id,
                entidad_ref=nuevo.username,
                detalle="Alta inicial del cotizador como administrador.",
                usuario=nuevo,
                username=nuevo.username,
            )
            return redirect(url_for("login"))

    return render_template("setup_admin.html", error=error, setup_token_configurado=setup_token_configurado)


@app.route("/login", methods=["GET", "POST"])
def login():
    if no_hay_usuarios():
        return redirect(url_for("setup_admin"))

    token = request.cookies.get("x-access-token")
    if token and obtener_usuario_desde_token(token):
        return redirect(url_for("index"))

    error = None
    if request.method == "POST":
        username = (request.form.get("username") or "").strip()
        password = request.form.get("password") or ""
        user = buscar_usuario_por_username(username)
        if user and user.check_password(password):
            token = generar_token_usuario(user)
            destino = request.args.get("next") or request.form.get("next") or url_for("index")
            if not str(destino).startswith("/"):
                destino = url_for("index")
            res = redirect(destino)
            res.set_cookie("x-access-token", token, **auth_cookie_kwargs())
            return res
        error = "Usuario o contraseña incorrectos."

    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    res = redirect(url_for("login"))
    res.delete_cookie("x-access-token")
    return res


@app.route("/usuarios", methods=["GET", "POST"])
@token_required
@admin_required
def usuarios_page():
    error = None
    if request.method == "POST":
        username = (request.form.get("username") or "").strip()
        password = request.form.get("password") or ""
        password_confirm = request.form.get("password_confirm") or ""
        is_admin = request.form.get("is_admin") == "1"

        if not username:
            error = "El usuario es obligatorio."
        elif buscar_usuario_por_username(username):
            error = "Ese usuario ya existe."
        elif not password:
            error = "La contraseña es obligatoria."
        elif password != password_confirm:
            error = "Las contraseñas no coinciden."
        elif len(password) < 8:
            error = "La contraseña debe tener al menos 8 caracteres."
        else:
            nuevo = Usuario(username=username, is_admin=is_admin)
            nuevo.set_password(password)
            db.session.add(nuevo)
            db.session.commit()
            registrar_auditoria(
                "Creo usuario",
                "Usuario",
                entidad_id=nuevo.id,
                entidad_ref=nuevo.username,
                detalle=f"Rol asignado: {'Administrador' if nuevo.is_admin else 'Operador'}.",
            )
            flash(f"Usuario {nuevo.username} creado correctamente.", "success")
            return redirect(url_for("usuarios_page"))

    usuarios = Usuario.query.order_by(Usuario.is_admin.desc(), func.lower(Usuario.username).asc()).all()
    total_usuarios = len(usuarios)
    total_admins = sum(1 for usuario in usuarios if usuario.is_admin)
    return render_template(
        "usuarios.html",
        error=error,
        usuarios=usuarios,
        total_usuarios=total_usuarios,
        total_admins=total_admins,
    )


@app.route("/api/clientes", methods=["POST"])
@token_required
def agregar_cliente():
    data = request.get_json(silent=True) or {}
    payload, error = validar_payload_cliente(data)
    if error:
        return jsonify({"error": error}), 400

    nuevo = Cliente(**payload)
    db.session.add(nuevo)
    db.session.commit()
    return jsonify(
        {
            "id": nuevo.id,
            "nombre": nuevo.nombre,
            "razon_social": nuevo.razon_social,
            "cuit": nuevo.cuit,
            "domicilio": nuevo.domicilio,
            "sector": nuevo.sector,
            "subsector": nuevo.subsector,
            "email": nuevo.email,
            "telefono": nuevo.telefono,
            "condicion_iva": nuevo.condicion_iva,
        }
    ), 201


@app.route("/api/clientes/<int:id>", methods=["PUT"])
@token_required
def actualizar_cliente(id):
    cliente = Cliente.query.get_or_404(id)
    data = request.get_json(silent=True) or {}
    payload, error = validar_payload_cliente(data, cliente_actual=cliente)
    if error:
        return jsonify({"error": error}), 400

    for campo, valor in payload.items():
        setattr(cliente, campo, valor)

    db.session.add(cliente)
    db.session.commit()
    return jsonify(
        {
            "id": cliente.id,
            "nombre": cliente.nombre,
            "razon_social": cliente.razon_social,
            "cuit": cliente.cuit,
            "domicilio": cliente.domicilio,
            "sector": cliente.sector,
            "subsector": cliente.subsector,
            "email": cliente.email,
            "telefono": cliente.telefono,
            "condicion_iva": cliente.condicion_iva,
        }
    ), 200


@app.before_request
def ensure_followup_worker():
    if app.config.get("TESTING"):
        return
    if app.debug and os.environ.get("WERKZEUG_RUN_MAIN") not in (None, "true"):
        return
    iniciar_worker_recordatorios()


@app.route("/cotizacion/<int:id>/estado", methods=["POST"])
@token_required
def actualizar_estado_cotizacion(id):
    cotizacion = Cotizacion.query.get_or_404(id)
    data = request.get_json(silent=True) or {}
    estado = normalizar_estado_cotizacion(data.get("estado"))
    if not estado:
        return jsonify({"error": "Estado invalido"}), 400
    estado_anterior = normalizar_estado_cotizacion(cotizacion.estado) or "En progreso"
    if estado == estado_anterior:
        return jsonify({"id": cotizacion.id, "estado": cotizacion.estado}), 200

    cotizacion.estado = estado
    db.session.commit()
    registrar_auditoria(
        "Cambió estado de cotización",
        "Cotización",
        entidad_id=cotizacion.id,
        entidad_ref=cotizacion.numero_cotizacion or str(cotizacion.id),
        detalle=f"Estado anterior: {estado_anterior}. Estado nuevo: {estado}.",
    )
    return jsonify({"id": cotizacion.id, "estado": cotizacion.estado}), 200


@app.route("/", methods=["GET", "POST"])
@token_required
def index():
    if request.method == "POST":
        _, redirect_response = persistir_cotizacion_desde_form()
        if redirect_response:
            return redirect_response
        return redirect(url_for("historial_page"))

    return renderizar_cotizador()


@app.route("/cotizacion/<int:id>/editar", methods=["GET", "POST"])
@token_required
def editar_cotizacion(id):
    cotizacion = Cotizacion.query.get_or_404(id)
    if request.method == "POST":
        _, redirect_response = persistir_cotizacion_desde_form(cotizacion)
        if redirect_response:
            return redirect_response
        return redirect(url_for("historial_page"))

    return renderizar_cotizador(cotizacion)


@app.route("/historial")
@token_required
def historial_page():
    clientes = Cliente.query.order_by(Cliente.nombre).all()
    return render_template("historial.html", clientes=clientes)


@app.route("/dashboard")
@token_required
def dashboard_page():
    filtros = resolver_filtros_dashboard_request()
    estado = filtros["estado"]
    familia = filtros["familia"]
    sector = filtros["sector"]
    subsector = filtros["subsector"]
    moneda = filtros["moneda"]
    periodo = filtros["periodo"]
    desde = filtros["desde"]
    hasta = filtros["hasta"]
    desde_date = filtros["desde_date"]
    hasta_date = filtros["hasta_date"]

    query_periodo = construir_query_dashboard_periodo(filtros)
    cotizaciones_periodo = query_periodo.order_by(Cotizacion.fecha.desc(), Cotizacion.id.desc()).all()
    contexto_operativo = construir_contexto_dashboard_operativo(query_periodo, filtros)
    cotizaciones = contexto_operativo["cotizaciones"]

    total = len(cotizaciones_periodo)
    aceptadas = sum(1 for cot in cotizaciones_periodo if normalizar_estado_cotizacion(cot.estado) == "Aceptada")
    rechazadas = sum(1 for cot in cotizaciones_periodo if normalizar_estado_cotizacion(cot.estado) == "Rechazada")
    en_progreso = sum(1 for cot in cotizaciones_periodo if normalizar_estado_cotizacion(cot.estado) == "En progreso")
    cerradas = aceptadas + rechazadas
    total_importe = round(sum((cot.total_final or 0.0) for cot in cotizaciones_periodo), 2)
    importe_pipeline = round(
        sum((cot.total_final or 0.0) for cot in cotizaciones_periodo if normalizar_estado_cotizacion(cot.estado) == "En progreso"),
        2,
    )
    importe_aceptado = round(
        sum((cot.total_final or 0.0) for cot in cotizaciones_periodo if normalizar_estado_cotizacion(cot.estado) == "Aceptada"),
        2,
    )
    ticket_promedio = round(total_importe / total, 2) if total else 0.0
    tasa_cierre = round((cerradas / total) * 100.0, 1) if total else 0.0
    tasa_aceptacion = round((aceptadas / total) * 100.0, 1) if total else 0.0

    dias_periodo = max((hasta_date - desde_date).days + 1, 1)
    desde_anterior = desde_date - timedelta(days=dias_periodo)
    hasta_anterior = desde_date - timedelta(days=1)
    query_previo = Cotizacion.query.outerjoin(Cliente, Cotizacion.cliente_id == Cliente.id)
    query_previo = aplicar_filtro_fecha_cotizaciones(query_previo, desde_anterior.isoformat(), hasta_anterior.isoformat())
    query_previo = aplicar_filtros_segmentacion_dashboard(
        query_previo, familia=familia, sector=sector, subsector=subsector, moneda=moneda
    )
    cotizaciones_previas = query_previo.all()
    total_previo = len(cotizaciones_previas)
    aceptadas_previas = sum(1 for cot in cotizaciones_previas if normalizar_estado_cotizacion(cot.estado) == "Aceptada")

    resumen = {
        "total": total,
        "cerradas": cerradas,
        "aceptadas": aceptadas,
        "rechazadas": rechazadas,
        "en_progreso": en_progreso,
        "total_importe": total_importe,
        "importe_pipeline": importe_pipeline,
        "importe_aceptado": importe_aceptado,
        "ticket_promedio": ticket_promedio,
        "tasa_cierre": tasa_cierre,
        "tasa_aceptacion": tasa_aceptacion,
        "promedio_diario": round(total / dias_periodo, 1),
        "delta_total": calcular_variacion_porcentual(total, total_previo),
        "delta_aceptadas": calcular_variacion_porcentual(aceptadas, aceptadas_previas),
    }

    series = construir_series_dashboard(cotizaciones_periodo, desde_date, hasta_date)
    monedas = {
        "ARS": {
            "cantidad": sum(1 for cot in cotizaciones_periodo if (cot.moneda or "ARS").upper() == "ARS"),
            "total": round(sum((cot.total_final or 0.0) for cot in cotizaciones_periodo if (cot.moneda or "ARS").upper() == "ARS"), 2),
        },
        "USD": {
            "cantidad": sum(1 for cot in cotizaciones_periodo if (cot.moneda or "ARS").upper() == "USD"),
            "total": round(sum((cot.total_final or 0.0) for cot in cotizaciones_periodo if (cot.moneda or "ARS").upper() == "USD"), 2),
        },
    }
    pipeline_por_moneda = {
        "ARS": {
            "cantidad": sum(
                1
                for cot in cotizaciones_periodo
                if normalizar_estado_cotizacion(cot.estado) == "En progreso" and (cot.moneda or "ARS").upper() == "ARS"
            ),
            "total": round(
                sum(
                    (cot.total_final or 0.0)
                    for cot in cotizaciones_periodo
                    if normalizar_estado_cotizacion(cot.estado) == "En progreso" and (cot.moneda or "ARS").upper() == "ARS"
                ),
                2,
            ),
        },
        "USD": {
            "cantidad": sum(
                1
                for cot in cotizaciones_periodo
                if normalizar_estado_cotizacion(cot.estado) == "En progreso" and (cot.moneda or "ARS").upper() == "USD"
            ),
            "total": round(
                sum(
                    (cot.total_final or 0.0)
                    for cot in cotizaciones_periodo
                    if normalizar_estado_cotizacion(cot.estado) == "En progreso" and (cot.moneda or "ARS").upper() == "USD"
                ),
                2,
            ),
        },
    }
    aceptado_por_moneda = {
        "ARS": {
            "cantidad": sum(
                1
                for cot in cotizaciones_periodo
                if normalizar_estado_cotizacion(cot.estado) == "Aceptada" and (cot.moneda or "ARS").upper() == "ARS"
            ),
            "total": round(
                sum(
                    (cot.total_final or 0.0)
                    for cot in cotizaciones_periodo
                    if normalizar_estado_cotizacion(cot.estado) == "Aceptada" and (cot.moneda or "ARS").upper() == "ARS"
                ),
                2,
            ),
        },
        "USD": {
            "cantidad": sum(
                1
                for cot in cotizaciones_periodo
                if normalizar_estado_cotizacion(cot.estado) == "Aceptada" and (cot.moneda or "ARS").upper() == "USD"
            ),
            "total": round(
                sum(
                    (cot.total_final or 0.0)
                    for cot in cotizaciones_periodo
                    if normalizar_estado_cotizacion(cot.estado) == "Aceptada" and (cot.moneda or "ARS").upper() == "USD"
                ),
                2,
            ),
        },
    }
    top_clientes = construir_top_clientes_dashboard(cotizaciones_periodo)
    familias_breakdown = construir_desglose_dashboard(
        cotizaciones_periodo, lambda cot: cot.familia, default_label="Sin familia", limite=6
    )
    sectores_breakdown = construir_desglose_dashboard(
        cotizaciones_periodo,
        lambda cot: cot.cliente_ref.sector if cot.cliente_ref else "",
        default_label="Sin sector",
    )
    subsectores_breakdown = construir_desglose_dashboard(
        cotizaciones_periodo,
        lambda cot: cot.cliente_ref.subsector if cot.cliente_ref else "",
        default_label="Sin subsector",
        limite=8,
    )
    estado_foco_label = "Vista general del periodo"
    filtros_activos = []
    if familia:
        filtros_activos.append({"icon": "bi-diagram-3", "label": f"Familia: {familia}"})
    if sector:
        filtros_activos.append({"icon": "bi-building", "label": f"Sector: {sector}"})
    if subsector:
        filtros_activos.append({"icon": "bi-tags", "label": f"Subsector: {subsector}"})
    if moneda:
        filtros_activos.append({"icon": "bi-cash-stack", "label": f"Moneda: {moneda}"})
    if filtros["cliente"]:
        filtros_activos.append({"icon": "bi-person-vcard", "label": f"Cliente: {filtros['cliente']}"})

    top_familia = familias_breakdown[0]["nombre"] if familias_breakdown else "Sin datos"
    top_sector = sectores_breakdown[0]["nombre"] if sectores_breakdown else "Sin datos"
    top_subsector = subsectores_breakdown[0]["nombre"] if subsectores_breakdown else "Sin datos"

    return render_template(
        "dashboard.html",
        resumen=resumen,
        series=series,
        monedas=monedas,
        pipeline_por_moneda=pipeline_por_moneda,
        aceptado_por_moneda=aceptado_por_moneda,
        top_clientes=top_clientes,
        familias_breakdown=familias_breakdown,
        sectores_breakdown=sectores_breakdown,
        subsectores_breakdown=subsectores_breakdown,
        estado_foco_label=estado_foco_label,
        filtro_desde_legible=desde_date.strftime("%d/%m/%Y"),
        filtro_hasta_legible=hasta_date.strftime("%d/%m/%Y"),
        dias_periodo=dias_periodo,
        filtros_activos=filtros_activos,
        familias_disponibles=FAMILIAS_COTIZACION,
        sectores_cliente=SECTORES_CLIENTE,
        top_familia=top_familia,
        top_sector=top_sector,
        top_subsector=top_subsector,
        **contexto_operativo,
    )


@app.route("/dashboard/detalle-operativo")
@token_required
def dashboard_detalle_operativo():
    filtros = resolver_filtros_dashboard_request()
    query_periodo = construir_query_dashboard_periodo(filtros)
    contexto_operativo = construir_contexto_dashboard_operativo(query_periodo, filtros)
    return render_template("_dashboard_operativo.html", **contexto_operativo)


@app.route("/auditoria")
@token_required
def auditoria_page():
    usuario_filtro = (request.args.get("usuario") or "").strip()
    page = parsear_entero_positivo(request.args.get("page"), default=1) or 1
    query = Auditoria.query
    if usuario_filtro:
        query = query.filter(Auditoria.username == usuario_filtro)
    paginacion = paginar_query(
        query.order_by(Auditoria.fecha.desc(), Auditoria.id.desc()),
        page=page,
        per_page=AUDITORIA_PER_PAGE,
    )
    usuarios_auditoria = [row[0] for row in db.session.query(Auditoria.username).distinct().order_by(Auditoria.username.asc()).all()]
    return render_template(
        "auditoria.html",
        registros=paginacion["items"],
        usuario_filtro=usuario_filtro,
        usuarios_auditoria=usuarios_auditoria,
        paginacion={
            **paginacion,
            "pages_visible": construir_paginas_visibles(paginacion["page"], paginacion["pages"]),
        },
    )


@app.route("/filtrar_historial")
@token_required
def filtrar_historial():
    cliente_id_raw = (request.args.get("cliente_id") or "").strip()
    cliente = (request.args.get("cliente") or "").strip()
    desde = (request.args.get("desde") or "").strip()
    hasta = (request.args.get("hasta") or "").strip()
    moneda = (request.args.get("moneda") or "").strip().upper()
    page = parsear_entero_positivo(request.args.get("page"), default=1) or 1

    query = Cotizacion.query.outerjoin(Cliente, Cotizacion.cliente_id == Cliente.id)
    if cliente_id_raw.isdigit():
        cliente_sel = db.session.get(Cliente, int(cliente_id_raw))
        if cliente_sel:
            condiciones_cliente = [Cotizacion.cliente_id == cliente_sel.id]
            if cliente_sel.nombre:
                condiciones_cliente.append(Cotizacion.cliente.ilike(f"%{cliente_sel.nombre}%"))
                condiciones_cliente.append(Cliente.nombre.ilike(f"%{cliente_sel.nombre}%"))
            if cliente_sel.razon_social:
                condiciones_cliente.append(Cotizacion.cliente_razon_social.ilike(f"%{cliente_sel.razon_social}%"))
                condiciones_cliente.append(Cliente.razon_social.ilike(f"%{cliente_sel.razon_social}%"))
            if cliente_sel.cuit:
                condiciones_cliente.append(Cotizacion.cliente_cuit == cliente_sel.cuit)
                condiciones_cliente.append(Cliente.cuit == cliente_sel.cuit)
            query = query.filter(or_(*condiciones_cliente))
        else:
            query = query.filter(Cotizacion.id == -1)
    elif cliente:
        patron = f"%{cliente}%"
        query = query.filter(
            or_(
                Cotizacion.cliente.ilike(patron),
                Cotizacion.cliente_razon_social.ilike(patron),
                Cliente.nombre.ilike(patron),
                Cliente.razon_social.ilike(patron),
                Cliente.cuit.ilike(patron),
            )
        )
    query = aplicar_filtro_fecha_cotizaciones(query, desde, hasta)
    if moneda in ("ARS", "USD"):
        query = query.filter(Cotizacion.moneda == moneda)

    paginacion = paginar_query(query.order_by(Cotizacion.id.desc()), page=page, per_page=HISTORIAL_PER_PAGE)
    cotizaciones = paginacion["items"]
    resultados = [
        {
            "id": c.id,
            "numero_cotizacion": c.numero_cotizacion or str(c.id).zfill(4),
            "estado": normalizar_estado_cotizacion(c.estado) or "En progreso",
            "cliente": c.cliente,
            "cliente_nombre": (c.cliente_ref.nombre if c.cliente_ref else c.cliente) or "Sin Cliente",
            "familia": c.familia or "",
            "fecha": c.fecha.strftime("%d/%m/%Y"),
            "moneda": c.moneda or "ARS",
            "total_final": c.total_final or 0.0,
            "total": f"${c.total_final:,.2f}",
        }
        for c in cotizaciones
    ]
    return jsonify(
        {
            "items": resultados,
            "pagination": {
                **{k: paginacion[k] for k in ("page", "per_page", "total", "pages", "has_prev", "has_next", "prev_page", "next_page", "start", "end")},
                "pages_visible": construir_paginas_visibles(paginacion["page"], paginacion["pages"]),
            },
        }
    )


@app.route("/cotizacion/<int:id>")
@token_required
def ver_cotizacion(id):
    cot = Cotizacion.query.get_or_404(id)
    return render_template("cotizacion_cliente.html", cot=cot, domicilio_empresa=DOMICILIO)


@app.route("/cotizacion/<int:id>/xlsx")
@token_required
def exportar_cotizacion_xlsx(id):
    cotizacion = Cotizacion.query.get_or_404(id)
    contenido = generar_excel_cotizacion(cotizacion)
    nombre_base = (cotizacion.numero_cotizacion or f"cotizacion-{cotizacion.id}").replace("/", "-").replace("\\", "-")
    return Response(
        contenido,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{nombre_base}.xlsx"'},
    )


@app.route("/cotizacion/<int:id>/eliminar", methods=["POST"])
@token_required
@admin_required
def eliminar_cotizacion(id):
    cotizacion = Cotizacion.query.get_or_404(id)
    numero_ref = cotizacion.numero_cotizacion or str(cotizacion.id)
    cliente_ref = cotizacion.cliente or "Sin cliente"
    familia_ref = cotizacion.familia or "Sin familia"
    moneda_ref = cotizacion.moneda or "ARS"
    total_ref = cotizacion.total_final or 0.0
    estado_ref = cotizacion.estado or "En progreso"

    for item in cotizacion.items:
        eliminar_imagen_local(item.imagen_url)

    db.session.delete(cotizacion)
    db.session.commit()

    registrar_auditoria(
        "Eliminó cotización",
        "Cotización",
        entidad_id=id,
        entidad_ref=numero_ref,
        detalle=f"Cliente: {cliente_ref}. Familia: {familia_ref}. Estado: {estado_ref}. Total: {moneda_ref} {total_ref:,.2f}.",
    )

    return jsonify({"ok": True, "id": id, "numero_cotizacion": numero_ref}), 200


if __name__ == "__main__":
    app.run(debug=True)
