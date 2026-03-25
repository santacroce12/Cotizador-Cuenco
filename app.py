import os
import json
import re
import smtplib
import threading
import time
from datetime import datetime, timedelta
from email.message import EmailMessage
from io import BytesIO
from pathlib import Path

from flask import Flask, Response, jsonify, redirect, render_template, request, url_for
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func, or_
from werkzeug.utils import secure_filename


app = Flask(__name__)
db_path = Path(app.root_path) / "database.db"
app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{db_path.as_posix()}"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db = SQLAlchemy(app)

NOMBRE_FANTASIA = "Cuenco Tech"
RAZON_SOCIAL = "Cuenco Tech S.A."
CUIT = "30-71831614-2"
DOMICILIO = "Rafael Cubillos 2056, M5500 Godoy Cruz, Mendoza"
DEFAULT_FOLLOWUP_EMAIL = "jsantacroce@cuencotech.com"
DEFAULT_SMTP_HOST = "a0021139.ferozo.com"
DEFAULT_SMTP_PORT = 465
DEFAULT_SMTP_USERNAME = "cotizador@cuencotech.com"
DEFAULT_SMTP_FROM = "cotizador@cuencotech.com"
DEFAULT_APP_BASE_URL = "https://cuencotech.com"
LOCAL_SETTINGS_PATH = Path(app.root_path) / "local_settings.json"

PLACEHOLDER_PRODUCTO = "placeholder_product.png"
UPLOADS_PRODUCTOS_DIR = Path(app.static_folder) / "uploads" / "productos"
UPLOADS_PRODUCTOS_DIR.mkdir(parents=True, exist_ok=True)
ALLOWED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".gif"}
ESTADOS_COTIZACION = ("En progreso", "Aceptada", "Rechazada")
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
    email = db.Column(db.String(100))
    telefono = db.Column(db.String(50))
    condicion_iva = db.Column(db.String(50), default="Consumidor Final")
    cotizaciones = db.relationship("Cotizacion", backref="cliente_ref", lazy=True)


def formatear_numero_cotizacion(anio, secuencia):
    return f"{NUMERO_COTIZACION_PREFIX}-{int(anio):04d}-{int(secuencia):04d}"


def parsear_numero_cotizacion(valor):
    numero = (valor or "").strip().upper()
    match = re.fullmatch(rf"{NUMERO_COTIZACION_PREFIX}-(\d{{4}})-(\d+)", numero)
    if not match:
        return None
    return int(match.group(1)), int(match.group(2))


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
    db.session.execute(
        db.text("UPDATE cotizacion SET estado = 'En progreso' WHERE estado IS NULL OR TRIM(estado) = ''")
    )
    db.session.commit()
    columnas_cliente = [col[1] for col in db.session.execute(db.text("PRAGMA table_info(cliente)")).fetchall()]
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


def archivo_imagen_permitido(filename):
    return Path(filename or "").suffix.lower() in ALLOWED_IMAGE_EXTENSIONS


def guardar_imagen_producto(file_storage, descripcion, row_id):
    if not file_storage or not file_storage.filename:
        return None
    if not archivo_imagen_permitido(file_storage.filename):
        return None

    base = secure_filename(descripcion) or "producto"
    ext = Path(file_storage.filename).suffix.lower()
    nombre_archivo = f"{datetime.utcnow():%Y%m%d%H%M%S%f}_{secure_filename(str(row_id))}_{base}{ext}"
    destino = UPLOADS_PRODUCTOS_DIR / nombre_archivo
    file_storage.save(destino)
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


def normalizar_estado_cotizacion(valor):
    valor = (valor or "").strip()
    if not valor:
        return "En progreso"
    for estado in ESTADOS_COTIZACION:
        if valor.lower() == estado.lower():
            return estado
    return None


def parsear_entero_positivo(valor, default=None):
    valor = (valor or "").strip()
    if not valor:
        return default
    try:
        numero = int(valor)
    except ValueError:
        return default
    return numero if numero > 0 else default


def normalizar_followup_email(valor, cliente_sel=None):
    email = (valor or "").strip()
    if email:
        return email
    config = obtener_config_smtp()
    if config["default_to"]:
        return config["default_to"]
    if cliente_sel and cliente_sel.email:
        return cliente_sel.email.strip()
    return ""


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
        cotizacion.seguimiento_email = email_raw or None
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


def resolver_rango_dashboard(periodo, desde_raw="", hasta_raw=""):
    hoy = datetime.now().date()
    periodo = (periodo or "30").strip().lower()
    desde = (desde_raw or "").strip()
    hasta = (hasta_raw or "").strip()

    if periodo == "30":
        hasta = hoy.isoformat()
        desde = (hoy - timedelta(days=29)).isoformat()
    elif periodo == "60":
        hasta = hoy.isoformat()
        desde = (hoy - timedelta(days=59)).isoformat()
    else:
        periodo = "custom"
        if not hasta:
            hasta = hoy.isoformat()

    if desde and hasta and desde > hasta:
        desde, hasta = hasta, desde

    return periodo, desde, hasta


def aplicar_filtro_estado_dashboard(query, estado):
    estado = (estado or "todos").strip().lower()
    if estado == "cerradas":
        return query.filter(Cotizacion.estado.in_(("Aceptada", "Rechazada")))
    if estado == "aceptadas":
        return query.filter(Cotizacion.estado == "Aceptada")
    if estado == "rechazadas":
        return query.filter(Cotizacion.estado == "Rechazada")
    if estado == "en_progreso":
        return query.filter(Cotizacion.estado == "En progreso")
    return query


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
        imagen_local = guardar_imagen_producto(request.files.get(f"foto_{row_id}"), desc, row_id)

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

    if not cotizacion.cliente or not cotizacion.items:
        destino = "editar_cotizacion" if es_edicion else "index"
        kwargs = {"id": cotizacion.id} if es_edicion else {}
        return None, redirect(url_for(destino, **kwargs))

    cotizacion.total_neto = round(neto_total, 2)
    cotizacion.total_iva = round(iva_total, 2)
    cotizacion.total_final = round(sum(item.subtotal for item in cotizacion.items), 2)

    db.session.add(cotizacion)
    db.session.commit()
    return cotizacion, None


@app.route("/api/clientes", methods=["POST"])
def agregar_cliente():
    data = request.get_json(silent=True) or {}
    cuit = (data.get("cuit") or "").strip()
    if cuit:
        repetido = Cliente.query.filter(Cliente.cuit == cuit).first()
        if repetido:
            return jsonify({"error": "Ya existe un cliente con ese CUIT"}), 400
    nuevo = Cliente(
        nombre=(data.get("nombre") or "").strip(),
        razon_social=(data.get("razon_social") or "").strip(),
        cuit=cuit,
        email=(data.get("email") or "").strip(),
        telefono=(data.get("telefono") or "").strip(),
        condicion_iva=(data.get("condicion_iva") or "Consumidor Final").strip() or "Consumidor Final",
    )
    if nuevo.nombre:
        db.session.add(nuevo)
        db.session.commit()
        return jsonify(
            {
                "id": nuevo.id,
                "nombre": nuevo.nombre,
                "razon_social": nuevo.razon_social,
                "cuit": nuevo.cuit,
                "email": nuevo.email,
                "telefono": nuevo.telefono,
                "condicion_iva": nuevo.condicion_iva,
            }
        ), 201
    return jsonify({"error": "Nombre requerido"}), 400


@app.before_request
def ensure_followup_worker():
    if app.config.get("TESTING"):
        return
    if app.debug and os.environ.get("WERKZEUG_RUN_MAIN") not in (None, "true"):
        return
    iniciar_worker_recordatorios()


@app.route("/cotizacion/<int:id>/estado", methods=["POST"])
def actualizar_estado_cotizacion(id):
    cotizacion = Cotizacion.query.get_or_404(id)
    data = request.get_json(silent=True) or {}
    estado = normalizar_estado_cotizacion(data.get("estado"))
    if not estado:
        return jsonify({"error": "Estado invalido"}), 400

    cotizacion.estado = estado
    db.session.commit()
    return jsonify({"id": cotizacion.id, "estado": cotizacion.estado}), 200


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        _, redirect_response = persistir_cotizacion_desde_form()
        if redirect_response:
            return redirect_response
        return redirect(url_for("historial_page"))

    return renderizar_cotizador()


@app.route("/cotizacion/<int:id>/editar", methods=["GET", "POST"])
def editar_cotizacion(id):
    cotizacion = Cotizacion.query.get_or_404(id)
    if request.method == "POST":
        _, redirect_response = persistir_cotizacion_desde_form(cotizacion)
        if redirect_response:
            return redirect_response
        return redirect(url_for("historial_page"))

    return renderizar_cotizador(cotizacion)


@app.route("/historial")
def historial_page():
    clientes = Cliente.query.order_by(Cliente.nombre).all()
    return render_template("historial.html", clientes=clientes)


@app.route("/dashboard")
def dashboard_page():
    estado = (request.args.get("estado") or "todos").strip().lower()
    if estado not in ("todos", "cerradas", "en_progreso", "aceptadas", "rechazadas"):
        estado = "todos"

    periodo, desde, hasta = resolver_rango_dashboard(
        request.args.get("periodo"),
        request.args.get("desde"),
        request.args.get("hasta"),
    )

    query_periodo = aplicar_filtro_fecha_cotizaciones(Cotizacion.query, desde, hasta)
    query_tabla = aplicar_filtro_estado_dashboard(query_periodo, estado)
    cotizaciones = query_tabla.order_by(Cotizacion.fecha.desc(), Cotizacion.id.desc()).all()

    hoy = datetime.now().date()
    ultimos_30_desde = (hoy - timedelta(days=29)).isoformat()
    ultimos_60_desde = (hoy - timedelta(days=59)).isoformat()
    cerradas_30 = aplicar_filtro_fecha_cotizaciones(Cotizacion.query, ultimos_30_desde, hoy.isoformat()).filter(
        Cotizacion.estado.in_(("Aceptada", "Rechazada"))
    ).count()
    cerradas_60 = aplicar_filtro_fecha_cotizaciones(Cotizacion.query, ultimos_60_desde, hoy.isoformat()).filter(
        Cotizacion.estado.in_(("Aceptada", "Rechazada"))
    ).count()

    resumen = {
        "total": query_periodo.count(),
        "cerradas": query_periodo.filter(Cotizacion.estado.in_(("Aceptada", "Rechazada"))).count(),
        "aceptadas": query_periodo.filter(Cotizacion.estado == "Aceptada").count(),
        "rechazadas": query_periodo.filter(Cotizacion.estado == "Rechazada").count(),
        "en_progreso": query_periodo.filter(Cotizacion.estado == "En progreso").count(),
        "cerradas_30": cerradas_30,
        "cerradas_60": cerradas_60,
    }

    return render_template(
        "dashboard.html",
        cotizaciones=cotizaciones,
        resumen=resumen,
        filtro_estado=estado,
        filtro_periodo=periodo,
        filtro_desde=desde,
        filtro_hasta=hasta,
    )


@app.route("/filtrar_historial")
def filtrar_historial():
    cliente_id_raw = (request.args.get("cliente_id") or "").strip()
    cliente = (request.args.get("cliente") or "").strip()
    desde = (request.args.get("desde") or "").strip()
    hasta = (request.args.get("hasta") or "").strip()
    moneda = (request.args.get("moneda") or "").strip().upper()

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

    cotizaciones = query.order_by(Cotizacion.id.desc()).all()
    resultados = [
        {
            "id": c.id,
            "numero_cotizacion": c.numero_cotizacion or str(c.id).zfill(4),
            "estado": normalizar_estado_cotizacion(c.estado) or "En progreso",
            "cliente": c.cliente,
            "cliente_nombre": (c.cliente_ref.nombre if c.cliente_ref else c.cliente) or "Sin Cliente",
            "fecha": c.fecha.strftime("%d/%m/%Y"),
            "moneda": c.moneda or "ARS",
            "total_final": c.total_final or 0.0,
            "total": f"${c.total_final:,.2f}",
        }
        for c in cotizaciones
    ]
    return jsonify(resultados)


@app.route("/cotizacion/<int:id>")
def ver_cotizacion(id):
    cot = Cotizacion.query.get_or_404(id)
    return render_template("cotizacion_cliente.html", cot=cot, domicilio_empresa=DOMICILIO)


@app.route("/cotizacion/<int:id>/xlsx")
def exportar_cotizacion_xlsx(id):
    cotizacion = Cotizacion.query.get_or_404(id)
    contenido = generar_excel_cotizacion(cotizacion)
    nombre_base = (cotizacion.numero_cotizacion or f"cotizacion-{cotizacion.id}").replace("/", "-").replace("\\", "-")
    return Response(
        contenido,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{nombre_base}.xlsx"'},
    )


if __name__ == "__main__":
    app.run(debug=True)
