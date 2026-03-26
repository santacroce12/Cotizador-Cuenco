from __future__ import annotations

import random
import sys
from datetime import datetime, timedelta
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from app import (
    CUIT,
    DOMICILIO,
    FAMILIAS_COTIZACION,
    NOMBRE_FANTASIA,
    RAZON_SOCIAL,
    Cliente,
    Cotizacion,
    ItemCotizacion,
    app,
    db,
    formatear_numero_cotizacion,
)


RANDOM_SEED = 20260326
PUBLIC_BATCH_SIZE = 12
ITEMS_CATALOGO = {
    "SEGURIDAD URBANA": [
        ("Camara IP urbana 4MP", 185.0, 21.0),
        ("Switch PoE gestionado", 220.0, 21.0),
        ("Servidor VMS compacto", 940.0, 21.0),
    ],
    "PARKING": [
        ("Barrera vehicular", 760.0, 21.0),
        ("Controlador de acceso", 330.0, 21.0),
        ("Lector UHF", 215.0, 21.0),
    ],
    "TRANSPORTE INTELIGENTE": [
        ("Panel de mensajeria variable", 1250.0, 21.0),
        ("Controlador ITS", 910.0, 21.0),
        ("GPS de flota", 130.0, 21.0),
    ],
    "CONECTIVIDAD SATELITAL": [
        ("Terminal satelital vehicular", 1650.0, 21.0),
        ("Router satelital", 420.0, 21.0),
        ("Unidad de energia outdoor", 240.0, 10.5),
    ],
    "SALAS DE CONTROL": [
        ("Videowall profesional", 1180.0, 21.0),
        ("Consola operativa", 860.0, 21.0),
        ("Workstation grafica", 980.0, 21.0),
    ],
    "SMART CITIES": [
        ("Gateway IoT urbano", 290.0, 21.0),
        ("Sensor ambiental", 125.0, 21.0),
        ("Totem ciudadano interactivo", 530.0, 21.0),
    ],
}

PLAN_PUBLICO = [
    ("Demo Municipalidad de Godoy Cruz", "Municipal", "SEGURIDAD URBANA", "En progreso", "ARS"),
    ("Demo Municipalidad de Guaymallen", "Municipal", "SMART CITIES", "Aceptada", "ARS"),
    ("Demo EPRE Mendoza", "Provincial", "SALAS DE CONTROL", "En progreso", "USD"),
    ("Demo Ministerio de Seguridad Mendoza", "Provincial", "SEGURIDAD URBANA", "Rechazada", "ARS"),
    ("Demo Vialidad Nacional Cuyo", "Nacional", "TRANSPORTE INTELIGENTE", "Aceptada", "USD"),
    ("Demo ANSV Cuyo", "Nacional", "PARKING", "En progreso", "ARS"),
    ("Demo UNCuyo", "Otro", "SMART CITIES", "Aceptada", "ARS"),
    ("Demo Hospital Central", "Otro", "SALAS DE CONTROL", "En progreso", "ARS"),
    ("Demo Municipalidad de Godoy Cruz", "Municipal", "PARKING", "Aceptada", "ARS"),
    ("Demo EPRE Mendoza", "Provincial", "CONECTIVIDAD SATELITAL", "En progreso", "USD"),
    ("Demo Vialidad Nacional Cuyo", "Nacional", "SEGURIDAD URBANA", "Rechazada", "ARS"),
    ("Demo Hospital Central", "Otro", "SMART CITIES", "Aceptada", "ARS"),
]


def obtener_clientes_publicos() -> dict[str, Cliente]:
    clientes = Cliente.query.filter_by(sector="Publico").all()
    return {cliente.nombre: cliente for cliente in clientes}


def siguiente_numero_cotizacion(fecha: datetime) -> str:
    prefijo = f"CT-{fecha.year:04d}-"
    ultima = (
        Cotizacion.query.filter(Cotizacion.numero_cotizacion.like(f"{prefijo}%"))
        .order_by(Cotizacion.numero_cotizacion.desc())
        .first()
    )
    if not ultima or not ultima.numero_cotizacion:
        return formatear_numero_cotizacion(fecha.year, 1)
    try:
        secuencia = int(ultima.numero_cotizacion.rsplit("-", 1)[1])
    except Exception:
        secuencia = Cotizacion.query.filter(Cotizacion.fecha >= datetime(fecha.year, 1, 1)).count()
    return formatear_numero_cotizacion(fecha.year, secuencia + 1)


def crear_items(familia: str, condicion_iva: str, rng: random.Random) -> tuple[list[ItemCotizacion], float, float, float]:
    neto_total = 0.0
    iva_total = 0.0
    items: list[ItemCotizacion] = []
    catalogo = ITEMS_CATALOGO[familia]
    cantidad_items = rng.randint(2, 3)
    seleccionados = rng.sample(catalogo, k=cantidad_items)

    for descripcion, costo_base, iva_default in seleccionados:
        cantidad = rng.randint(1, 6)
        costo_unitario = round(costo_base * rng.uniform(0.92, 1.18), 2)
        costo_extra = round(rng.uniform(4.0, 8.0), 1)
        margen = round(rng.uniform(0.45, 0.95), 2)
        costo_con_extra = costo_unitario * (1 + costo_extra / 100)
        precio_venta = round(costo_con_extra * (1 + margen), 2)
        subtotal_neto = round(precio_venta * cantidad, 2)
        iva_pct = iva_default if condicion_iva != "Exento" else 0.0
        iva_item = round(subtotal_neto * (iva_pct / 100), 2)
        subtotal = round(subtotal_neto + iva_item, 2)

        items.append(
            ItemCotizacion(
                descripcion=descripcion,
                cantidad=float(cantidad),
                costo_unitario=costo_unitario,
                costo_extra=costo_extra,
                margen=margen,
                iva_item=iva_pct,
                precio_venta=precio_venta,
                subtotal=subtotal,
                imagen_url=None,
            )
        )
        neto_total += subtotal_neto
        iva_total += iva_item

    return items, round(neto_total, 2), round(iva_total, 2), round(neto_total + iva_total, 2)


def main() -> None:
    rng = random.Random(RANDOM_SEED)
    with app.app_context():
        clientes = obtener_clientes_publicos()
        faltantes = [nombre for nombre, _, _, _, _ in PLAN_PUBLICO if nombre not in clientes]
        if faltantes:
            raise RuntimeError(f"Faltan clientes publicos requeridos para el seed: {faltantes}")

        ahora = datetime.utcnow()
        creadas = []
        for idx, (cliente_nombre, _subsector, familia, estado, moneda) in enumerate(PLAN_PUBLICO[:PUBLIC_BATCH_SIZE]):
            cliente = clientes[cliente_nombre]
            fecha = ahora - timedelta(days=idx, hours=(idx * 3) % 11, minutes=idx * 7)
            items, total_neto, total_iva, total_final = crear_items(familia, cliente.condicion_iva or "Exento", rng)

            cotizacion = Cotizacion(
                numero_cotizacion=siguiente_numero_cotizacion(fecha),
                estado=estado,
                nombre_fantasia=NOMBRE_FANTASIA,
                razon_social=RAZON_SOCIAL,
                cuit=CUIT,
                cliente=cliente.nombre,
                cliente_id=cliente.id,
                cliente_razon_social=cliente.razon_social,
                cliente_cuit=cliente.cuit,
                familia=familia,
                fecha=fecha,
                moneda=moneda,
                condicion_iva=cliente.condicion_iva,
                total_neto=total_neto,
                total_iva=total_iva,
                total_final=total_final,
            )
            for item in items:
                cotizacion.items.append(item)
            db.session.add(cotizacion)
            creadas.append((cotizacion.numero_cotizacion, cliente.nombre, cliente.subsector, familia, estado, moneda))

        db.session.commit()

        print(f"cotizaciones_publicas_agregadas: {len(creadas)}")
        for row in creadas:
            print(row)


if __name__ == "__main__":
    main()
