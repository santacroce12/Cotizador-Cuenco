from __future__ import annotations

import random
import sys
from collections import Counter, defaultdict
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
    SECTORES_CLIENTE,
    Auditoria,
    Cliente,
    Cotizacion,
    ItemCotizacion,
    app,
    db,
    formatear_numero_cotizacion,
    normalizar_estado_cotizacion,
)


RANDOM_SEED = 20260325
DEMO_NAME_PREFIX = "Demo "
DEMO_CUIT_START = 30_97000000
START_DATE = datetime(2026, 1, 2, 9, 0, 0)

PUBLIC_SUBS = list(SECTORES_CLIENTE["Publico"])
PRIVATE_SUBS = list(SECTORES_CLIENTE["Privado"])
FAMILIES = list(FAMILIAS_COTIZACION)

CLIENT_SPECS = {
    ("Publico", PUBLIC_SUBS[0]): [
        ("Municipalidad de Godoy Cruz", "Municipalidad de Godoy Cruz"),
        ("Municipalidad de Guaymallen", "Municipalidad de Guaymallen"),
    ],
    ("Publico", PUBLIC_SUBS[1]): [
        ("Ministerio de Seguridad Mendoza", "Ministerio de Seguridad de Mendoza"),
        ("EPRE Mendoza", "Ente Provincial Regulador Electrico"),
    ],
    ("Publico", PUBLIC_SUBS[2]): [
        ("Vialidad Nacional Cuyo", "Direccion Nacional de Vialidad - Distrito Cuyo"),
        ("ANSV Cuyo", "Agencia Nacional de Seguridad Vial - Region Cuyo"),
    ],
    ("Publico", PUBLIC_SUBS[3]): [
        ("UNCuyo", "Universidad Nacional de Cuyo"),
        ("Hospital Central", "Hospital Central de Mendoza"),
    ],
    ("Privado", PRIVATE_SUBS[0]): [
        ("Energia Andina", "Energia Andina S.A."),
        ("Solar del Oeste", "Solar del Oeste S.A."),
    ],
    ("Privado", PRIVATE_SUBS[1]): [
        ("AgroPack Cuyo", "AgroPack Cuyo S.A."),
        ("Bodega Valle Norte", "Bodega Valle Norte S.A."),
    ],
    ("Privado", PRIVATE_SUBS[2]): [
        ("Hotel Cordillera", "Hotel Cordillera S.A."),
        ("Hotel Plaza Andina", "Hotel Plaza Andina S.R.L."),
    ],
    ("Privado", PRIVATE_SUBS[3]): [
        ("Instituto Tecnico Sur", "Instituto Tecnico Sur S.A."),
        ("Colegio Nuevo Cuyo", "Colegio Nuevo Cuyo S.R.L."),
    ],
    ("Privado", PRIVATE_SUBS[4]): [
        ("Turismo Alta Montana", "Turismo Alta Montana S.A."),
        ("Andes Experience", "Andes Experience S.R.L."),
    ],
    ("Privado", PRIVATE_SUBS[5]): [
        ("Retail Centro", "Retail Centro S.A."),
        ("Supermercados del Oeste", "Supermercados del Oeste S.A."),
    ],
    ("Privado", PRIVATE_SUBS[6]): [
        ("Logistica Cuyo", "Logistica Cuyo S.A."),
        ("Parque Industrial Oeste", "Parque Industrial Oeste S.A."),
    ],
}

FAMILY_BY_SEGMENT = {
    ("Publico", PUBLIC_SUBS[0]): [FAMILIES[0], FAMILIES[5], FAMILIES[1]],
    ("Publico", PUBLIC_SUBS[1]): [FAMILIES[0], FAMILIES[4], FAMILIES[2]],
    ("Publico", PUBLIC_SUBS[2]): [FAMILIES[2], FAMILIES[3], FAMILIES[0]],
    ("Publico", PUBLIC_SUBS[3]): [FAMILIES[4], FAMILIES[5], FAMILIES[0]],
    ("Privado", PRIVATE_SUBS[0]): [FAMILIES[3], FAMILIES[4], FAMILIES[5]],
    ("Privado", PRIVATE_SUBS[1]): [FAMILIES[5], FAMILIES[3], FAMILIES[1]],
    ("Privado", PRIVATE_SUBS[2]): [FAMILIES[1], FAMILIES[5], FAMILIES[3]],
    ("Privado", PRIVATE_SUBS[3]): [FAMILIES[4], FAMILIES[5], FAMILIES[0]],
    ("Privado", PRIVATE_SUBS[4]): [FAMILIES[1], FAMILIES[5], FAMILIES[2]],
    ("Privado", PRIVATE_SUBS[5]): [FAMILIES[1], FAMILIES[0], FAMILIES[5]],
    ("Privado", PRIVATE_SUBS[6]): [FAMILIES[2], FAMILIES[3], FAMILIES[4]],
}

PRODUCT_CATALOG = {
    FAMILIES[0]: [
        ("Camara analitica 4MP", 185.0, 21.0),
        ("Servidor de video gestionado", 940.0, 21.0),
        ("Licencia VMS avanzada", 310.0, 21.0),
        ("Poste inteligente reforzado", 520.0, 10.5),
        ("Cabina antivandalica urbana", 270.0, 21.0),
    ],
    FAMILIES[1]: [
        ("Barrera vehicular industrial", 760.0, 21.0),
        ("Sensor de ocupacion LPR", 140.0, 21.0),
        ("Totem de acceso parking", 460.0, 21.0),
        ("Lector UHF de largo alcance", 215.0, 21.0),
        ("Controlador de carril", 330.0, 21.0),
    ],
    FAMILIES[2]: [
        ("Camara embarcada dual", 280.0, 21.0),
        ("Panel de mensajeria variable", 1250.0, 21.0),
        ("Controlador ITS de interseccion", 910.0, 21.0),
        ("Switch industrial PoE", 190.0, 21.0),
        ("GPS de flota reforzado", 130.0, 21.0),
    ],
    FAMILIES[3]: [
        ("Terminal satelital vehicular", 1650.0, 21.0),
        ("Antena Ka portable", 890.0, 21.0),
        ("Router satelital failover", 420.0, 21.0),
        ("Abono gestion de enlace", 180.0, 21.0),
        ("Unidad de energia outdoor", 240.0, 10.5),
    ],
    FAMILIES[4]: [
        ("Videowall panel profesional", 1180.0, 21.0),
        ("Consola operativa modular", 860.0, 21.0),
        ("Decoder de videowall", 410.0, 21.0),
        ("Workstation grafica", 980.0, 21.0),
        ("Procesador de control KVM", 560.0, 21.0),
    ],
    FAMILIES[5]: [
        ("Gateway IoT urbano", 290.0, 21.0),
        ("Sensor ambiental multiparametro", 125.0, 21.0),
        ("Luminaria inteligente LED", 210.0, 10.5),
        ("Totem ciudadano interactivo", 530.0, 21.0),
        ("Nodo de telemetria urbana", 170.0, 21.0),
    ],
}


def demo_slug(texto: str) -> str:
    return (
        texto.lower()
        .replace(" ", "-")
        .replace(".", "")
        .replace("/", "-")
        .replace(",", "")
    )


def demo_domicilio(idx: int, sector: str, subsector: str) -> str:
    base = 400 + (idx * 17)
    ciudad = "Godoy Cruz" if sector == "Privado" else "Ciudad de Mendoza"
    return f"Demo {subsector} {base}, {ciudad}, Mendoza"


def elegir_condicion_iva(rng: random.Random, sector: str) -> str:
    if sector == "Publico":
        opciones = ["Exento", "Responsable Inscrito", "Consumidor Final"]
        pesos = [0.55, 0.25, 0.20]
    else:
        opciones = ["Responsable Inscrito", "Consumidor Final", "Exento"]
        pesos = [0.65, 0.25, 0.10]
    return rng.choices(opciones, weights=pesos, k=1)[0]


def elegir_estado(client_idx: int, quote_idx: int) -> str:
    estados = ["Aceptada", "En progreso", "Rechazada"]
    return estados[(client_idx + quote_idx) % len(estados)]


def elegir_moneda(rng: random.Random, familia: str) -> str:
    if familia in {FAMILIES[2], FAMILIES[3], FAMILIES[4]}:
        return "USD" if rng.random() < 0.55 else "ARS"
    return "USD" if rng.random() < 0.22 else "ARS"


def costo_rango(base: float, rng: random.Random) -> float:
    factor = rng.uniform(0.88, 1.24)
    return round(base * factor, 2)


def purge_existing_demo_data() -> None:
    demo_clientes = Cliente.query.filter(Cliente.nombre.like(f"{DEMO_NAME_PREFIX}%")).all()
    demo_cliente_ids = {cliente.id for cliente in demo_clientes}
    demo_cotizaciones = (
        Cotizacion.query.filter(
            (Cotizacion.cliente.like(f"{DEMO_NAME_PREFIX}%")) | (Cotizacion.cliente_id.in_(demo_cliente_ids))
        )
        .order_by(Cotizacion.id.asc())
        .all()
    )

    demo_refs = {cot.numero_cotizacion for cot in demo_cotizaciones if cot.numero_cotizacion}
    demo_quote_ids = {cot.id for cot in demo_cotizaciones}

    if demo_quote_ids:
        Auditoria.query.filter(
            (Auditoria.tipo_entidad == "Cotización") & (Auditoria.entidad_id.in_(demo_quote_ids))
        ).delete(synchronize_session=False)
    if demo_refs:
        Auditoria.query.filter(
            (Auditoria.tipo_entidad == "Cotización") & (Auditoria.entidad_ref.in_(demo_refs))
        ).delete(synchronize_session=False)
    if demo_cliente_ids:
        Auditoria.query.filter(
            (Auditoria.tipo_entidad == "Cliente") & (Auditoria.entidad_id.in_(demo_cliente_ids))
        ).delete(synchronize_session=False)

    for cotizacion in demo_cotizaciones:
        db.session.delete(cotizacion)
    for cliente in demo_clientes:
        db.session.delete(cliente)
    db.session.commit()


def renumber_all_quotes() -> None:
    secuencias_por_anio = defaultdict(int)
    cotizaciones = Cotizacion.query.order_by(Cotizacion.fecha.asc(), Cotizacion.id.asc()).all()
    for cotizacion in cotizaciones:
        anio = (cotizacion.fecha or datetime.utcnow()).year
        secuencias_por_anio[anio] += 1
        cotizacion.numero_cotizacion = formatear_numero_cotizacion(anio, secuencias_por_anio[anio])
    db.session.commit()


def seed_demo_dataset() -> dict[str, int]:
    rng = random.Random(RANDOM_SEED)
    purge_existing_demo_data()

    created_clients = []
    cuit_num = DEMO_CUIT_START
    for (sector, subsector), specs in CLIENT_SPECS.items():
        for nombre_corto, razon_social in specs:
            cuit_num += 1
            cliente = Cliente(
                nombre=f"{DEMO_NAME_PREFIX}{nombre_corto}",
                razon_social=razon_social,
                cuit=f"30-{cuit_num:08d}-9",
                domicilio=demo_domicilio(len(created_clients) + 1, sector, subsector),
                sector=sector,
                subsector=subsector,
                email=f"{demo_slug(nombre_corto)}@demo.cuencotech.local",
                telefono=f"+54 261 {700000 + len(created_clients) * 37}",
                condicion_iva=elegir_condicion_iva(rng, sector),
            )
            db.session.add(cliente)
            created_clients.append(cliente)
    db.session.commit()

    created_quotes = []
    start_date = START_DATE
    total_days = 82
    for client_idx, cliente in enumerate(created_clients):
        family_options = FAMILY_BY_SEGMENT[(cliente.sector, cliente.subsector)]
        for quote_idx in range(3):
            fecha = start_date + timedelta(
                days=((client_idx * 3) + (quote_idx * 11)) % total_days,
                hours=8 + ((client_idx + quote_idx) % 8),
                minutes=((client_idx * 13) + (quote_idx * 17)) % 60,
            )
            familia = family_options[(client_idx + quote_idx) % len(family_options)]
            estado = elegir_estado(client_idx, quote_idx)
            moneda = elegir_moneda(rng, familia)
            cotizacion = Cotizacion(
                numero_cotizacion="PENDING",
                estado=estado,
                seguimiento_activo=False,
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
                total_neto=0.0,
                total_iva=0.0,
                total_final=0.0,
            )

            neto_total = 0.0
            iva_total = 0.0
            catalog = PRODUCT_CATALOG[familia]
            item_count = 2 + ((client_idx + quote_idx) % 3)
            for item_idx in range(item_count):
                descripcion, costo_base, iva_base = catalog[(client_idx + quote_idx + item_idx) % len(catalog)]
                cantidad = 1 + ((client_idx + item_idx + quote_idx) % 4)
                costo = costo_rango(costo_base, rng)
                extra = [5.0, 7.5, 10.0][(client_idx + item_idx) % 3]
                margen = round(0.32 + (((client_idx * 7) + (quote_idx * 5) + item_idx) % 45) / 100, 2)
                iva_item = iva_base if cliente.condicion_iva != "Exento" else 0.0
                costo_con_extra = costo * (1 + (extra / 100.0))
                precio_venta = round(costo_con_extra * (1 + margen), 2)
                subtotal_neto = cantidad * precio_venta
                monto_iva = subtotal_neto * (iva_item / 100.0)
                subtotal_final = round(subtotal_neto + monto_iva, 2)

                item = ItemCotizacion(
                    descripcion=descripcion,
                    cantidad=float(cantidad),
                    costo_unitario=costo,
                    costo_extra=extra,
                    margen=margen,
                    iva_item=iva_item,
                    precio_venta=precio_venta,
                    subtotal=subtotal_final,
                    imagen_url=None,
                )
                cotizacion.items.append(item)
                neto_total += subtotal_neto
                iva_total += monto_iva

            cotizacion.total_neto = round(neto_total, 2)
            cotizacion.total_iva = round(iva_total, 2)
            cotizacion.total_final = round(neto_total + iva_total, 2)

            db.session.add(cotizacion)
            created_quotes.append(cotizacion)

    db.session.commit()
    renumber_all_quotes()

    return {
        "clientes_demo": len(created_clients),
        "cotizaciones_demo": len(created_quotes),
        "items_demo": sum(len(cot.items) for cot in created_quotes),
    }


def print_summary() -> None:
    demo_clientes = Cliente.query.filter(Cliente.nombre.like(f"{DEMO_NAME_PREFIX}%")).all()
    demo_cliente_ids = [cliente.id for cliente in demo_clientes]
    demo_cotizaciones = (
        Cotizacion.query.filter(
            (Cotizacion.cliente.like(f"{DEMO_NAME_PREFIX}%")) | (Cotizacion.cliente_id.in_(demo_cliente_ids))
        )
        .order_by(Cotizacion.fecha.asc(), Cotizacion.id.asc())
        .all()
    )

    estados = Counter(normalizar_estado_cotizacion(cot.estado) or "En progreso" for cot in demo_cotizaciones)
    familias = Counter((cot.familia or "Sin familia") for cot in demo_cotizaciones)
    sectores = Counter((cot.cliente_ref.sector if cot.cliente_ref else "Sin sector") for cot in demo_cotizaciones)
    subsectores = Counter((cot.cliente_ref.subsector if cot.cliente_ref else "Sin subsector") for cot in demo_cotizaciones)

    print("Resumen demo")
    print("  clientes:", len(demo_clientes))
    print("  cotizaciones:", len(demo_cotizaciones))
    print("  estados:", dict(estados))
    print("  familias:", dict(familias))
    print("  sectores:", dict(sectores))
    print("  subsectores:", dict(subsectores))
    if demo_cotizaciones:
        print("  primera:", demo_cotizaciones[0].numero_cotizacion, demo_cotizaciones[0].fecha)
        print("  ultima:", demo_cotizaciones[-1].numero_cotizacion, demo_cotizaciones[-1].fecha)


def main() -> None:
    with app.app_context():
        created = seed_demo_dataset()
        print("Carga demo completada:", created)
        print_summary()


if __name__ == "__main__":
    main()
