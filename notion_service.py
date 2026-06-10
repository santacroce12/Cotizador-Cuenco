import os
from datetime import datetime, timedelta
from typing import Optional

import requests


NOTION_VERSION = "2022-06-28"


def notion_enabled() -> bool:
    return str(os.getenv("NOTION_ENABLED", "")).strip().lower() in ("1", "true", "yes")


def notion_headers() -> dict:
    token = os.getenv("NOTION_TOKEN", "").strip()
    return {
        "Authorization": f"Bearer {token}",
        "Notion-Version": NOTION_VERSION,
        "Content-Type": "application/json",
    }


def notion_request(method: str, url: str, **kwargs) -> Optional[dict]:
    if not notion_enabled():
        return None

    token = os.getenv("NOTION_TOKEN", "").strip()
    if not token:
        print("[notion] NOTION_TOKEN no configurado")
        return None

    try:
        response = requests.request(
            method,
            url,
            headers=notion_headers(),
            timeout=15,
            **kwargs,
        )

        if response.status_code >= 400:
            print(f"[notion] Error {response.status_code}: {response.text}")
            return None

        if not response.text:
            return {}

        return response.json()

    except Exception as exc:
        print(f"[notion] Error de conexion: {exc}")
        return None


def notion_query_database(database_id: str, filter_payload: dict) -> Optional[dict]:
    url = f"https://api.notion.com/v1/databases/{database_id}/query"
    return notion_request("POST", url, json={"filter": filter_payload})


def notion_create_page(database_id: str, properties: dict) -> Optional[dict]:
    url = "https://api.notion.com/v1/pages"
    payload = {
        "parent": {"database_id": database_id},
        "properties": properties,
    }
    return notion_request("POST", url, json=payload)


def notion_update_page(page_id: str, properties: dict) -> Optional[dict]:
    url = f"https://api.notion.com/v1/pages/{page_id}"
    return notion_request("PATCH", url, json={"properties": properties})


def text_prop(value: object) -> dict:
    value = "" if value is None else str(value)
    return {"rich_text": [{"text": {"content": value[:2000]}}]}


def title_prop(value: object) -> dict:
    value = "" if value is None else str(value)
    return {"title": [{"text": {"content": value[:2000]}}]}


def select_prop(value: object) -> dict:
    value = "" if value is None else str(value).strip()
    if not value:
        return {"select": None}
    return {"select": {"name": value}}


def email_prop(value: object) -> dict:
    value = "" if value is None else str(value).strip()
    return {"email": value or None}


def phone_prop(value: object) -> dict:
    value = "" if value is None else str(value).strip()
    return {"phone_number": value or None}


def url_prop(value: object) -> dict:
    value = "" if value is None else str(value).strip()
    return {"url": value or None}


def number_prop(value: object) -> dict:
    try:
        return {"number": float(value)}
    except (TypeError, ValueError):
        return {"number": None}


def date_prop(value: object = None) -> dict:
    if value is None:
        value = datetime.utcnow()

    if isinstance(value, datetime):
        return {"date": {"start": value.date().isoformat()}}

    return {"date": {"start": str(value)}}


def relation_prop(page_id: Optional[str]) -> dict:
    if not page_id:
        return {"relation": []}
    return {"relation": [{"id": page_id}]}


def buscar_page_por_propiedad(database_id: str, property_name: str, value: object) -> Optional[str]:
    value = "" if value is None else str(value).strip()
    if not value:
        return None

    data = notion_query_database(
        database_id,
        {
            "property": property_name,
            "rich_text": {
                "equals": value,
            },
        },
    )

    if not data:
        return None

    results = data.get("results") or []
    if not results:
        return None

    return results[0].get("id")


def map_estado_cotizacion_a_notion(estado: str) -> str:
    estado = (estado or "").strip().lower()

    if estado == "aceptada":
        return "Aprobado"

    if estado == "rechazada":
        return "Rechazado"

    if estado == "en progreso":
        return "Enviado"

    return "Enviado"


def map_familia_a_notion(familia: str) -> str:
    familia = (familia or "").strip().upper()
    permitidas = {
        "SEGURIDAD URBANA",
        "PARKING",
        "TRANSPORTE INTELIGENTE",
        "CONECTIVIDAD SATELITAL",
        "SALAS DE CONTROL",
        "SMART CITIES",
    }
    return familia if familia in permitidas else "Otro"


def map_estado_cotizador(valor: str) -> str:
    valor = (valor or "").strip().lower()

    if valor == "aceptada":
        return "Aceptada"

    if valor == "rechazada":
        return "Rechazada"

    if valor == "en progreso":
        return "En progreso"

    return "En progreso"


def map_etapa_comercial_desde_estado(valor: str) -> str:
    valor = (valor or "").strip().lower()

    if valor == "aceptada":
        return "Ganado"

    if valor == "rechazada":
        return "Perdido"

    return "Cotizado"


def map_estado_seguimiento_desde_estado(valor: str) -> str:
    valor = (valor or "").strip().lower()

    if valor == "aceptada":
        return "Ganado"

    if valor == "rechazada":
        return "Perdido"

    return "Pendiente"


def cotizacion_esta_aceptada(cotizacion) -> bool:
    return ((getattr(cotizacion, "estado", "") or "").strip().lower() == "aceptada")


def construir_link_cotizador(cotizacion_id: int) -> str:
    base_url = (os.getenv("APP_BASE_URL") or "").strip().rstrip("/")
    if not base_url:
        return ""
    return f"{base_url}/cotizacion/{cotizacion_id}"


def sync_cliente_to_notion(cliente) -> Optional[str]:
    if not notion_enabled() or not cliente:
        return None

    database_id = os.getenv("NOTION_CLIENTES_DB_ID", "").strip()
    if not database_id:
        print("[notion] NOTION_CLIENTES_DB_ID no configurado")
        return None

    id_sistema = str(cliente.id)
    page_id = buscar_page_por_propiedad(database_id, "ID Sistema", id_sistema)

    tipo_cliente = "Administracion"
    if (cliente.sector or "").lower() == "privado":
        tipo_cliente = "Seguridad electronica"

    properties = {
        "Cliente": title_prop(cliente.nombre or cliente.razon_social or f"Cliente {cliente.id}"),
        "ID Sistema": text_prop(id_sistema),
        "CUIT": text_prop(cliente.cuit or ""),
        "Tipo de cliente": select_prop(tipo_cliente),
        "Estado": select_prop("Activo"),
        "Contacto principal": text_prop(cliente.razon_social or ""),
        "Telefono": phone_prop(cliente.telefono or ""),
        "Email": email_prop(cliente.email or ""),
        "Direccion": text_prop(cliente.domicilio or ""),
        "Notas": text_prop(
            f"Sector: {cliente.sector or ''}. Subsector: {cliente.subsector or ''}. "
            f"Condicion IVA: {cliente.condicion_iva or ''}."
        ),
        "Origen": select_prop("Cotizador"),
        "Última sync": date_prop(),
    }

    if page_id:
        result = notion_update_page(page_id, properties)
    else:
        result = notion_create_page(database_id, properties)

    if not result:
        return page_id

    return result.get("id") or page_id


def sync_seguimiento_comercial_from_cotizacion(
    cotizacion,
    cliente_page_id: Optional[str] = None,
    presupuesto_page_id: Optional[str] = None,
    aviso_cotizador: bool = False,
) -> Optional[str]:
    if not notion_enabled() or not cotizacion:
        return None

    seguimiento_db_id = os.getenv("NOTION_SEGUIMIENTO_DB_ID", "").strip()
    if not seguimiento_db_id:
        print("[notion] NOTION_SEGUIMIENTO_DB_ID no configurado")
        return None

    if not cliente_page_id and getattr(cotizacion, "cliente_ref", None):
        cliente_page_id = sync_cliente_to_notion(cotizacion.cliente_ref)

    id_cotizacion = str(cotizacion.id)
    numero = cotizacion.numero_cotizacion or f"COT-{cotizacion.id}"
    link = construir_link_cotizador(cotizacion.id)

    page_id = buscar_page_por_propiedad(seguimiento_db_id, "Numero cotizacion", numero)

    estado_cotizador = map_estado_cotizador(cotizacion.estado)
    etapa = map_etapa_comercial_desde_estado(cotizacion.estado)
    estado_seguimiento = map_estado_seguimiento_desde_estado(cotizacion.estado)

    fecha_base = cotizacion.fecha or datetime.utcnow()
    proximo_seguimiento = fecha_base + timedelta(days=7)

    if estado_cotizador == "Aceptada":
        proxima_accion = "Coordinar pase a trabajo operativo / ejecucion."
        probabilidad = 100
    elif estado_cotizador == "Rechazada":
        proxima_accion = "Registrar motivo de perdida y evaluar alternativa futura."
        probabilidad = 0
    else:
        proxima_accion = "Contactar al cliente para consultar estado de aprobacion."
        probabilidad = 50

    titulo = f"{numero} - {cotizacion.cliente or 'Cliente'}"

    properties = {
        "Seguimiento": title_prop(titulo),
        "Cliente": relation_prop(cliente_page_id),
        "Presupuesto": relation_prop(presupuesto_page_id),
        "Estado": select_prop(estado_seguimiento),
        "Etapa comercial": select_prop(etapa),
        "Tipo de oportunidad": select_prop("Cotizacion"),
        "Estado cotizador": select_prop(estado_cotizador),
        "Numero cotizacion": text_prop(numero),
        "Link cotizador": url_prop(link),
        "Monto estimado": number_prop(cotizacion.total_final or 0),
        "Moneda": select_prop((cotizacion.moneda or "ARS").upper()),
        "Probabilidad": number_prop(probabilidad),
        "Prioridad": select_prop("Media"),
        "Responsable": text_prop("Comercial"),
        "Canal": select_prop("WhatsApp"),
        "Proximo seguimiento": date_prop(proximo_seguimiento),
        "Proxima accion": text_prop(proxima_accion),
        "Dolor / necesidad": text_prop(cotizacion.observacion_cliente or ""),
        "Objecion principal": text_prop("Pendiente de relevar por comercial."),
        "Resumen comercial": text_prop(
            f"Cotizacion {numero} para {cotizacion.cliente or ''}. "
            f"Estado en cotizador: {estado_cotizador}. "
            f"Familia: {cotizacion.familia or ''}. "
            f"Total: {(cotizacion.moneda or 'ARS').upper()} {cotizacion.total_final or 0:,.2f}."
        ),
        "Origen": select_prop("Cotizador"),
    }

    if aviso_cotizador:
        properties["Ultimo aviso cotizador"] = date_prop()
        properties["Avisos recibidos"] = number_prop(1)

    if page_id:
        result = notion_update_page(page_id, properties)
    else:
        result = notion_create_page(seguimiento_db_id, properties)

    if not result:
        return page_id

    return result.get("id") or page_id


def sync_cotizacion_to_notion(cotizacion) -> Optional[str]:
    if not notion_enabled() or not cotizacion:
        return None

    presupuestos_db_id = os.getenv("NOTION_PRESUPUESTOS_DB_ID", "").strip()
    trabajos_db_id = os.getenv("NOTION_TRABAJOS_DB_ID", "").strip()

    if not presupuestos_db_id:
        print("[notion] NOTION_PRESUPUESTOS_DB_ID no configurado")
        return None

    cliente_page_id = None
    if getattr(cotizacion, "cliente_ref", None):
        cliente_page_id = sync_cliente_to_notion(cotizacion.cliente_ref)

    id_cotizacion = str(cotizacion.id)
    numero = cotizacion.numero_cotizacion or f"COT-{cotizacion.id}"
    link = construir_link_cotizador(cotizacion.id)

    page_id = buscar_page_por_propiedad(presupuestos_db_id, "ID Cotizacion", id_cotizacion)

    properties = {
        "Presupuesto": title_prop(numero),
        "ID Cotizacion": text_prop(id_cotizacion),
        "Numero Cotizacion": text_prop(numero),
        "Cliente": relation_prop(cliente_page_id),
        "Estado": select_prop(map_estado_cotizacion_a_notion(cotizacion.estado)),
        "Monto": number_prop(cotizacion.total_final or 0),
        "Moneda": select_prop((cotizacion.moneda or "ARS").upper()),
        "Familia": select_prop(map_familia_a_notion(cotizacion.familia)),
        "Fecha enviado": date_prop(cotizacion.fecha),
        "Fecha seguimiento": date_prop((cotizacion.fecha or datetime.utcnow()) + timedelta(days=7)),
        "Responsable": text_prop("Cotizador Cuenco"),
        "Link Cotizador": url_prop(link),
        "Notas": text_prop(
            f"Cliente: {cotizacion.cliente or ''}. "
            f"Forma de pago: {cotizacion.forma_pago or ''}. "
            f"Condicion IVA: {cotizacion.condicion_iva or ''}."
        ),
        "Origen": select_prop("Cotizador"),
        "Ultima sync": date_prop(),
    }

    if page_id:
        result = notion_update_page(page_id, properties)
    else:
        result = notion_create_page(presupuestos_db_id, properties)

    presupuesto_page_id = (result or {}).get("id") or page_id

    sync_seguimiento_comercial_from_cotizacion(
        cotizacion,
        cliente_page_id=cliente_page_id,
        presupuesto_page_id=presupuesto_page_id,
    )

    if trabajos_db_id and cotizacion_esta_aceptada(cotizacion):
        sync_trabajo_from_cotizacion(cotizacion, cliente_page_id)

    return presupuesto_page_id


def sync_trabajo_from_cotizacion(cotizacion, cliente_page_id: Optional[str] = None) -> Optional[str]:
    trabajos_db_id = os.getenv("NOTION_TRABAJOS_DB_ID", "").strip()
    if not trabajos_db_id or not cotizacion or not cotizacion_esta_aceptada(cotizacion):
        return None

    id_cotizacion = str(cotizacion.id)
    numero = cotizacion.numero_cotizacion or f"COT-{cotizacion.id}"
    link = construir_link_cotizador(cotizacion.id)

    page_id = buscar_page_por_propiedad(trabajos_db_id, "ID Cotizacion", id_cotizacion)

    properties = {
        "Trabajo": title_prop(f"{numero} - {cotizacion.cliente or 'Cliente'}"),
        "Cliente": relation_prop(cliente_page_id),
        "ID Cotizacion": text_prop(id_cotizacion),
        "Numero Cotizacion": text_prop(numero),
        "Tipo de servicio": select_prop("Otro"),
        "Estado": select_prop("Aprobado"),
        "Prioridad": select_prop("Media"),
        "Responsable": text_prop("Cotizador Cuenco"),
        "Fecha inicio": date_prop(cotizacion.fecha),
        "Link Cotizador": url_prop(link),
        "Proximo paso": text_prop("Revisar estado de la cotizacion y coordinar seguimiento."),
        "Notas": text_prop(f"Trabajo generado automaticamente desde cotizacion {numero}."),
        "Origen": select_prop("Cotizador"),
        "Ultima sync": date_prop(),
    }

    if page_id:
        result = notion_update_page(page_id, properties)
    else:
        result = notion_create_page(trabajos_db_id, properties)

    if not result:
        return page_id

    return result.get("id") or page_id
