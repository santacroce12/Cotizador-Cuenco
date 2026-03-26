from __future__ import annotations

import sys
import unicodedata
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from sqlalchemy import func

from app import Cotizacion, app, db


FAMILIA_KEYWORDS = {
    "SEGURIDAD URBANA": (
        "camara",
        "cctv",
        "hikvision",
        "poe",
        "switch",
        "nvr",
        "dvr",
        "domo",
        "bullet",
        "seguridad",
    ),
    "PARKING": (
        "parking",
        "estacionamiento",
        "barrera",
        "ticketera",
        "ticket",
        "acceso vehicular",
    ),
    "TRANSPORTE INTELIGENTE": (
        "transporte",
        "trafico",
        "transito",
        "radar",
        "semaforo",
        "movilidad",
        "via publica",
    ),
    "CONECTIVIDAD SATELITAL": (
        "satelital",
        "starlink",
        "vsat",
        "enlace",
        "antena",
        "radioenlace",
        "microonda",
    ),
    "SALAS DE CONTROL": (
        "sala de control",
        "control room",
        "videowall",
        "video wall",
        "monitor",
        "pantalla",
        "workstation",
    ),
    "SMART CITIES": (
        "smart city",
        "smart cities",
        "iot",
        "sensor",
        "ciudad",
        "urbano",
    ),
}

DEFAULT_FAMILIA = "SMART CITIES"


def normalizar(texto: str) -> str:
    base = unicodedata.normalize("NFKD", texto or "")
    ascii_only = base.encode("ascii", "ignore").decode("ascii")
    return ascii_only.lower().strip()


def inferir_familia(cotizacion: Cotizacion) -> str:
    piezas = [cotizacion.cliente or "", cotizacion.cliente_razon_social or ""]
    piezas.extend(item.descripcion or "" for item in cotizacion.items)
    corpus = " | ".join(normalizar(parte) for parte in piezas if parte)

    for familia, keywords in FAMILIA_KEYWORDS.items():
        if any(keyword in corpus for keyword in keywords):
            return familia
    return DEFAULT_FAMILIA


def main() -> None:
    with app.app_context():
        faltantes = (
            Cotizacion.query.filter((Cotizacion.familia == None) | (func.trim(Cotizacion.familia) == ""))  # noqa: E711
            .order_by(Cotizacion.id.asc())
            .all()
        )

        if not faltantes:
            print("No hay cotizaciones para backfill.")
            return

        resumen: dict[str, int] = {}
        for cotizacion in faltantes:
            familia = inferir_familia(cotizacion)
            cotizacion.familia = familia
            resumen[familia] = resumen.get(familia, 0) + 1
            print(f"{cotizacion.id} {cotizacion.numero_cotizacion or '-'} -> {familia}")

        db.session.commit()

        print("---")
        print(f"actualizadas: {len(faltantes)}")
        for familia, cantidad in sorted(resumen.items(), key=lambda item: item[0]):
            print(f"{familia}: {cantidad}")


if __name__ == "__main__":
    main()
