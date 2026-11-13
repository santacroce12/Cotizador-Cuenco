"""Calculo de importes de una cotizacion.

Unica fuente de verdad para el guardado, la vista, el PDF y el Excel (el
formulario repite la misma logica en JavaScript para mostrar los totales en
vivo). Reglas:

1. El precio unitario neto se redondea a 2 decimales y todo lo demas se
   calcula a partir de ese valor: cantidad x unitario siempre da el total de
   la linea, como en una factura.
2. Responsable Inscrito (precios sin IVA): el IVA se calcula sobre el neto
   acumulado de cada alicuota y se redondea por alicuota.
   Resto (precios con IVA incluido): el precio final unitario se redondea y el
   total de la linea es cantidad x precio final; la suma de lineas es el total.
3. La bonificacion de cierre se carga como un monto final (con IVA) y se
   reparte en forma proporcional entre el neto y el IVA de cada alicuota, como
   un descuento comercial antes de impuestos. Asi el IVA informado es el que
   corresponde a lo que el cliente paga y la ganancia descuenta solo la parte
   neta de la bonificacion.

Se usa Decimal para que los redondeos sean exactos (mitad hacia arriba).
"""

from decimal import ROUND_HALF_UP, Decimal

CENTAVO = Decimal("0.01")
CIEN = Decimal("100")
CERO = Decimal("0")


def _d(valor):
    return Decimal(str(valor or 0))


def redondear(valor):
    return _d(valor).quantize(CENTAVO, rounding=ROUND_HALF_UP)


def costo_con_imprevistos(costo, imprevisto_item_pct=0, imprevisto_general_pct=0):
    """Costo x (1 + Imprevisto del item%) x (1 + Imprevisto general%), sin redondear.

    El imprevisto del item cubre riesgos de ese producto; el general, de todo el
    proyecto, y se aplica sobre el costo ya con los imprevistos de cada item.
    """
    return _d(costo) * (1 + _d(imprevisto_item_pct) / CIEN) * (1 + _d(imprevisto_general_pct) / CIEN)


def precio_unitario_neto(costo, extra_pct, margen_pct, descuento_pct, imprevisto_general_pct=0):
    """Costo con imprevistos x (1 + Margen%) x (1 - Desc%), redondeado a centavos.

    extra_pct es el imprevisto del item (antes llamado "Otros %").
    """
    precio = (
        costo_con_imprevistos(costo, extra_pct, imprevisto_general_pct)
        * (1 + _d(margen_pct) / CIEN)
        * (1 - _d(descuento_pct) / CIEN)
    )
    return float(redondear(precio))


def calcular_totales(lineas, precio_incluye_iva, bonificacion_solicitada=0):
    """lineas: iterable de dicts con unitario_neto, cantidad, iva_pct y (opcional) carga_pct.

    Devuelve importes en float, ya redondeados a centavos.
    """
    detalle_lineas = []
    neto = CERO
    neto_por_tasa = {}
    iva_por_tasa = {}
    carga = CERO

    for linea in lineas:
        unitario = redondear(linea["unitario_neto"])
        cantidad = _d(linea["cantidad"])
        tasa = _d(linea["iva_pct"])
        neto_linea = redondear(unitario * cantidad)
        if precio_incluye_iva:
            unitario_final = redondear(unitario * (1 + tasa / CIEN))
            total_linea = redondear(unitario_final * cantidad)
            iva_linea = total_linea - neto_linea
            iva_por_tasa[tasa] = iva_por_tasa.get(tasa, CERO) + iva_linea
        else:
            unitario_final = unitario
            iva_linea = neto_linea * tasa / CIEN
            total_linea = redondear(neto_linea + iva_linea)
        neto += neto_linea
        neto_por_tasa[tasa] = neto_por_tasa.get(tasa, CERO) + neto_linea
        carga += neto_linea * _d(linea.get("carga_pct")) / CIEN
        detalle_lineas.append(
            {
                "unitario_neto": float(unitario),
                "unitario_final": float(unitario_final),
                "neto": float(neto_linea),
                "total": float(total_linea),
            }
        )

    if not precio_incluye_iva:
        iva_por_tasa = {tasa: redondear(base * tasa / CIEN) for tasa, base in neto_por_tasa.items()}

    iva_bruto = sum(iva_por_tasa.values(), CERO)
    bruto = neto + iva_bruto
    bonificacion = min(max(_d(bonificacion_solicitada), CERO), bruto)
    bonificacion = redondear(bonificacion)
    factor = (bonificacion / bruto) if bruto else CERO
    total = redondear(bruto - bonificacion)

    iva_final = {tasa: redondear(monto * (1 - factor)) for tasa, monto in iva_por_tasa.items()}
    iva_total = sum(iva_final.values(), CERO)
    # La parte neta de la bonificacion absorbe el centavo de redondeo para
    # que neto - bonificacion neta + IVA sea exactamente el total.
    bonificacion_neta = neto + iva_total - total
    neto_bonificado = neto - bonificacion_neta

    return {
        "lineas": detalle_lineas,
        "neto": float(neto),
        "iva_bruto": float(iva_bruto),
        "bruto": float(bruto),
        "bonificacion": float(bonificacion),
        "bonificacion_neta": float(bonificacion_neta),
        "neto_bonificado": float(neto_bonificado),
        "iva_por_tasa": {float(t): float(m) for t, m in sorted(iva_final.items(), key=lambda x: -x[0]) if t > 0},
        "iva_total": float(iva_total),
        "total": float(total),
        "carga_fiscal": float(redondear(carga * (1 - factor))),
    }
