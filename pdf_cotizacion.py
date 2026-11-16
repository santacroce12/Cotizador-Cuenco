"""Genera el PDF del presupuesto en el servidor.

A diferencia de "Imprimir > Guardar como PDF" del navegador, el resultado es
identico desde cualquier PC: mismos margenes, tipografia, saltos de pagina y
numeracion. Usa solo ReportLab (Python puro), sin dependencias del sistema.
"""

from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.platypus import (
    Image,
    KeepTogether,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

TINTA = colors.HexColor("#0f172a")
TEXTO = colors.HexColor("#334155")
GRIS = colors.HexColor("#64748b")
LINEA = colors.HexColor("#cbd5e1")
FONDO = colors.HexColor("#f1f5f9")
CABECERA = colors.HexColor("#1e293b")

MARGEN = 15 * mm
ANCHO_UTIL = A4[0] - 2 * MARGEN


def _estilo(nombre, **kwargs):
    base = {"fontName": "Helvetica", "fontSize": 9, "leading": 12, "textColor": TEXTO}
    base.update(kwargs)
    return ParagraphStyle(nombre, **base)


ESTILOS = {
    "normal": _estilo("normal"),
    "chico": _estilo("chico", fontSize=8, leading=10, textColor=GRIS),
    "empresa": _estilo("empresa", fontName="Helvetica-Bold", fontSize=12, leading=15, textColor=TINTA),
    "titulo": _estilo("titulo", fontName="Helvetica-Bold", fontSize=16, leading=19, textColor=TINTA, alignment=TA_RIGHT),
    "derecha": _estilo("derecha", alignment=TA_RIGHT),
    "seccion": _estilo("seccion", fontName="Helvetica-Bold", fontSize=8, leading=10, textColor=GRIS),
    "cliente": _estilo("cliente", fontName="Helvetica-Bold", fontSize=11, leading=14, textColor=TINTA),
    "th": _estilo("th", fontName="Helvetica-Bold", fontSize=8, leading=10, textColor=colors.white, alignment=TA_CENTER),
    "td": _estilo("td", fontSize=8.5, leading=11),
    "td_bold": _estilo("td_bold", fontName="Helvetica-Bold", fontSize=8.5, leading=11, textColor=TINTA),
    "td_num": _estilo("td_num", fontSize=8.5, leading=11, alignment=TA_RIGHT),
    "td_centro": _estilo("td_centro", fontSize=8.5, leading=11, alignment=TA_CENTER),
    "total_label": _estilo("total_label", fontName="Helvetica-Bold", fontSize=9),
    "total_num": _estilo("total_num", fontSize=9, alignment=TA_RIGHT),
    "final_label": _estilo("final_label", fontName="Helvetica-Bold", fontSize=12, leading=15, textColor=TINTA),
    "final_num": _estilo("final_num", fontName="Helvetica-Bold", fontSize=12, leading=15, textColor=TINTA, alignment=TA_RIGHT),
    "firma_nombre": _estilo("firma_nombre", fontName="Helvetica-Bold", fontSize=10, leading=13, textColor=TINTA, alignment=TA_CENTER),
    "firma": _estilo("firma", fontSize=8.5, leading=11, textColor=GRIS, alignment=TA_CENTER),
    "lista": _estilo("lista", fontSize=8.5, leading=11, leftIndent=8, bulletIndent=0),
}


def limpiar_texto(texto):
    """Quita caracteres que las fuentes estandar del PDF no pueden dibujar.

    Helvetica cubre el espanol completo (tildes, enie, signos), pero no emoji
    ni simbolos graficos: esos saldrian como un cuadrado negro.
    """
    limpio = []
    for caracter in str(texto or ""):
        codigo = ord(caracter)
        if codigo > 0xFFFF or 0x2600 <= codigo <= 0x27BF or 0xFE00 <= codigo <= 0xFE0F or codigo == 0x200D:
            continue
        limpio.append(caracter)
    return "".join(limpio)


def _p(texto, estilo="normal"):
    """Paragraph con el texto escapado (los datos vienen del usuario)."""
    contenido = escape(limpiar_texto(texto)).replace("\n", "<br/>")
    return Paragraph(contenido, ESTILOS[estilo] if isinstance(estilo, str) else estilo)


def formatear_importe(valor, simbolo):
    """Formato argentino: US$ 1.234,56"""
    numero = float(valor or 0.0)
    texto = f"{abs(numero):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{'-' if numero < 0 else ''}{simbolo} {texto}"


def formatear_porcentaje(valor):
    texto = f"{float(valor or 0):.2f}".rstrip("0").rstrip(".")
    return f"{texto.replace('.', ',')}%"


def _imagen(ruta, ancho, alto):
    """Imagen escalada para entrar en ancho x alto, o None si no se puede leer."""
    if not ruta:
        return None
    try:
        lector = ImageReader(str(ruta))
        w, h = lector.getSize()
    except Exception:
        return None
    escala = min(ancho / float(w), alto / float(h))
    return Image(str(ruta), width=w * escala, height=h * escala)


PREFIJO_UPLOADS = "uploads/productos/"


def _ruta_imagen_item(imagen_url, static_dir, uploads_dir=None):
    """Ruta local de la foto de un item, sin salir de las carpetas permitidas."""
    if not imagen_url or str(imagen_url).startswith(("http://", "https://")):
        return None
    imagen_url = str(imagen_url)
    candidatos = []
    if uploads_dir and imagen_url.startswith(PREFIJO_UPLOADS):
        candidatos.append((Path(uploads_dir), imagen_url[len(PREFIJO_UPLOADS):]))
    candidatos.append((Path(static_dir), imagen_url))
    for base, relativa in candidatos:
        ruta = (base / relativa).resolve()
        try:
            ruta.relative_to(base.resolve())
        except ValueError:
            continue
        if ruta.is_file():
            return ruta
    return None


def generar_pdf_cotizacion(contexto, *, static_dir, fecha_texto, empresa, nota_facturacion_usd="", uploads_dir=None):
    """Devuelve los bytes del PDF.

    contexto: el mismo que usa la vista HTML (construir_contexto_documento_cotizacion).
    empresa: dict con nombre, cuit y domicilio.
    """
    cot = contexto["cot"]
    firma = contexto["quote_signature"]
    precio_incluye_iva = contexto["precio_incluye_iva"]
    mostrar_descuento = contexto["mostrar_descuento"]
    simbolo = "$" if (cot.moneda or "").upper() == "ARS" else "US$"
    numero = cot.numero_cotizacion or f"{cot.id:04d}"
    moneda_texto = "Pesos argentinos (ARS)" if simbolo == "$" else "Dólares estadounidenses (USD)"

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=MARGEN,
        rightMargin=MARGEN,
        topMargin=MARGEN,
        bottomMargin=18 * mm,
        title=f"Presupuesto {numero}",
        author=empresa.get("nombre", ""),
        subject=f"Presupuesto {numero} - {cot.cliente or ''}",
    )
    historia = []

    # ---- Encabezado ----
    static_path = Path(static_dir)
    logo = _imagen(static_path / "LOGO color.png", 34 * mm, 18 * mm) or _imagen(static_path / "logo.png", 34 * mm, 18 * mm)
    datos_empresa = [
        _p(empresa.get("nombre"), "empresa"),
        _p(f"CUIT: {empresa.get('cuit', '')}", "chico"),
        _p(empresa.get("domicilio"), "chico"),
    ]
    izquierda = Table([[logo or "", datos_empresa]], colWidths=[38 * mm, None])
    izquierda.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    derecha = [
        _p("PRESUPUESTO", "titulo"),
        _p(f"N° {numero}", _estilo("num", fontName="Helvetica-Bold", fontSize=11, leading=14, alignment=TA_RIGHT, textColor=TINTA)),
        Spacer(1, 2),
        _p(f"Fecha: {fecha_texto}", "derecha"),
        _p(f"Moneda: {moneda_texto}", "derecha"),
    ]
    encabezado = Table([[izquierda, derecha]], colWidths=[ANCHO_UTIL * 0.6, ANCHO_UTIL * 0.4])
    encabezado.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("LINEBELOW", (0, 0), (-1, 0), 1.2, TINTA),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    historia += [encabezado, Spacer(1, 6 * mm)]

    # ---- Cliente ----
    lineas_cliente = [_p("CLIENTE", "seccion"), _p(cot.cliente_razon_social or cot.cliente, "cliente")]
    detalle_cliente = []
    if cot.cliente_cuit:
        detalle_cliente.append(f"CUIT: {cot.cliente_cuit}")
    if cot.condicion_iva:
        detalle_cliente.append(f"Condición IVA: {cot.condicion_iva}")
    if detalle_cliente:
        lineas_cliente.append(_p("   |   ".join(detalle_cliente), "chico"))
    if cot.cliente_contacto:
        lineas_cliente.append(_p(f"At.: {cot.cliente_contacto}"))

    lineas_condiciones = [_p("FORMA DE PAGO", "seccion"), _p(cot.forma_pago or "A convenir")]
    if cot.observacion_cliente:
        lineas_condiciones += [Spacer(1, 4), _p("REFERENCIA", "seccion"), _p(cot.observacion_cliente)]

    bloque_cliente = Table([[lineas_cliente, lineas_condiciones]], colWidths=[ANCHO_UTIL * 0.6, ANCHO_UTIL * 0.4])
    bloque_cliente.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BACKGROUND", (0, 0), (-1, -1), FONDO),
                ("BOX", (0, 0), (-1, -1), 0.5, LINEA),
                ("LINEBEFORE", (1, 0), (1, 0), 0.5, LINEA),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    historia += [bloque_cliente, Spacer(1, 6 * mm)]

    # ---- Items ----
    imagenes = [_imagen(_ruta_imagen_item(item.imagen_url, static_dir, uploads_dir), 13 * mm, 13 * mm) for item in cot.items]
    con_imagenes = any(imagenes)
    encabezados = ["Producto", "Descripción", "Cant."]
    anchos = [50 * mm, None, 12 * mm]
    if con_imagenes:
        encabezados.insert(0, "")
        anchos.insert(0, 16 * mm)
    if mostrar_descuento:
        encabezados.append("Desc.")
        anchos.append(13 * mm)
    encabezados += ["P. unitario" if not precio_incluye_iva else "Precio", "Total"]
    anchos += [24 * mm, 26 * mm]
    fijo = sum(a for a in anchos if a)
    anchos = [a if a else ANCHO_UTIL - fijo for a in anchos]

    filas = [[_p(h, "th") for h in encabezados]]
    totales = contexto["totales"]
    for item, imagen, linea in zip(cot.items, imagenes, totales["lineas"]):
        cantidad = int(item.cantidad or 0)
        if precio_incluye_iva:
            unitario, total = linea["unitario_final"], linea["total"]
        else:
            unitario, total = linea["unitario_neto"], linea["neto"]
        fila = [
            _p(item.descripcion, "td_bold"),
            _p((item.detalle or "").strip() or "-", "td"),
            _p(str(cantidad), "td_centro"),
        ]
        if con_imagenes:
            fila.insert(0, imagen or "")
        if mostrar_descuento:
            fila.append(_p(formatear_porcentaje(item.descuento_pct) if (item.descuento_pct or 0) > 0 else "-", "td_centro"))
        fila += [_p(formatear_importe(unitario, simbolo), "td_num"), _p(formatear_importe(total, simbolo), "td_num")]
        filas.append(fila)

    tabla_items = Table(filas, colWidths=anchos, repeatRows=1)
    estilo_items = [
        ("ALIGN", (0, 1), (0, -1), "CENTER" if con_imagenes else "LEFT"),
        ("BACKGROUND", (0, 0), (-1, 0), CABECERA),
        ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
        ("VALIGN", (0, 1), (-1, -1), "TOP"),
        ("LINEBELOW", (0, 1), (-1, -1), 0.4, LINEA),
        ("BOX", (0, 0), (-1, -1), 0.5, LINEA),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]
    for indice in range(2, len(filas), 2):
        estilo_items.append(("BACKGROUND", (0, indice), (-1, indice), FONDO))
    tabla_items.setStyle(TableStyle(estilo_items))
    historia += [tabla_items, Spacer(1, 5 * mm)]

    # ---- Totales ----
    filas_totales = []
    hay_bonificacion = totales["bonificacion"] > 0

    def fila_total(etiqueta, importe):
        filas_totales.append([_p(etiqueta, "total_label"), _p(formatear_importe(importe, simbolo), "total_num")])

    if not precio_incluye_iva:
        # Factura A: la bonificacion es un descuento sobre el neto, antes del IVA.
        fila_total("Subtotal neto", totales["neto"])
        if hay_bonificacion:
            fila_total("Bonificación de cierre", -totales["bonificacion_neta"])
            fila_total("Neto bonificado", totales["neto_bonificado"])
        for linea in contexto["desglose_iva"]:
            fila_total(f"IVA {formatear_porcentaje(linea['rate'])}", linea["amount"])
        fila_total("IVA total", totales["iva_total"])
    elif hay_bonificacion:
        fila_total("Subtotal", totales["bruto"])
        fila_total("Bonificación de cierre", -totales["bonificacion"])
    fila_final = len(filas_totales)
    filas_totales.append([_p("TOTAL", "final_label"), _p(formatear_importe(totales["total"], simbolo), "final_num")])
    if precio_incluye_iva:
        filas_totales.append([_p("Los precios incluyen IVA.", "chico"), ""])

    tabla_totales = Table(filas_totales, colWidths=[45 * mm, 38 * mm])
    estilo_totales = [
        ("BACKGROUND", (0, 0), (-1, -1), FONDO),
        ("BOX", (0, 0), (-1, -1), 0.5, LINEA),
        ("LINEABOVE", (0, fila_final), (-1, fila_final), 1, TINTA),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    if precio_incluye_iva:
        estilo_totales.append(("SPAN", (0, len(filas_totales) - 1), (-1, len(filas_totales) - 1)))
    tabla_totales.setStyle(TableStyle(estilo_totales))
    contenedor_totales = Table([["", tabla_totales]], colWidths=[ANCHO_UTIL - 83 * mm, 83 * mm])
    contenedor_totales.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    historia += [KeepTogether(contenedor_totales), Spacer(1, 7 * mm)]

    # ---- Condiciones y firma ----
    condiciones = []
    if contexto.get("validez"):
        condiciones.append(contexto["validez"]["texto"])
    condiciones += list(contexto["condiciones_cotizacion"])
    if firma.get("footer_note"):
        condiciones.append(firma["footer_note"])
    condiciones.append("La aceptación de esta cotización confirma cantidades, moneda y condiciones indicadas.")
    if simbolo == "US$" and nota_facturacion_usd:
        condiciones.append(nota_facturacion_usd)

    columna_condiciones = [_p("CONDICIONES COMERCIALES", "seccion"), Spacer(1, 3)]
    columna_condiciones += [Paragraph(escape(limpiar_texto(c)), ESTILOS["lista"], bulletText="•") for c in condiciones]

    columna_firma = [_p("EMITIDO POR", _estilo("sec_c", fontName="Helvetica-Bold", fontSize=8, textColor=GRIS, alignment=TA_CENTER))]
    imagen_firma = _imagen(static_path / firma["image"], 40 * mm, 16 * mm) if firma.get("image") else None
    columna_firma += [imagen_firma] if imagen_firma else [Spacer(1, 14 * mm)]
    linea_firma = Table([[""]], colWidths=[50 * mm])
    linea_firma.setStyle(TableStyle([("LINEABOVE", (0, 0), (-1, -1), 0.8, GRIS)]))
    columna_firma += [linea_firma, _p(firma.get("name"), "firma_nombre")]
    for dato in (firma.get("role"), firma.get("email"), firma.get("phone"), cot.razon_social or empresa.get("nombre")):
        if dato:
            columna_firma.append(_p(dato, "firma"))

    cierre = Table([[columna_condiciones, columna_firma]], colWidths=[ANCHO_UTIL * 0.6, ANCHO_UTIL * 0.4])
    cierre.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (0, 0), 0),
                ("ALIGN", (1, 0), (1, 0), "CENTER"),
            ]
        )
    )
    historia.append(KeepTogether(cierre))

    texto_pie = limpiar_texto(f"{empresa.get('nombre', '')} - Presupuesto {numero}")

    class CanvasNumerado(rl_canvas.Canvas):
        """Dibuja el pie al final, cuando ya se conoce el total de paginas."""

        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self._paginas = []

        def showPage(self):
            self._paginas.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            total = len(self._paginas)
            for estado in self._paginas:
                self.__dict__.update(estado)
                self.saveState()
                self.setStrokeColor(LINEA)
                self.setLineWidth(0.5)
                self.line(MARGEN, 12 * mm, A4[0] - MARGEN, 12 * mm)
                self.setFont("Helvetica", 7.5)
                self.setFillColor(GRIS)
                self.drawString(MARGEN, 8 * mm, texto_pie)
                self.drawRightString(A4[0] - MARGEN, 8 * mm, f"Página {self._pageNumber} de {total}")
                self.restoreState()
                super().showPage()
            super().save()

    doc.build(historia, canvasmaker=CanvasNumerado)
    return buffer.getvalue()
