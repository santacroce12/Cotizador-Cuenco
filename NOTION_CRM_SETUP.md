# CRM en Notion

## Objetivo

Separar correctamente:

- `Lead`
- `Cliente`
- `Presupuesto`
- `Trabajo / Proyecto`

El cotizador no debe usarse como CRM general. El cotizador arranca cuando ya existe un interes real y vale la pena presupuestar.

## Bases en Notion

### `Clientes`

Base maestra de clientes formales.

Va aca:

- clientes creados desde el cotizador
- clientes ya calificados
- clientes con los que ya se opera de forma comercial o tecnica

No va aca:

- leads frios
- listados sin validar de Excel o Drive

### `Seguimiento Comercial`

Es el CRM diario. Cada fila es una oportunidad o gestion comercial.

Puede nacer de dos formas:

1. Manual:
   un lead cargado directamente en Notion desde Excel, Drive o relevamiento comercial
2. Desde cotizacion:
   una oportunidad creada o actualizada automaticamente cuando se genera una cotizacion en el sistema

### `Presupuestos`

Base de cotizaciones reales creadas en el cotizador.

No representa todos los leads, solo los casos que llegaron a presupuesto.

### `Trabajos / Proyectos`

Base operativa. Solo se crea cuando una cotizacion queda `Aceptada`.

## Reglas del sistema

### Cliente creado en cotizador

- crea o actualiza `Cliente` en Notion

### Cotizacion creada o editada en cotizador

- crea o actualiza `Presupuesto`
- crea o actualiza `Seguimiento Comercial`
- no crea `Trabajo / Proyecto`

### Cotizacion aceptada

- actualiza `Presupuesto` a aprobado
- actualiza `Seguimiento Comercial` a ganado
- crea o actualiza `Trabajo / Proyecto`

### Cotizacion rechazada

- actualiza `Presupuesto` a rechazado
- actualiza `Seguimiento Comercial` a perdido
- no crea `Trabajo / Proyecto`

### Lead desde Excel o Drive

- se carga directamente en `Seguimiento Comercial`
- no crea `Cliente`
- no crea `Presupuesto`
- no crea `Trabajo / Proyecto`

## Flujo recomendado

1. El comercial carga o pega leads manualmente en `Seguimiento Comercial`.
2. Trabaja desde ahi las etapas: `Sin contactar`, `Contactado`, `Interesado`, `A cotizar`, `Ganado`, `Perdido`.
3. Cuando el lead pide presupuesto, recien ahi se crea el cliente en el cotizador.
4. Se arma la cotizacion en el cotizador.
5. El sistema sincroniza `Clientes`, `Presupuestos` y la oportunidad ligada a esa cotizacion.
6. Si la cotizacion se acepta, se crea `Trabajo / Proyecto`.

## Campos recomendados para `Seguimiento Comercial`

Minimos:

- `Seguimiento` (title)
- `Cliente`
- `Presupuesto`
- `Estado`
- `Etapa comercial`
- `Tipo de oportunidad`
- `Responsable`
- `Canal`
- `Proximo seguimiento`
- `Proxima accion`
- `Probabilidad`
- `Monto estimado`
- `Moneda`
- `Resumen comercial`
- `Origen`

Utiles para leads manuales:

- `Empresa / Lead`
- `Contacto`
- `Telefono`
- `Email`
- `Origen lead`
- `Ultimo contacto`
- `Observaciones`

Utiles para oportunidades creadas desde cotizacion:

- `Estado cotizador`
- `Numero cotizacion`
- `Link cotizador`
- `Ultimo aviso cotizador`
- `Avisos recibidos`

## Criterio operativo

- Un `lead` no es un `cliente`.
- Una `oportunidad` no es un `presupuesto`.
- Un `presupuesto` no es un `trabajo`.
- Un `trabajo` solo nace cuando la cotizacion queda aceptada.

## Estado actual del codigo

El codigo ya hace esto:

- sincroniza clientes desde el cotizador a `Clientes`
- sincroniza cotizaciones a `Presupuestos`
- sincroniza oportunidades de cotizacion a `Seguimiento Comercial`
- solo crea `Trabajos / Proyectos` cuando la cotizacion esta aceptada

Todavia no hace esto:

- importar leads desde `.xlsx`
- crear leads manuales desde la app
- enlazar automaticamente una oportunidad manual de Notion con una cotizacion futura

Eso ultimo conviene resolverlo despues de estabilizar el uso diario del CRM.
