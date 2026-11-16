# Guia para usar `cotizador.local` solo en tu PC

## Objetivo

Hacer que `http://cotizador.local` abra el cotizador **solo en tu PC**, sin tocar el DNS global del router ni afectar otros servicios de la red.

## Cuando conviene este metodo

Usalo si:

- no queres cambiar el router
- no queres que toda la red dependa del Ubuntu
- solo necesitas que funcione en tu PC de trabajo

No usalo si queres que tambien funcione en todos los telefonos y PCs de la red. Para eso conviene DNS local en router o DHCP.

## Datos actuales del sistema

- servidor Ubuntu: `192.168.0.50`
- dominio interno deseado: `cotizador.local`
- cotizador: publicado en el server y accesible por red local

## Paso a paso en Windows

### 1. Abrir el Bloc de notas como administrador

- apreta `Inicio`
- busca `Bloc de notas`
- click derecho
- `Ejecutar como administrador`

### 2. Abrir el archivo `hosts`

Desde el Bloc de notas:

- `Archivo`
- `Abrir`
- pega esta ruta:

```text
C:\Windows\System32\drivers\etc\hosts
```

Si no lo ves:

- en el selector de archivos cambia `Documentos de texto (*.txt)` por `Todos los archivos (*.*)`

### 3. Agregar esta linea al final

```text
192.168.0.50 cotizador.local
```

Guarda el archivo.

### 4. Limpiar cache DNS

Abri `PowerShell` o `CMD` como administrador y corre:

```powershell
ipconfig /flushdns
```

### 5. Verificar que resolvio bien

Corre:

```powershell
ping cotizador.local
```

Esperado:

```text
Haciendo ping a cotizador.local [192.168.0.50]
```

### 6. Abrir el sistema

Proba primero:

```text
http://cotizador.local
```

Si no abre, proba estas dos URL de control:

```text
http://192.168.0.50
http://192.168.0.50:9000
```

## Como deshacerlo

Si en algun momento queres volver atras:

1. volve a abrir el archivo `hosts`
2. borra esta linea:

```text
192.168.0.50 cotizador.local
```

3. limpia cache otra vez:

```powershell
ipconfig /flushdns
```

## Que NO cambia este metodo

Este metodo:

- no toca el router
- no cambia el DNS de toda la red
- no afecta el otro servicio publicado
- no afecta telefonos ni otras PCs

## Si despues queres que funcione tambien en el telefono

Tenes dos caminos:

1. agregar DNS manual solo en el telefono
2. configurar DNS local en el router

Mientras tanto, este metodo es el mas seguro para no mover toda la red.
