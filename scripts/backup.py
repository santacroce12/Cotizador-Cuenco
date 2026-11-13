"""Backup del cotizador: base SQLite + fotos de productos en un .zip con fecha.

- Copia la base con la API de backup de SQLite: es segura con la app andando
  (no copia un archivo a medio escribir, como pasaria con un cp).
- Verifica la copia con PRAGMA integrity_check antes de darla por buena.
- Borra los backups mas viejos que --dias (por defecto 30).

Uso (en Docker, desde la carpeta del proyecto en el servidor):
    docker compose exec -T cotizador python scripts/backup.py

Uso local:
    python scripts/backup.py --base data/database.db --uploads storage/uploads/productos --destino backups

Sale con codigo 1 si algo falla, para que cron o el monitoreo lo detecten.
"""

from __future__ import annotations

import argparse
import os
import sqlite3
import sys
import tempfile
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

PREFIJO = "cotizador-backup-"


def copiar_base(origen: Path, destino: Path) -> None:
    fuente = sqlite3.connect(f"file:{origen.as_posix()}?mode=ro", uri=True, timeout=30)
    try:
        copia = sqlite3.connect(destino)
        try:
            fuente.backup(copia)
            resultado = copia.execute("PRAGMA integrity_check").fetchone()[0]
            if resultado != "ok":
                raise RuntimeError(f"La copia de la base no paso integrity_check: {resultado}")
            cotizaciones = copia.execute("SELECT COUNT(*) FROM cotizacion").fetchone()[0]
            print(f"Base copiada y verificada ({cotizaciones} cotizaciones).")
        finally:
            copia.close()
    finally:
        fuente.close()


def armar_zip(base: Path, uploads: Path | None, destino_dir: Path) -> Path:
    destino_dir.mkdir(parents=True, exist_ok=True)
    marca = datetime.now().strftime("%Y%m%d-%H%M%S")
    archivo_zip = destino_dir / f"{PREFIJO}{marca}.zip"
    with tempfile.TemporaryDirectory() as tmp:
        copia_base = Path(tmp) / "database.db"
        copiar_base(base, copia_base)
        with zipfile.ZipFile(archivo_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.write(copia_base, "database.db")
            fotos = 0
            if uploads and uploads.is_dir():
                for archivo in uploads.rglob("*"):
                    if archivo.is_file() and archivo.name != ".gitkeep":
                        zf.write(archivo, Path("uploads/productos") / archivo.relative_to(uploads))
                        fotos += 1
            print(f"Fotos incluidas: {fotos}.")
    print(f"Backup creado: {archivo_zip} ({archivo_zip.stat().st_size / 1024 / 1024:.1f} MB)")
    return archivo_zip


def borrar_viejos(destino_dir: Path, dias: int) -> None:
    limite = datetime.now() - timedelta(days=dias)
    for archivo in destino_dir.glob(f"{PREFIJO}*.zip"):
        if datetime.fromtimestamp(archivo.stat().st_mtime) < limite:
            archivo.unlink()
            print(f"Borrado por antiguedad: {archivo.name}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--base", default=os.getenv("DATABASE_PATH", "/app/data/database.db"))
    parser.add_argument("--uploads", default=os.getenv("UPLOADS_PRODUCTOS_DIR", "/app/static/uploads/productos"))
    parser.add_argument("--destino", default=os.getenv("BACKUP_DIR", "/app/data/backups"))
    parser.add_argument("--dias", type=int, default=int(os.getenv("BACKUP_DIAS", "30")))
    args = parser.parse_args()

    base = Path(args.base)
    if not base.is_file():
        print(f"ERROR: no existe la base {base}", file=sys.stderr)
        return 1
    try:
        armar_zip(base, Path(args.uploads) if args.uploads else None, Path(args.destino))
        borrar_viejos(Path(args.destino), args.dias)
    except Exception as exc:  # noqa: BLE001 - cualquier falla debe salir con codigo 1
        print(f"ERROR: el backup fallo: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
