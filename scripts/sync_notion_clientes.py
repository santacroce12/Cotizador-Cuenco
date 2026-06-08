import argparse
import os
import sys
import time
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from app import Cliente, app  # noqa: E402
from notion_service import notion_enabled, sync_cliente_to_notion  # noqa: E402


def parse_ids(raw_ids):
    if not raw_ids:
        return []

    ids = []
    for raw_id in str(raw_ids).split(","):
        raw_id = raw_id.strip()
        if not raw_id:
            continue
        if not raw_id.isdigit():
            raise ValueError(f"ID invalido: {raw_id}")
        ids.append(int(raw_id))
    return ids


def build_parser():
    parser = argparse.ArgumentParser(
        description="Sincroniza clientes existentes del cotizador hacia Notion."
    )
    parser.add_argument(
        "--execute",
        action="store_true",
        help="Ejecuta la sincronizacion real. Sin este flag solo muestra un dry-run.",
    )
    parser.add_argument(
        "--ids",
        help="Lista de IDs de cliente separados por coma. Si se omite, toma todos.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        help="Cantidad maxima de clientes a procesar.",
    )
    parser.add_argument(
        "--sleep",
        type=float,
        default=0.15,
        help="Pausa en segundos entre clientes para evitar rate limits.",
    )
    parser.add_argument(
        "--fail-fast",
        action="store_true",
        help="Corta ante el primer cliente que no pueda sincronizarse.",
    )
    return parser


def main():
    args = build_parser().parse_args()

    try:
        ids = parse_ids(args.ids)
    except ValueError as exc:
        print(f"[notion-backfill] {exc}")
        return 2

    if args.limit is not None and args.limit <= 0:
        print("[notion-backfill] --limit debe ser mayor a 0")
        return 2

    with app.app_context():
        query = Cliente.query.order_by(Cliente.id.asc())
        if ids:
            query = query.filter(Cliente.id.in_(ids))
        if args.limit:
            query = query.limit(args.limit)

        clientes = query.all()

        print(f"[notion-backfill] DATABASE={app.config.get('SQLALCHEMY_DATABASE_URI')}")
        print(f"[notion-backfill] NOTION_ENABLED={notion_enabled()}")
        print(f"[notion-backfill] CLIENTES_ENCONTRADOS={len(clientes)}")

        if not clientes:
            return 0

        if not args.execute:
            print("[notion-backfill] Dry-run. No se enviara nada a Notion.")
            for cliente in clientes[:20]:
                print(f"[notion-backfill] cliente id={cliente.id} nombre={cliente.nombre}")
            if len(clientes) > 20:
                print(f"[notion-backfill] ... y {len(clientes) - 20} clientes mas")
            print("[notion-backfill] Para ejecutar: python scripts/sync_notion_clientes.py --execute")
            return 0

        if not notion_enabled():
            print("[notion-backfill] Abortado: NOTION_ENABLED debe ser true para ejecutar.")
            return 2

        if not os.getenv("NOTION_TOKEN", "").strip():
            print("[notion-backfill] Abortado: NOTION_TOKEN no esta configurado.")
            return 2

        ok = 0
        errores = 0
        for cliente in clientes:
            page_id = sync_cliente_to_notion(cliente)
            if page_id:
                ok += 1
                print(f"[notion-backfill] OK cliente id={cliente.id} notion_page={page_id}")
            else:
                errores += 1
                print(f"[notion-backfill] ERROR cliente id={cliente.id} nombre={cliente.nombre}")
                if args.fail_fast:
                    break
            if args.sleep > 0:
                time.sleep(args.sleep)

        print(f"[notion-backfill] FINAL ok={ok} errores={errores}")
        return 0 if errores == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
