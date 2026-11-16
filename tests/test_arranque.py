"""El arranque tiene que aguantar varios procesos a la vez (workers de gunicorn)."""

import sqlite3
import subprocess
import sys
import tempfile
import textwrap
import unittest
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]

ARRANQUE = textwrap.dedent(
    """
    import os, sys
    base = sys.argv[1]
    os.environ["DATABASE_PATH"] = base
    os.environ["DATA_DIR"] = os.path.dirname(base)
    os.environ["UPLOADS_PRODUCTOS_DIR"] = os.path.join(os.path.dirname(base), "uploads")
    os.environ["APP_SECRET_KEY"] = "clave-secreta-solo-para-tests"
    sys.path.insert(0, sys.argv[2])
    import app  # corre las migraciones del arranque
    print("OK")
    """
)


class ArranqueConcurrenteTest(unittest.TestCase):
    def test_tres_procesos_arrancan_a_la_vez_sobre_una_base_nueva(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp) / "database.db"
            script = Path(tmp) / "arrancar.py"
            script.write_text(ARRANQUE, encoding="utf-8")
            procesos = [
                subprocess.Popen(
                    [sys.executable, str(script), str(base), str(RAIZ)],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                )
                for _ in range(3)
            ]
            resultados = [p.communicate(timeout=120) for p in procesos]
            for proceso, (salida, error) in zip(procesos, resultados):
                self.assertEqual(proceso.returncode, 0, error[-800:])
                self.assertIn("OK", salida)

            conexion = sqlite3.connect(base)
            self.assertEqual(conexion.execute("PRAGMA journal_mode").fetchone()[0], "wal")
            columnas = {fila[1] for fila in conexion.execute("PRAGMA table_info(usuario)")}
            self.assertTrue({"activo", "cargo", "email", "telefono"} <= columnas)
            conexion.close()


class SinInternetTest(unittest.TestCase):
    def test_las_plantillas_no_cargan_nada_de_internet(self):
        usadas = ["layout.html", "login.html", "dashboard.html", "cotizador.html", "historial.html",
                  "usuarios.html", "familias.html", "auditoria.html", "mi_cuenta.html",
                  "cotizacion_cliente.html", "cotizacion_llave_en_mano.html", "_dashboard_operativo.html"]
        for nombre in usadas:
            html = (RAIZ / "templates" / nombre).read_text(encoding="utf-8")
            for externo in ("cdn.jsdelivr.net", "fonts.googleapis.com", "fonts.gstatic.com", "unpkg.com", "cdnjs"):
                self.assertNotIn(externo, html, f"{nombre} carga {externo}")

    def test_las_librerias_locales_existen(self):
        vendor = RAIZ / "static" / "vendor"
        for archivo in (
            "bootstrap/bootstrap.min.css",
            "bootstrap/bootstrap.bundle.min.js",
            "bootstrap-icons/bootstrap-icons.min.css",
            "bootstrap-icons/fonts/bootstrap-icons.woff2",
            "tom-select/tom-select.bootstrap5.min.css",
            "tom-select/tom-select.complete.min.js",
            "chartjs/chart.umd.min.js",
            "inter/inter.css",
            "inter/inter-latin-wght-normal.woff2",
        ):
            ruta = vendor / archivo
            self.assertTrue(ruta.is_file(), archivo)
            self.assertGreater(ruta.stat().st_size, 500, archivo)


if __name__ == "__main__":
    unittest.main()
