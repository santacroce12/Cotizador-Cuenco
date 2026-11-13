import sqlite3
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from calculos_cotizacion import calcular_totales, precio_unitario_neto  # noqa: E402


def linea(unitario, cantidad=1, iva=21, carga=0):
    return {"unitario_neto": unitario, "cantidad": cantidad, "iva_pct": iva, "carga_pct": carga}


class CalculosTest(unittest.TestCase):
    def test_unitario_se_redondea_y_las_lineas_suman_el_total(self):
        unitario = precio_unitario_neto(33.333, 5, 37, 0)
        self.assertEqual(unitario, 47.95)
        totales = calcular_totales([linea(unitario, 1000)], precio_incluye_iva=False)
        # Antes: la linea decia 47.950,00 y el subtotal 47.949,52.
        self.assertEqual(totales["lineas"][0]["neto"], 47950.00)
        self.assertEqual(totales["neto"], 47950.00)

    def test_redondeo_mitad_hacia_arriba(self):
        self.assertEqual(precio_unitario_neto(1.005, 0, 0, 0), 1.01)
        self.assertEqual(precio_unitario_neto(2.675, 0, 0, 0), 2.68)

    def test_precio_con_iva_incluido_cantidad_por_precio_final(self):
        totales = calcular_totales([linea(159.60, 2)], precio_incluye_iva=True)
        self.assertEqual(totales["lineas"][0]["unitario_final"], 193.12)
        self.assertEqual(totales["lineas"][0]["total"], 386.24)
        self.assertEqual(totales["total"], 386.24)
        self.assertAlmostEqual(totales["neto"] + totales["iva_total"], totales["total"], places=2)

    def test_responsable_inscrito_iva_por_alicuota(self):
        totales = calcular_totales([linea(150), linea(60, 2, iva=10.5)], precio_incluye_iva=False)
        self.assertEqual(totales["neto"], 270.0)
        self.assertEqual(totales["iva_por_tasa"], {21.0: 31.5, 10.5: 12.6})
        self.assertEqual(totales["total"], 314.1)

    def test_bonificacion_se_reparte_entre_neto_e_iva(self):
        totales = calcular_totales([linea(159.60, carga=5)], precio_incluye_iva=False, bonificacion_solicitada=10)
        self.assertEqual(totales["bonificacion"], 10.0)
        self.assertEqual(totales["total"], 183.12)
        self.assertEqual(totales["bonificacion_neta"], 8.26)
        self.assertEqual(totales["neto_bonificado"], 151.34)
        self.assertEqual(totales["iva_total"], 31.78)
        # El IVA informado es el 21% del neto que realmente se cobra.
        self.assertAlmostEqual(totales["neto_bonificado"] * 0.21, totales["iva_total"], delta=0.01)
        self.assertAlmostEqual(totales["neto_bonificado"] + totales["iva_total"], totales["total"], places=2)
        self.assertEqual(totales["carga_fiscal"], 7.57)

    def test_bonificacion_no_supera_el_total_y_no_es_negativa(self):
        self.assertEqual(calcular_totales([linea(100)], False, 999999)["total"], 0.0)
        self.assertEqual(calcular_totales([linea(100)], False, -50)["bonificacion"], 0.0)

    def test_imprevisto_general_se_aplica_sobre_el_costo_con_imprevisto_del_item(self):
        # 100 de costo, 5% imprevisto del item, 10% general, 50% margen:
        # 100 x 1,05 x 1,10 x 1,50 = 173,25
        self.assertEqual(precio_unitario_neto(100, 5, 50, 0, 10), 173.25)
        # Sin imprevisto general da lo mismo que antes.
        self.assertEqual(precio_unitario_neto(100, 5, 50, 0), precio_unitario_neto(100, 5, 50, 0, 0))

    def test_sin_lineas(self):
        totales = calcular_totales([], precio_incluye_iva=True, bonificacion_solicitada=10)
        self.assertEqual((totales["neto"], totales["total"], totales["bonificacion"]), (0.0, 0.0, 0.0))


class BackupTest(unittest.TestCase):
    def test_backup_genera_zip_verificado_con_base_y_fotos(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            base = tmp / "database.db"
            conexion = sqlite3.connect(base)
            conexion.execute("PRAGMA journal_mode=WAL")
            conexion.execute("CREATE TABLE cotizacion (id INTEGER PRIMARY KEY)")
            conexion.execute("INSERT INTO cotizacion DEFAULT VALUES")
            conexion.commit()
            uploads = tmp / "uploads"
            uploads.mkdir()
            (uploads / "foto.jpg").write_bytes(b"jpg")
            # La app sigue con la conexion abierta mientras corre el backup.
            resultado = subprocess.run(
                [sys.executable, str(RAIZ / "scripts" / "backup.py"), "--base", str(base),
                 "--uploads", str(uploads), "--destino", str(tmp / "backups")],
                capture_output=True, text=True,
            )
            conexion.close()
            self.assertEqual(resultado.returncode, 0, resultado.stderr)
            zips = list((tmp / "backups").glob("cotizador-backup-*.zip"))
            self.assertEqual(len(zips), 1)
            with zipfile.ZipFile(zips[0]) as zf:
                self.assertEqual(sorted(zf.namelist()), ["database.db", "uploads/productos/foto.jpg"])
                zf.extract("database.db", tmp / "restaurado")
            copia = sqlite3.connect(tmp / "restaurado" / "database.db")
            self.assertEqual(copia.execute("SELECT COUNT(*) FROM cotizacion").fetchone()[0], 1)
            copia.close()

    def test_backup_falla_con_codigo_1_si_no_hay_base(self):
        resultado = subprocess.run(
            [sys.executable, str(RAIZ / "scripts" / "backup.py"), "--base", "no-existe.db", "--destino", "x"],
            capture_output=True, text=True,
        )
        self.assertEqual(resultado.returncode, 1)


if __name__ == "__main__":
    unittest.main()
