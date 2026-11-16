"""Prueba el redondeo del formulario (JavaScript) con Node, si esta instalado.

El formulario calcula en el navegador con las mismas reglas que el servidor.
Un error en el redondeo de JavaScript no se ve en los tests de Python.
"""

import json
import re
import shutil
import subprocess
import unittest
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]


def extraer_funcion(nombre):
    html = (RAIZ / "templates" / "cotizador.html").read_text(encoding="utf-8")
    inicio = html.index(f"function {nombre}(")
    fin = re.search(r"\n}\n", html[inicio:]).end()
    return html[inicio : inicio + fin]


@unittest.skipUnless(shutil.which("node"), "Node no esta instalado")
class RedondeoNavegadorTest(unittest.TestCase):
    def test_redondear_centavos(self):
        casos = {
            "0.1 + 0.2 - 0.3": 0,  # residuo 5.5e-17: antes daba NaN
            "2.9e-11": 0,
            "-2.9e-11": 0,
            "1.005": 1.01,  # mitad hacia arriba, como Decimal ROUND_HALF_UP
            "2.675": 2.68,
            "47.95 * 1000": 47950,
            "-8.255": -8.26,
            "193.116": 193.12,
            "NaN": 0,
        }
        codigo = extraer_funcion("redondearCentavos") + "\nconst r = {" + ",".join(
            f"{json.dumps(expr)}: redondearCentavos({expr})" for expr in casos
        ) + "};\nconsole.log(JSON.stringify(r));"
        salida = subprocess.run(["node", "-e", codigo], capture_output=True, text=True, timeout=30)
        self.assertEqual(salida.returncode, 0, salida.stderr)
        resultados = json.loads(salida.stdout)
        for expr, esperado in casos.items():
            self.assertEqual(resultados[expr], esperado, expr)


if __name__ == "__main__":
    unittest.main()
