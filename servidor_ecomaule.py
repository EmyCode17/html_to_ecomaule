
"""
ECOMAULE Analytics - Servidor web

Compatible con:
- Render (hosting)
- Ejecución local
- HTML de ECOMAULE Analytics
- Excel maestro con hojas REAL y PPTO

No requiere bibliotecas externas.
"""

from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from datetime import datetime
import json
import os
import posixpath
import re
import zipfile
from xml.etree import ElementTree as ET


# ==========================================================
# CONFIGURACIÓN
# ==========================================================

RAIZ = Path(__file__).resolve().parent

ARCHIVO = RAIZ / "ECOMAULE_PRUEBA_VINCULADA.xlsx"

HTML = "ECOMAULE_Analytics_VINCULADO.html"

PUERTO = int(os.environ.get("PORT", 10000))

N = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
P = "{http://schemas.openxmlformats.org/package/2006/relationships}"

UPS = (
    "LODOS",
    "COMPOST",
    "RELLENO",
    "BIOGAS",
    "VAL ENV",
    "RCD"
)

ESCENARIOS = ("REAL", "PPTO")

CUENTAS = {
    "INGRESOS",
    "COGS",
    "GAV",
    "COMBUSTIBLE",
    "ARRIENDO MAQUINARIA",
    "VIRUTA",
    "TON TRATADAS",
    "LITROS COMBUSTIBLE",
    "HORAS MAQUINA",
    "VOLUMEN VIRUTA",
    "ACOPIO FINAL",
    "TRATAMIENTO ACOPIO",
    "ACOPIO SMA",
    "HORAS EXCAVADORA",
    "VUELTAS TOLVA",
    "GENERACION BIOGAS",
    "TON BIODIGESTION",
    "RILES TRATADOS",
    "LODOS BIODIGESTION",
}


# ==========================================================
# LECTURA DE CELDAS EXCEL
# ==========================================================

def texto_celda(celda, compartidas):
    tipo = celda.attrib.get("t", "")

    inline = celda.find(N + "is")

    if inline is not None:
        return "".join(
            n.text or ""
            for n in inline.iter(N + "t")
        )

    v = celda.find(N + "v")

    if v is None or v.text is None:
        return None

    if tipo == "s":
        return compartidas[int(v.text)]

    if tipo in ("str", "inlineStr", "e"):
        return v.text

    if tipo == "b":
        return v.text == "1"

    try:
        numero = float(v.text)
        return int(numero) if numero.is_integer() else numero
    except ValueError:
        return v.text


# ==========================================================
# LECTURA DEL LIBRO EXCEL
# ==========================================================

def leer_libro(archivo):
    """
    Lee las hojas REAL y PPTO sin bibliotecas externas.
    """

    with zipfile.ZipFile(archivo) as z:

        compartidas = []

        if "xl/sharedStrings.xml" in z.namelist():
            raiz = ET.fromstring(
                z.read("xl/sharedStrings.xml")
            )

            compartidas = [
                "".join(
                    n.text or ""
                    for n in si.iter(N + "t")
                )
                for si in raiz.findall(N + "si")
            ]

        libro = ET.fromstring(
            z.read("xl/workbook.xml")
        )

        rels = ET.fromstring(
            z.read("xl/_rels/workbook.xml.rels")
        )

        destinos = {
            x.get("Id"): x.get("Target")
            for x in rels.findall(P + "Relationship")
        }

        resultados = {}

        for hoja in libro.findall(
            N + "sheets/" + N + "sheet"
        ):

            nombre = hoja.get("name")

            if nombre not in ESCENARIOS:
                continue

            target = destinos[
                hoja.get(R + "id")
            ]

            if target.startswith("/"):
                entry = target.lstrip("/")
            else:
                entry = posixpath.normpath(
                    "xl/" + target
                )

            xml = ET.fromstring(
                z.read(entry)
            )

            tabla = []

            for row in xml.findall(
                N + "sheetData/" + N + "row"
            ):

                columnas = {}

                for cell in row.findall(N + "c"):

                    referencia = cell.attrib.get(
                        "r", ""
                    )

                    letras = re.match(
                        r"[A-Z]+",
                        referencia
                    )

                    if not letras:
                        continue

                    indice = 0

                    for letra in letras.group():
                        indice = (
                            indice * 26
                            + ord(letra)
                            - 64
                        )

                    columnas[indice - 1] = texto_celda(
                        cell,
                        compartidas
                    )

                if columnas:
                    tabla.append([
                        columnas.get(i)
                        for i in range(
                            max(columnas) + 1
                        )
                    ])

            resultados[nombre] = tabla

        return resultados


# ==========================================================
# VALIDACIÓN Y CONSOLIDACIÓN
# ==========================================================

def convertir(archivo):

    hojas = leer_libro(archivo)

    if (
        "REAL" not in hojas
        or "PPTO" not in hojas
    ):
        raise ValueError(
            "El Excel debe contener las hojas REAL y PPTO."
        )

    registros = []
    avisos = []
    unicos = set()

    for escenario, tabla in hojas.items():

        if not tabla:
            continue

        encabezados = {
            str(valor).strip().upper(): indice
            for indice, valor in enumerate(tabla[0])
            if valor is not None
        }

        if (
            "PERIODO" not in encabezados
            or "INDICADOR" not in encabezados
        ):
            raise ValueError(
                f"Faltan encabezados PERIODO "
                f"e INDICADOR en {escenario}."
            )

        def obtener(fila, nombre):
            indice = encabezados.get(nombre)

            if (
                indice is not None
                and indice < len(fila)
            ):
                return fila[indice]

            return None

        for numero_fila, fila in enumerate(
            tabla[1:], 2
        ):

            periodo = str(
                obtener(fila, "PERIODO") or ""
            ).strip()

            indicador = str(
                obtener(fila, "INDICADOR") or ""
            ).strip().upper()

            if not periodo and not indicador:
                continue

            if not re.fullmatch(
                r"20\d{2}-(0[1-9]|1[0-2])",
                periodo
            ):
                avisos.append(
                    f"{escenario} fila {numero_fila}: "
                    "período no válido"
                )
                continue

            if indicador not in CUENTAS:
                avisos.append(
                    f"{escenario} fila {numero_fila}: "
                    f"indicador no reconocido: {indicador}"
                )
                continue

            for up in UPS:

                valor_original = obtener(
                    fila, up
                )

                if (
                    valor_original is None
                    or str(valor_original).strip() == ""
                ):
                    continue

                try:
                    valor = float(valor_original)
                except (ValueError, TypeError):
                    avisos.append(
                        f"{escenario} fila {numero_fila} "
                        f"{up}: valor no numérico"
                    )
                    continue

                clave = (
                    periodo,
                    up,
                    indicador,
                    escenario
                )

                if clave in unicos:
                    avisos.append(
                        f"{escenario} fila {numero_fila}: "
                        f"dato duplicado {clave}"
                    )
                    continue

                unicos.add(clave)

                registros.append({
                    "periodo": periodo,
                    "up": up,
                    "cuenta": indicador,
                    "escenario": escenario,
                    "valor": valor,
                    "origen": "EXCEL_VINCULADO"
                })

    if not registros:
        raise ValueError(
            "No se encontraron registros válidos "
            "en REAL y PPTO."
        )

    return {
        "registros": registros,
        "archivo": archivo.name,
        "actualizado_archivo": datetime.fromtimestamp(
            archivo.stat().st_mtime
        ).strftime("%d-%m-%Y %H:%M:%S"),
        "sincronizado": datetime.now().strftime(
            "%d-%m-%Y %H:%M:%S"
        ),
        "advertencias": avisos[:25],
        "total": len(registros)
    }


# ==========================================================
# SERVIDOR WEB
# ==========================================================

class Handler(SimpleHTTPRequestHandler):

    def __init__(self, *args, **kwargs):
        super().__init__(
            *args,
            directory=str(RAIZ),
            **kwargs
        )

    def end_headers(self):
        self.send_header(
            "Cache-Control",
            "no-store, no-cache, "
            "must-revalidate, max-age=0"
        )
        self.send_header(
            "X-Content-Type-Options",
            "nosniff"
        )
        super().end_headers()

    def list_directory(self, path):
        self.send_error(
            403,
            "No está permitido listar la carpeta"
        )
        return None

    def do_GET(self):

        ruta = self.path.split("?", 1)[0]

        # Página principal
        if ruta == "/":
            self.path = "/" + HTML
            return super().do_GET()

        # Comprobación de servidor
        if ruta == "/health":
            return self.enviar_json({
                "status": "ok",
                "servicio": "ECOMAULE Analytics"
            })

        # API de datos
        if ruta == "/api/maestro":

            if not ARCHIVO.exists():
                return self.enviar_json({
                    "error": (
                        "Fuente de datos no configurada. "
                        "El servidor todavía no está "
                        "conectado a SharePoint."
                    )
                }, 503)

            try:
                datos = convertir(ARCHIVO)
                return self.enviar_json(datos)

            except Exception as exc:
                return self.enviar_json({
                    "error": str(exc)
                }, 500)

        # Evitar descargar archivos internos
        if (
            ruta.endswith(".py")
            or ruta.endswith(".xlsx")
            or ruta.endswith(".xls")
            or ruta.startswith("/.")
        ):
            self.send_error(403)
            return

        return super().do_GET()

    def enviar_json(self, datos, estado=200):

        contenido = json.dumps(
            datos,
            ensure_ascii=False
        ).encode("utf-8")

        self.send_response(estado)

        self.send_header(
            "Content-Type",
            "application/json; charset=utf-8"
        )

        self.send_header(
            "Content-Length",
            str(len(contenido))
        )

        self.end_headers()
        self.wfile.write(contenido)


# ==========================================================
# INICIO DEL SERVIDOR
# ==========================================================

if __name__ == "__main__":

    print("=" * 60)
    print("ECOMAULE ANALYTICS")
    print("Servidor web Python")
    print(f"Puerto: {PUERTO}")
    print(f"Panel: {HTML}")
    print("=" * 60)

    servidor = ThreadingHTTPServer(
        ("0.0.0.0", PUERTO),
        Handler
    )

    try:
        servidor.serve_forever()

    except KeyboardInterrupt:
        print("Servidor detenido.")

    finally:
        servidor.server_close()

