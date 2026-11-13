"""Servidor local de demostración: vuelve a leer el Excel en cada consulta.
No requiere paquetes externos. Solo se comunica con este computador (127.0.0.1).
"""
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from datetime import datetime
import json
import os
import posixpath
import re
import threading
import webbrowser
import zipfile
from xml.etree import ElementTree as ET

RAIZ = Path(__file__).resolve().parent
ARCHIVO = RAIZ / "ECOMAULE_PRUEBA_VINCULADA.xlsx"
PUERTO = 8765
N = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
P = "{http://schemas.openxmlformats.org/package/2006/relationships}"
UPS = ("LODOS", "COMPOST", "RELLENO", "BIOGAS", "VAL ENV", "RCD")
ESCENARIOS = ("REAL", "PPTO")
CUENTAS = {
    "INGRESOS", "COGS", "GAV", "COMBUSTIBLE", "ARRIENDO MAQUINARIA", "VIRUTA",
    "TON TRATADAS", "LITROS COMBUSTIBLE", "HORAS MAQUINA", "VOLUMEN VIRUTA",
    "ACOPIO FINAL", "TRATAMIENTO ACOPIO", "ACOPIO SMA", "HORAS EXCAVADORA",
    "VUELTAS TOLVA", "GENERACION BIOGAS", "TON BIODIGESTION", "RILES TRATADOS",
    "LODOS BIODIGESTION",
}


def texto_celda(celda, compartidas):
    tipo = celda.attrib.get("t", "")
    inline = celda.find(N + "is")
    if inline is not None:
        return "".join(n.text or "" for n in inline.iter(N + "t"))
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
        n = float(v.text)
        return int(n) if n.is_integer() else n
    except ValueError:
        return v.text


def leer_libro(archivo):
    """Lee solo celdas ordinarias de las pestañas REAL y PPTO, sin librerías externas."""
    with zipfile.ZipFile(archivo) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            raiz = ET.fromstring(z.read("xl/sharedStrings.xml"))
            shared = ["".join(n.text or "" for n in si.iter(N + "t")) for si in raiz.findall(N + "si")]
        libro = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        destinations = {x.get("Id"): x.get("Target") for x in rels.findall(P + "Relationship")}
        resultados = {}
        for hoja in libro.findall(N + "sheets/" + N + "sheet"):
            nombre = hoja.get("name")
            if nombre not in ESCENARIOS:
                continue
            target = destinations[hoja.get(R + "id")]
            entry = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
            xml = ET.fromstring(z.read(entry))
            tabla = []
            for row in xml.findall(N + "sheetData/" + N + "row"):
                cols = {}
                for cell in row.findall(N + "c"):
                    ref = cell.attrib.get("r", "")
                    letras = re.match(r"[A-Z]+", ref)
                    if not letras:
                        continue
                    idx = 0
                    for letra in letras.group():
                        idx = idx * 26 + ord(letra) - 64
                    cols[idx - 1] = texto_celda(cell, shared)
                if cols:
                    tabla.append([cols.get(i) for i in range(max(cols) + 1)])
            resultados[nombre] = tabla
        return resultados


def convertir(archivo):
    hojas = leer_libro(archivo)
    if "REAL" not in hojas or "PPTO" not in hojas:
        raise ValueError("El Excel debe contener las hojas REAL y PPTO.")
    registros = []
    avisos = []
    unicos = set()
    for escenario, tabla in hojas.items():
        if not tabla:
            continue
        encabezados = {str(v).strip().upper(): i for i, v in enumerate(tabla[0]) if v is not None}
        if "PERIODO" not in encabezados or "INDICADOR" not in encabezados:
            raise ValueError(f"Faltan los encabezados periodo e indicador en {escenario}.")
        def obtener(fila, nombre):
            i = encabezados.get(nombre)
            return fila[i] if i is not None and i < len(fila) else None
        for n, fila in enumerate(tabla[1:], 2):
            periodo = str(obtener(fila, "PERIODO") or "").strip()
            indicador = str(obtener(fila, "INDICADOR") or "").strip().upper()
            if not periodo and not indicador:
                continue
            if not re.fullmatch(r"20\d{2}-(0[1-9]|1[0-2])", periodo):
                avisos.append(f"{escenario} fila {n}: período no válido")
                continue
            if indicador not in CUENTAS:
                avisos.append(f"{escenario} fila {n}: indicador no reconocido: {indicador}")
                continue
            for up in UPS:
                v = obtener(fila, up)
                if v is None or str(v).strip() == "":
                    continue
                try:
                    valor = float(v)
                except (ValueError, TypeError):
                    avisos.append(f"{escenario} fila {n} {up}: valor no numérico")
                    continue
                clave = (periodo, up, indicador, escenario)
                if clave in unicos:
                    avisos.append(f"{escenario} fila {n}: dato duplicado {clave}")
                    continue
                unicos.add(clave)
                registros.append({"periodo": periodo, "up": up, "cuenta": indicador,
                                  "escenario": escenario, "valor": valor, "origen": "EXCEL_VINCULADO"})
    if not registros:
        raise ValueError("No se encontraron registros válidos en REAL y PPTO.")
    return {"registros": registros,
            "archivo": archivo.name,
            "actualizado_archivo": datetime.fromtimestamp(archivo.stat().st_mtime).strftime("%d-%m-%Y %H:%M:%S"),
            "sincronizado": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
            "advertencias": avisos[:25], "total": len(registros)}


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=str(RAIZ), **kw)

    def end_headers(self):
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        super().end_headers()

    def list_directory(self, path):
        self.send_error(403, "No está permitido listar la carpeta")
        return None

    def do_GET(self):
        if self.path.split("?", 1)[0] == "/api/maestro":
            try:
                data = convertir(ARCHIVO)
                status = 200
            except Exception as exc:
                data = {"error": str(exc)}
                status = 500
            blob = json.dumps(data, ensure_ascii=False).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(blob)))
            self.end_headers()
            self.wfile.write(blob)
        else:
            super().do_GET()


if __name__ == "__main__":
    if not ARCHIVO.exists():
        raise SystemExit(f"Falta el Excel junto al servidor: {ARCHIVO}")
    print("=" * 66)
    print("ECOMAULE - Panel vinculado al Excel local")
    print("Mantén esta ventana abierta mientras utilizas el panel.")
    print(f"Archivo: {ARCHIVO.name}")
    print(f"Panel: http://127.0.0.1:{PUERTO}/ECOMAULE_Analytics_VINCULADO.html")
    print("Para detener: Ctrl+C")
    print("=" * 66)
    try:
        httpd = ThreadingHTTPServer(("127.0.0.1", PUERTO), Handler)
    except OSError as ex:
        raise SystemExit(f"No se pudo iniciar el puerto {PUERTO}. ¿Ya hay otra instancia abierta? {ex}")
    threading.Timer(0.7, lambda: webbrowser.open(f"http://127.0.0.1:{PUERTO}/ECOMAULE_Analytics_VINCULADO.html")).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("Servidor detenido.")
    finally:
        httpd.server_close()
