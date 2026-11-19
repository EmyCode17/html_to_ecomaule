"""ECOMAULE: panel web y recepción de PanelWebDAN desde Power Automate.
Sin dependencias externas. BIOGAS_SYNC_KEY debe configurarse en Render.
El Excel de SharePoint sigue siendo la fuente maestra. Cada POST reemplaza
la fotografía completa; tras perder el disco temporal debe reenviarse la tabla.
"""
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from datetime import datetime, timedelta, timezone
from decimal import Decimal, InvalidOperation
from zoneinfo import ZoneInfo
from collections import defaultdict
import hmac
import json
import os
import tempfile
import threading
from urllib.parse import urlsplit

RAIZ = Path(__file__).resolve().parent
HTML = RAIZ / 'ECOMAULE_Analytics_VINCULADO.html'
DATOS = Path(os.environ.get('BIOGAS_DATA_FILE', str(Path(tempfile.gettempdir()) / 'ecomaule-biogas.json')))
LOCK = threading.Lock()
MAX_BODY = 5 * 1024 * 1024


def fecha(valor):
    texto = str(valor).strip()
    try:
        serial = Decimal(texto)
        if not serial.is_finite() or serial < 1 or serial > 80000:
            raise ValueError('Fecha serial fuera de rango')
        return (datetime(1899, 12, 30) + timedelta(days=float(serial))).date()
    except InvalidOperation:
        pass
    for formato in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y'):
        try:
            return datetime.strptime(texto[:10], formato).date()
        except ValueError:
            pass
    raise ValueError('Fecha no válida')


CAMPOS = {
    'Biogas Generado M3': 'GENERACION BIOGAS',
    'Toneladas tratadas': 'TON BIODIGESTION',
    'RILES tratados': 'RILES TRATADOS',
    'Lodos tratados': 'LODOS BIODIGESTION',
}


def hoy_chile():
    try:
        return datetime.now(ZoneInfo("America/Santiago")).date()
    except Exception as exc:
        raise ValueError("No se pudo obtener la zona America/Santiago; instala tzdata") from exc


def convertir(payload, corte=None):
    tabla = payload if isinstance(payload, list) else payload.get('value') if isinstance(payload, dict) else None
    if not isinstance(tabla, list) or not tabla or len(tabla) > 20000:
        raise ValueError('Se espera la lista completa de filas de PanelWebDAN (1 a 20000 filas)')
    corte = corte or hoy_chile()
    meses = defaultdict(Decimal)
    biogas_dias = defaultdict(set)
    futuras = 0
    dias = set()
    omitidas = 0
    presentes = set()
    for numero, fila in enumerate(tabla, 2):
        if not isinstance(fila, dict):
            raise ValueError(f'Fila {numero}: debe ser un objeto')
        valores = {nombre: fila.get(nombre) for nombre in CAMPOS
                   if fila.get(nombre) is not None and str(fila.get(nombre)).strip() != ''}
        if not valores:
            omitidas += 1
            continue
        try:
            dia = fecha(fila.get('Fecha', ''))
            if dia > corte:
                futuras += 1
                continue
            if dia in dias:
                raise ValueError(f'Fecha duplicada: {dia.isoformat()}')
            for nombre, raw in valores.items():
                texto = str(raw).strip().replace(' ', '').replace('\u00a0', '')
                if ',' in texto:
                    texto = texto.replace('.', '').replace(',', '.')
                try:
                    cantidad = Decimal(texto)
                except InvalidOperation as exc:
                    raise ValueError(f'{nombre}: valor no numérico') from exc
                if not cantidad.is_finite() or cantidad < 0 or cantidad > Decimal('1e12'):
                    raise ValueError(f'{nombre}: debe ser un número finito entre 0 y 1e12')
                meses[(dia.strftime('%Y-%m'), CAMPOS[nombre])] += cantidad
                presentes.add(nombre)
                if nombre == 'Biogas Generado M3':
                    biogas_dias[dia.strftime('%Y-%m')].add(dia)
        except (ValueError, InvalidOperation, OverflowError) as exc:
            raise ValueError(f'Fila {numero}: {exc}') from exc
        dias.add(dia)
    if not dias:
        raise ValueError('No hay filas con fecha e indicadores válidos; se conserva la carga anterior')
    ahora = datetime.now(timezone.utc).isoformat(timespec='seconds')
    avisos = [f'{omitidas} filas sin indicadores omitidas'] if omitidas else []
    avisos += [f'Sin valores recibidos: {nombre}' for nombre in CAMPOS if nombre not in presentes]
    resumen_promedio = {}
    for mes, fechas in sorted(biogas_dias.items()):
        fin = max(fechas)
        dias_periodo = fin.day
        faltantes = [f'{mes}-{d:02d}' for d in range(1, dias_periodo + 1)
                     if datetime(fin.year, fin.month, d).date() not in fechas]
        total = meses[(mes, 'GENERACION BIOGAS')]
        promedio = float(total / Decimal(dias_periodo)) if not faltantes else None
        resumen_promedio[mes] = {
            'corte': fin.isoformat(), 'dias': dias_periodo,
            'dias_con_datos': len(fechas), 'faltantes': faltantes,
            'total_m3': float(total), 'promedio_m3_dia': promedio,
        }
        if promedio is not None:
            meses[(mes, 'BIOGAS PROMEDIO DIARIO')] = Decimal(str(promedio))
        else:
            avisos.append(f'{mes}: promedio no disponible; faltan {len(faltantes)} días de biogás')
    if futuras:
        avisos.append(f'{futuras} filas futuras excluidas; corte máximo {corte.isoformat()}')
    return {
        'biogas_promedio_mensual': resumen_promedio,
        'fecha_corte_maxima': corte.isoformat(),
        '_filas_fuente': tabla,
        'registros': [{'periodo': mes, 'up': 'BIOGAS', 'cuenta': cuenta,
                       'escenario': 'REAL', 'valor': float(total), 'origen': 'SHAREPOINT_POWER_AUTOMATE'}
                      for (mes, cuenta), total in sorted(meses.items())],
        'archivo': 'Control de Operación DAN Ecomaule.xlsx · PanelWebDAN',
        'sincronizado': ahora, 'actualizado_archivo': 'No informado por Excel',
        'total': len(meses), 'meses': len({mes for mes, cuenta in meses}),
        'indicadores_recibidos': sorted(presentes), 'filas_diarias': len(dias),
        'advertencias': avisos, 'ultima_fecha': max(dias).isoformat(),
    }


def guardar(datos):
    DATOS.parent.mkdir(parents=True, exist_ok=True)
    with LOCK:
        fd, nombre = tempfile.mkstemp(dir=DATOS.parent, prefix='biogas-', suffix='.tmp')
        try:
            with os.fdopen(fd, 'w', encoding='utf-8') as archivo:
                json.dump(datos, archivo, ensure_ascii=False, allow_nan=False)
                archivo.flush()
                os.fsync(archivo.fileno())
            os.replace(nombre, DATOS)
        finally:
            if os.path.exists(nombre):
                os.unlink(nombre)


class Handler(BaseHTTPRequestHandler):
    protocol_version = 'HTTP/1.0'

    def enviar(self, contenido, estado=200, tipo='application/json; charset=utf-8'):
        blob = contenido if isinstance(contenido, bytes) else json.dumps(contenido, ensure_ascii=False, allow_nan=False).encode('utf-8')
        self.send_response(estado)
        self.send_header('Content-Type', tipo)
        self.send_header('Content-Length', str(len(blob)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.end_headers()
        self.wfile.write(blob)

    def do_GET(self):
        ruta = urlsplit(self.path).path
        if ruta == '/health':
            return self.enviar({'status': 'ok', 'receptor_configurado': bool(os.environ.get('BIOGAS_SYNC_KEY')), 'datos_disponibles': DATOS.is_file()})
        if ruta == '/api/maestro':
            try:
                with LOCK:
                    datos = json.loads(DATOS.read_text(encoding='utf-8'))
                if '_filas_fuente' in datos:
                    sincronizado = datos['sincronizado']
                    datos = convertir(datos['_filas_fuente'])
                    datos['sincronizado'] = sincronizado
                datos.pop('_filas_fuente', None)
                return self.enviar(datos)
            except FileNotFoundError:
                return self.enviar({'error': 'Esperando el primer envío de Power Automate', 'registros': []}, 503)
            except (OSError, ValueError):
                return self.enviar({'error': 'No se pudo leer la última carga; vuelve a ejecutar el flujo'}, 503)
        if ruta in ('/', '/' + HTML.name):
            try:
                return self.enviar(HTML.read_bytes(), tipo='text/html; charset=utf-8')
            except OSError:
                return self.enviar({'error': 'Falta el archivo HTML del panel'}, 503)
        return self.enviar({'error': 'Ruta no disponible'}, 404)

    def do_POST(self):
        if urlsplit(self.path).path != '/api/biogas':
            return self.enviar({'error': 'Ruta no disponible'}, 404)
        clave = os.environ.get('BIOGAS_SYNC_KEY', '')
        if len(clave) < 32:
            return self.enviar({'error': 'Configura BIOGAS_SYNC_KEY con al menos 32 caracteres en Render'}, 503)
        recibido = self.headers.get('X-ECOMAULE-Key', '')
        if not hmac.compare_digest(recibido.encode('utf-8'), clave.encode('utf-8')):
            return self.enviar({'error': 'Clave de envío incorrecta'}, 401)
        if self.headers.get('Content-Type', '').split(';')[0].strip().lower() != 'application/json':
            return self.enviar({'error': 'Se requiere Content-Type: application/json'}, 415)
        try:
            tamano = int(self.headers.get('Content-Length', '0'))
            if not 0 < tamano <= MAX_BODY:
                return self.enviar({'error': 'Tamaño del envío no permitido'}, 413)
            self.connection.settimeout(30)
            payload = json.loads(self.rfile.read(tamano))
            datos = convertir(payload)
        except (ValueError, UnicodeError, TimeoutError, OSError) as exc:
            return self.enviar({'error': str(exc)}, 400)
        try:
            guardar(datos)
        except OSError:
            return self.enviar({'error': 'No se pudo guardar la carga; se debe reintentar'}, 500)
        return self.enviar({'ok': True, 'meses': datos['meses'], 'registros': datos['total'], 'indicadores': datos['indicadores_recibidos'], 'filas_diarias': datos['filas_diarias'], 'sincronizado': datos['sincronizado']})


if __name__ == '__main__':
    puerto = int(os.environ.get('PORT', '10000'))
    print(f'ECOMAULE listo en puerto {puerto}', flush=True)
    with ThreadingHTTPServer(('0.0.0.0', puerto), Handler) as servidor:
        servidor.serve_forever()
