import os
import io
import time
import logging
import threading
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional
from http.server import HTTPServer, BaseHTTPRequestHandler

# Slack Bolt SDK
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler
from slack_sdk.errors import SlackApiError

# Google Sheets SDK
import gspread
from google.oauth2.service_account import Credentials
from google.auth.exceptions import GoogleAuthError

# ReportLab PDF Engine
from reportlab.lib.pagesizes import letter
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfgen import canvas


# ==============================================================================
# 1. CONFIGURACIÓN Y AUDITORÍA (LOGGING)
# ==============================================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] [%(filename)s:%(lineno)d] - %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger("DragadosBot")

class Config:
    SLACK_BOT_TOKEN = os.environ.get("SLACK_BOT_TOKEN")
    SLACK_SIGNING_SECRET = os.environ.get("SLACK_SIGNING_SECRET")
    SLACK_APP_TOKEN = os.environ.get("SLACK_APP_TOKEN")
    PORT = int(os.environ.get("PORT", 10000))
    CREDENTIALS_FILE = os.environ.get("GOOGLE_CREDENTIALS_FILE", "credentials.json")
    SHEET_NAME = "Mantencion-Componentes-Dragados"
    CACHE_TTL_SECONDS = 300  # 5 minutos de caché en memoria

    @classmethod
    def validar(cls):
        faltantes = []
        if not cls.SLACK_BOT_TOKEN: faltantes.append("SLACK_BOT_TOKEN")
        if not cls.SLACK_SIGNING_SECRET: faltantes.append("SLACK_SIGNING_SECRET")
        if not cls.SLACK_APP_TOKEN: faltantes.append("SLACK_APP_TOKEN")
        
        if faltantes:
            logger.warning(f"Variables de entorno faltantes: {', '.join(faltantes)}")

Config.validar()


# ==============================================================================
# 2. SERVIDOR DE MONITOREO Y SALUD (HEALTH CHECK)
# ==============================================================================
class HealthCheckHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path in ["/", "/health", "/healthz"]:
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            response = '{"status": "ok", "service": "Dragados Slack Bot", "uptime": "active"}'
            self.wfile.write(response.encode("utf-8"))
        else:
            self.send_response(404)
            self.end_headers()

    def log_message(self, format, *args):
        return

def iniciar_servidor_healthcheck():
    server = HTTPServer(("0.0.0.0", Config.PORT), HealthCheckHandler)
    logger.info(f"Servidor HealthCheck activo en el puerto {Config.PORT}")
    server.serve_forever()

threading.Thread(target=iniciar_servidor_healthcheck, daemon=True).start()


# ==============================================================================
# 3. CLIENTE GOOGLE SHEETS CON CAPA DE CACHÉ Y REINTENTOS
# ==============================================================================
class CacheManager:
    def __init__(self, ttl_seconds: int = 300):
        self.ttl = ttl_seconds
        self._cache: Dict[str, Any] = {}
        self._timestamps: Dict[str, float] = {}

    def get(self, key: str) -> Optional[Any]:
        if key in self._cache:
            if time.time() - self._timestamps[key] < self.ttl:
                logger.info(f"Caché HIT para clave: '{key}'")
                return self._cache[key]
            else:
                logger.info(f"Caché EXPIRADO para clave: '{key}'")
                del self._cache[key]
                del self._timestamps[key]
        return None

    def set(self, key: str, value: Any):
        self._cache[key] = value
        self._timestamps[key] = time.time()
        logger.info(f"Caché ALMACENADO para clave: '{key}'")

    def invalidate_all(self):
        self._cache.clear()
        self._timestamps.clear()
        logger.info("Caché global invalidado.")

cache_sheets = CacheManager(ttl_seconds=Config.CACHE_TTL_SECONDS)

class SheetsRepository:
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    @classmethod
    def _obtener_cliente(cls):
        creds = Credentials.from_service_account_file(Config.CREDENTIALS_FILE, scopes=cls.SCOPES)
        return gspread.authorize(creds)

    @classmethod
    def obtener_filas_pestaña(cls, nombre_pestaña: str, force_refresh: bool = False) -> List[Dict[str, Any]]:
        cache_key = f"sheet_{nombre_pestaña.upper().strip()}"
        
        if not force_refresh:
            datos_cached = cache_sheets.get(cache_key)
            if datos_cached is not None:
                return datos_cached

        intentos = 3
        for intento in range(intentos):
            try:
                cliente = cls._obtener_cliente()
                libro = cliente.open(Config.SHEET_NAME)
                hoja = libro.worksheet(nombre_pestaña)
                registros = hoja.get_all_records()
                
                registros_normalizados = []
                for fila in registros:
                    fila_norm = {str(k).strip().upper(): str(v).strip() for k, v in fila.items()}
                    registros_normalizados.append(fila_norm)

                cache_sheets.set(cache_key, registros_normalizados)
                return registros_normalizados

            except Exception as e:
                logger.error(f"Intento {intento + 1} fallido al leer '{nombre_pestaña}': {e}")
                if intento < intentos - 1:
                    time.sleep(2 ** intento)
                else:
                    logger.critical(f"No se pudo conectar a Google Sheets tras {intentos} intentos.")
                    return []
    @classmethod
    def buscar_por_coincidencia(cls, nombre_pestaña: str, termino: str, campo_especifico: Optional[str] = None, columna_clave: Optional[str] = None) -> List[Dict[str, Any]]:
        campo_especifico = campo_especifico or columna_clave
        filas = cls.obtener_filas_pestaña(nombre_pestaña)
        if not termino:
            return filas

        termino_clean = str(termino).upper().strip()
        resultados = []

        for fila in filas:
            if campo_especifico:
                valor = fila.get(campo_especifico.upper(), "")
                if termino_clean in str(valor).upper():
                    resultados.append(fila)
            else:
                if any(termino_clean in str(val).upper() for val in fila.values()):
                    resultados.append(fila)
        return resultados 
# ==============================================================================
# 4. FÁBRICA DE COMPONENTES DE INTERFAZ (BLOCK KIT BUILDER)
# ==============================================================================
class BlockKitFactory:
    @staticmethod
    def crear_tarjeta_motor(titulo: str, datos: Optional[Dict[str, Any]]) -> List[Dict[str, Any]]:
        if not datos:
            return [
                {
                    "type": "section",
                    "text": {"type": "mrkdwn", "text": f"❌ *No se encontró información para:* `{titulo}`"}
                }
            ]

        cant = datos.get("CANTIDA", datos.get("CANTIDAD", "1"))
        kw = datos.get("KW", "N/A")
        amp = datos.get("A", datos.get("AMPERAJE", "N/A"))
        volts = datos.get("V", datos.get("VOLTAJE", "N/A"))
        conexion = datos.get("CONEXIÓN", datos.get("CONEXION", "N/A"))
        fp = datos.get("FP", "N/A")
        rpm = datos.get("RPM", "N/A")
        hz = datos.get("HZ", "N/A")
        rod_vent = datos.get("RODAMIENTO LADO VENTILADOF", datos.get("RODAMIENTO VENTILADOR", "N/A"))
        rod_bomba = datos.get("RODAMIENTO LADO BOMBA", "N/A")

        return [
            {
                "type": "header",
                "text": {"type": "plain_text", "text": f"⚙️ ESPECIFICACIONES: {titulo.upper()}", "emoji": True}
            },
            {"type": "divider"},
            {
                "type": "section",
                "fields": [
                    {"type": "mrkdwn", "text": f"*Cantidad:* {cant}"},
                    {"type": "mrkdwn", "text": f"*Potencia:* {kw} kW"},
                    {"type": "mrkdwn", "text": f"*Corriente Nominal:* {amp} A"},
                    {"type": "mrkdwn", "text": f"*Tensión:* {volts} V"},
                    {"type": "mrkdwn", "text": f"*Conexión:* {conexion}"},
                    {"type": "mrkdwn", "text": f"*Factor de Potencia:* {fp}"},
                    {"type": "mrkdwn", "text": f"*Velocidad:* {rpm} RPM"},
                    {"type": "mrkdwn", "text": f"*Frecuencia:* {hz} Hz"}
                ]
            },
            {
                "type": "section",
                "fields": [
                    {"type": "mrkdwn", "text": f"*Rodamiento Ventilación:* {rod_vent}"},
                    {"type": "mrkdwn", "text": f"*Rodamiento Lado Bomba:* {rod_bomba}"}
                ]
            },
            {"type": "divider"},
            {
                "type": "context",
                "elements": [
                    {"type": "mrkdwn", "text": f"📋 Fuente: Pestaña `DECANTERS` | Actualizado: {datetime.now().strftime('%H:%M:%S')}"}
                ]
            }
        ]

    @staticmethod
    def crear_lista_tablero(titulo_tablero: str, componentes: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        blocks = [
            {
                "type": "header",
                "text": {"type": "plain_text", "text": f"⚡ {titulo_tablero.upper()}", "emoji": True}
            },
            {"type": "divider"}
        ]

        if not componentes:
            blocks.append({
                "type": "section",
                "text": {"type": "mrkdwn", "text": "❌ No se registraron componentes para este tablero."}
            })
            return blocks

        cuerpo_texto = ""
        for idx, item in enumerate(componentes[:15], start=1):
            comp = item.get("TABLEROS", item.get("COMPONENTE", "N/A"))
            marca = item.get("MA", item.get("MARCA", "N/A"))
            modelo = item.get("MODELO", "N/A")
            cant = item.get("CANTID", item.get("CANTIDAD", "1"))
            link = item.get("LINK", "")

            linea = f"*{idx}. {comp}* — {marca} {modelo} `[Cant: {cant}]`"
            if link and link.upper() != "N/A":
                linea += f" | <{link}|📄 Ficha>"
            cuerpo_texto += linea + "\n"

        blocks.append({
            "type": "section",
            "text": {"type": "mrkdwn", "text": cuerpo_texto}
        })

        if len(componentes) > 15:
            blocks.append({
                "type": "context",
                "elements": [
                    {"type": "mrkdwn", "text": f"⚠️ Mostrando 15 de {len(componentes)} componentes registrados."}
                ]
            })

        blocks.append({"type": "divider"})
        return blocks


# ==============================================================================
# 5. GENERADOR DE REPORTES PDF (REPORTLAB ENGINE)
# ==============================================================================
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_page_decorations(self, page_count):
        self.saveState()
        self.setFont("Helvetica-Bold", 8)
        self.setFillColor(colors.HexColor("#003366"))
        
        # Encabezado superior del documento
        self.drawString(54, 750, "SISTEMA DE GESTIÓN DE MANTENIMIENTO - DRAGADOS RESITER")
        self.setStrokeColor(colors.HexColor("#CCCCCC"))
        self.setLineWidth(0.5)
        self.line(54, 742, 612 - 54, 742)

        # Pie de página
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#666666"))
        self.drawString(54, 36, "Documento de carácter técnico interno. Prohibida su reproducción no autorizada.")
        
        paginacion = f"Página {self._pageNumber} de {page_count}"
        self.drawRightString(612 - 54, 36, paginacion)
        self.restoreState()

class PDFReportGenerator:
    @staticmethod
    def generar_reporte_operacion(datos: List[Dict[str, Any]]) -> io.BytesIO:
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            buffer,
            pagesize=letter,
            leftMargin=54,
            rightMargin=54,
            topMargin=72,
            bottomMargin=54
        )

        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle(
            'ReportTitle',
            parent=styles['Heading1'],
            fontName='Helvetica-Bold',
            fontSize=16,
            leading=20,
            textColor=colors.HexColor("#003366"),
            spaceAfter=4
        )

        subtitle_style = ParagraphStyle(
            'ReportSubtitle',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=9,
            leading=12,
            textColor=colors.HexColor("#555555"),
            spaceAfter=15
        )

        cell_style = ParagraphStyle(
            'GridCell',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=8,
            leading=10,
            alignment=1
        )

        header_style = ParagraphStyle(
            'HeaderGridCell',
            parent=styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=8,
            leading=10,
            textColor=colors.whitesmoke,
            alignment=1
        )

        story = []

        # Título y metadata
        story.append(Paragraph("INFORME DE PARÁMETROS DE OPERACIÓN", title_style))
        fecha_str = datetime.now().strftime("%d/%m/%Y a las %H:%M:%S")
        story.append(Paragraph(f"Generado automáticamente por Bot Slack | Fecha: {fecha_str}", subtitle_style))
        story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#003366"), spaceAfter=15))

        # Configuración de tabla
        encabezados = ["EQUIPO", "VOLTAJE (V)", "AMPERAJE (A)", "RPM", "HZ", "CAUDAL"]
        data_tabla = [[Paragraph(h, header_style) for h in encabezados]]

        for fila in datos:
            eq = fila.get("EQUIPO", fila.get("EQUIPOS", "N/A"))
            v = fila.get("VOLTAJE", "N/A")
            a = fila.get("AMPERA", fila.get("AMPERAJE", "N/A"))
            rpm = fila.get("RPM", "N/A")
            hz = fila.get("HZ", "N/A")
            caudal = fila.get("CAUDAL", "N/A")

            data_tabla.append([
                Paragraph(str(eq), cell_style),
                Paragraph(str(v), cell_style),
                Paragraph(str(a), cell_style),
                Paragraph(str(rpm), cell_style),
                Paragraph(str(hz), cell_style),
                Paragraph(str(caudal), cell_style)
            ])

        t = Table(data_tabla, colWidths=[144, 72, 72, 72, 72, 72])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#003366")),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#D0D0D0")),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor("#F4F6F8")])
        ]))

        story.append(t)
        doc.build(story, canvasmaker=NumberedCanvas)
        buffer.seek(0)
        return buffer
        # ==============================================================================
# 6. INICIALIZACIÓN Y COMANDOS SLACK BOLT
# ==============================================================================
app = App(
    token=Config.SLACK_BOT_TOKEN,
    signing_secret=Config.SLACK_SIGNING_SECRET
)

def procesar_consulta_motor(nombre_motor: str, ack, respond):
    ack()
    try:
        res = SheetsRepository.buscar_por_coincidencia("DECANTERS", nombre_motor, columna_clave="EQUIPOS")
        datos = res[0] if res else None
        bloques = BlockKitFactory.crear_tarjeta_motor(nombre_motor, datos)
        respond(blocks=bloques)
    except Exception as e:
        logger.error(f"Error procesando comando para motor '{nombre_motor}': {e}")
        respond(f"❌ Error interno al consultar `{nombre_motor}`.")


# --- COMANDOS SLASH PARA MOTORES INDIVIDUALES ---

@app.command("/mtrdec")
def handle_mtrdec(ack, respond):
    procesar_consulta_motor("MOTOR DECANTER", ack, respond)

@app.command("/mtrback")
def handle_mtrback(ack, respond):
    procesar_consulta_motor("MOTOR BAK DRIVE", ack, respond)

@app.command("/mtrco")
def handle_mtrco(ack, respond):
    procesar_consulta_motor("MOTOR CORREA", ack, respond)

@app.command("/mtrbbaa")
def handle_mtrbbaa(ack, respond):
    procesar_consulta_motor("MOTOR ALIMENTACION", ack, respond)

@app.command("/mtrbbal")
def handle_mtrbbal(ack, respond):
    procesar_consulta_motor("MOTOR LIMPIEZA", ack, respond)


# --- BUSCADOR MULTIPROPÓSITO DE EQUIPOS ---

@app.command("/equipos")
def handle_equipos(ack, respond, command):
    ack()
    busqueda = command.get("text", "").strip()

    if not busqueda:
        respond("⚠️ *Uso del comando:* `/equipos <nombre o término>`\n*Ejemplo:* `/equipos BOMBA`")
        return

    try:
        resultados = SheetsRepository.buscar_por_coincidencia("DECANTERS", busqueda)
        if not resultados:
            respond(f"🔍 No se encontraron coincidencias para `{busqueda}` en la pestaña `DECANTERS`.")
            return

        msg = f"🔎 *Resultados encontrados para '{busqueda}' ({len(resultados)} coincidencia(s)):*\n\n"
        for item in resultados[:10]:
            eq = item.get("EQUIPOS", "Sin nombre")
            kw = item.get("KW", "N/A")
            amp = item.get("A", "N/A")
            volts = item.get("V", "N/A")
            rpm = item.get("RPM", "N/A")
            msg += f"• *{eq}* | Potencia: `{kw} kW` | Corr: `{amp} A` | Tensión: `{volts} V` | `{rpm} RPM`\n"

        if len(resultados) > 10:
            msg += f"\n_...y {len(resultados) - 10} resultados adicionales. Refina tu búsqueda._"

        respond(msg)
    except Exception as e:
        logger.error(f"Error en comando /equipos con búsqueda '{busqueda}': {e}")
        respond("❌ Error al realizar la búsqueda en la base de datos.")


# --- COMANDOS PARA TABLEROS ELÉCTRICOS ---

import threading  # Asegúrate de tener esta importación al inicio de bot.py

@app.command("/tablerodec2")
@app.command("/tablerodec3")
@app.command("/tableros")
def handle_tableros(ack, respond, command):
    ack()  # Le responde a Slack instantáneamente
    
    def procesar_tarea():
        comando_usado = command.get("command", "/tableros")
        try:
            componentes = SheetsRepository.obtener_filas_pestaña("TABLEROS ELECTRICOS")
            bloques = BlockKitFactory.crear_lista_tablero(f"Componentes {comando_usado}", componentes)
            respond(blocks=bloques)
        except Exception as e:
            logger.error(f"Error en comando {comando_usado}: {e}")
            respond(f"❌ Ocurrió un error al cargar la lista de tableros eléctricos: {e}")

    # Ejecuta la lectura de Google Sheets en segundo plano
    threading.Thread(target=procesar_tarea).start()


# --- GENERADOR Y ENVIADOR DE REPORTES PDF (/resumen) ---

@app.command("/resumen")
def handle_resumen(ack, respond, client, command):
    ack()
    channel_id = command.get("channel_id")
    user_id = command.get("user_id")

    respond("⏳ *Generando informe en PDF...* Por favor espera un momento.")

    try:
        datos_operacion = SheetsRepository.obtener_filas_pestaña("PARAMETROS DE OPERACION")
        
        if not datos_operacion:
            client.chat_postMessage(
                channel=channel_id,
                text=f"❌ <@{user_id}>, no se encontraron registros en la pestaña `PARAMETROS DE OPERACION`."
            )
            return

        pdf_buffer = PDFReportGenerator.generar_reporte_operacion(datos_operacion)
        nombre_archivo = f"Reporte_Parametros_Operacion_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"

        client.files_upload_v2(
            channel=command["channel_id"],
            file=pdf_buffer,
            filename=nombre_archivo,
            title="Informe de Parámetros de Operación",
            initial_comment=f"📄 *Informe de Parámetros de Operación solicitado por <@{user_id}>*"
        )
        logger.info(f"Reporte PDF subido exitosamente por el usuario {user_id}")

    except SlackApiError as e:
        logger.error(f"Error de API de Slack al subir PDF: {e.response['error']}")
        client.chat_postMessage(
            channel=channel_id,
            text=f"❌ Error al enviar el archivo a Slack: `{e.response['error']}`"
        )
    except Exception as e:
        logger.error(f"Error general en comando /resumen: {e}")
        client.chat_postMessage(
            channel=channel_id,
            text="❌ Se produjo un fallo inesperado al construir el reporte PDF."
        )


# --- COMANDO DE AYUDA Y ESTADO DEL SISTEMA ---

@app.command("/ayuda")
@app.command("/status")
def handle_ayuda_status(ack, respond):
    ack()
    msg = (
        "🛠️ *SISTEMA DE MANTENIMIENTO DRAGADOS - COMANDOS DISPONIBLES*\n\n"
        "• `/mtrdec` - Consulta ficha técnica del Motor Decanter\n"
        "• `/mtrback` - Consulta ficha técnica del Motor Back Drive\n"
        "• `/mtrco` - Consulta ficha técnica del Motor Correa\n"
        "• `/mtrbbaa` - Consulta ficha técnica del Motor Bomba Alimentación\n"
        "• `/mtrbbal` - Consulta ficha técnica del Motor Bomba Limpieza\n"
        "• `/equipos <texto>` - Búsqueda general en la pestaña Decanters\n"
        "• `/tablerodec2` - Muestra componentes en Tableros Eléctricos\n"
        "• `/resumen` - Genera y descarga el PDF de Parámetros de Operación\n"
        "• `/status` - Verifica el estado de servicio y refresco de caché\n"
    )
    respond(msg)


# ==============================================================================
# 7. MANEJO GLOBAL DE ERRORES Y PUNTO DE ENTRADA
# ==============================================================================
@app.error
def global_error_handler(error, body, logger_bolt):
    logger.error(f"Error no controlado en la APP: {error}")
    logger.debug(f"Cuerpo del evento con error: {body}")

if __name__ == "__main__":
    if not Config.SLACK_APP_TOKEN:
        logger.critical("No es posible iniciar SocketMode sin SLACK_APP_TOKEN.")
    else:
        logger.info("Iniciando Bot en Socket Mode (Mantenimiento Dragados Resiter)...")
        handler = SocketModeHandler(app, Config.SLACK_APP_TOKEN)
        handler.start()
