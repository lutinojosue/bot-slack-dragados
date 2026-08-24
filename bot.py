import os
import re
import time
import threading
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta

import gspread
from google.oauth2.service_account import Credentials

from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler

# Librerías para generar PDF
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# -------------------------------------------------------------
# 1. CONFIGURACIÓN Y TOKENS DE SLACK Y GOOGLE SHEETS
# -------------------------------------------------------------
BOT_TOKEN = os.environ.get("BOT_TOKEN")
APP_TOKEN = os.environ.get("APP_TOKEN")
CANAL_SLACK = "#todo-bot-dragados"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_PATH = os.path.join(BASE_DIR, "credentials.json")
NOMBRE_SHEET_DRIVE = "Mantencion-Componentes-Dragados"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

app = App(token=BOT_TOKEN)

def obtener_sheet():
    creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=SCOPES)
    gc = gspread.authorize(creds)
    return gc.open(NOMBRE_SHEET_DRIVE)

# -------------------------------------------------------------
# 2. FUNCIONES DE FECHAS Y BÚSQUEDA EXACTA
# -------------------------------------------------------------
def obtener_fecha_valida(celda_val):
    if celda_val is None:
        return None
    if isinstance(celda_val, datetime):
        return celda_val.date()
    if isinstance(celda_val, date):
        return celda_val
    if isinstance(celda_val, (int, float)):
        try:
            return (datetime(1899, 12, 30) + timedelta(days=celda_val)).date()
        except Exception:
            pass
    if isinstance(celda_val, str):
        s = celda_val.strip()
        if not s or s.upper() in ["N/D", "N/A", "NONE", "-"]:
            return None
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y", "%d-%m-%Y", "%d.%m.%Y"):
            try:
                return datetime.strptime(s, fmt).date()
            except ValueError:
                pass
    return None

def extraer_texto_y_link_str(val_str):
    val = str(val_str or "").strip()
    link = None
    
    if val.startswith("http://") or val.startswith("https://"):
        link = val
    elif val.startswith("www."):
        link = f"http://{val}"
    elif val.startswith("="):
        match = re.search(r'(?:HYPERLINK|HIPERVINCULO)\s*\(\s*["\']([^"\']+)["\']', val, re.IGNORECASE)
        if match:
            link = match.group(1)

    return val, link

def calcular_proxima_fecha(nombre_equipo, fecha_base):
    nombre_upper = nombre_equipo.upper().strip()
    if "EJE" in nombre_upper or "CORREA" in nombre_upper:
        return fecha_base + timedelta(days=1)
    else:
        return fecha_base + relativedelta(months=3)

def encontrar_columnas_fechas_filas(filas):
    col_ult = None
    col_prox = None
    col_sello = None

    for r_idx in range(min(4, len(filas))):
        for c_idx, val_raw in enumerate(filas[r_idx]):
            val = str(val_raw or "").strip().upper()
            if not val:
                continue
            
            c = c_idx + 1
            if ("ULT" in val or "ÚLT" in val) and ("ENGRAS" in val or "FECHA" in val or "MANT" in val or "MANTENCION" in val):
                col_ult = c
            elif "ULTI" in val or "ÚLTIMA" in val:
                if "PROX" not in val and "PRÓX" not in val:
                    col_ult = c
            
            if ("PROX" in val or "PRÓX" in val) and ("ENGRAS" in val or "FECHA" in val or "MANT" in val or "MANTENCION" in val):
                col_prox = c

            if "SELLO" in val or "ACREDIT" in val:
                col_sello = c

    if col_ult and col_ult <= 3:
        col_ult = None
    if col_prox and col_prox <= 3:
        col_prox = None

    if not col_ult and not col_prox and not col_sello:
        col_ult = 16
        col_prox = 17

    return col_ult, col_prox, col_sello

def es_mismo_equipo_exacto(target, celda):
    t = str(target or "").upper().strip()
    c = str(celda or "").upper().strip()
    
    if not c or c.startswith("TOTAL"):
        return False
    
    if t == c:
        return True
    
    t_clean = t.replace("BAK", "BACK").replace(" ", "").replace("-", "")
    c_clean = c.replace("BAK", "BACK").replace(" ", "").replace("-", "")
    
    return t_clean == c_clean

def actualizar_excel_equipo(nombre_equipo):
    try:
        sh = obtener_sheet()
        hoy = date.today()
        nueva_prox = calcular_proxima_fecha(nombre_equipo, hoy)
        
        nueva_prox_str = nueva_prox.strftime("%d/%m/%Y")
        hoy_str = hoy.strftime("%d/%m/%Y")

        for ws in sh.worksheets():
            filas = ws.get_all_values()
            if len(filas) < 4:
                continue

            col_ult, col_prox, _ = encontrar_columnas_fechas_filas(filas)
            if not col_ult or not col_prox:
                continue

            for row_idx, fila in enumerate(filas[3:], start=4):
                equipo_celda = str(fila[0] if fila else "").strip()
                if es_mismo_equipo_exacto(nombre_equipo, equipo_celda):
                    ws.update_cell(row_idx, col_ult, hoy_str)
                    ws.update_cell(row_idx, col_prox, nueva_prox_str)
                    return nueva_prox_str
        return nueva_prox_str
    except Exception as e:
        print(f"Error actualizando Google Sheets: {e}")
        return date.today().strftime("%d/%m/%Y")

# -------------------------------------------------------------
# 3. EXTRAER DATOS Y COMPONENTES DE TABLEROS Y GENERADORES
# -------------------------------------------------------------
def obtener_datos_tablero(nombre_tablero):
    try:
        sh = obtener_sheet()
        resumen_tablero = {}
        componentes = []

        sheets_to_check = [sh.worksheet("TABLEROS ELECTRICOS")] if "TABLEROS ELECTRICOS" in [w.title for w in sh.worksheets()] else sh.worksheets()

        for ws in sheets_to_check:
            filas = ws.get_all_values()
            if len(filas) < 3:
                continue

            headers = [str(h or "").strip() for h in filas[2]]
            tablero_encontrado = False

            for fila in filas[3:]:
                col_1 = str(fila[0] if fila else "").strip().upper()

                if "TABLERO" in col_1 and es_mismo_equipo_exacto(nombre_tablero, col_1):
                    tablero_encontrado = True
                    for col_idx, val_cell in enumerate(fila):
                        val, _ = extraer_texto_y_link_str(val_cell)
                        h = headers[col_idx].upper() if col_idx < len(headers) else f"COL {col_idx+1}"
                        if val != "":
                            resumen_tablero[h] = val
                    continue

                if tablero_encontrado:
                    if col_1.startswith("TABLERO") and not es_mismo_equipo_exacto(nombre_tablero, col_1):
                        break
                    
                    if col_1:
                        comp_nombre = col_1
                        specs = {}
                        for col_idx, val_cell in enumerate(fila):
                            val, link = extraer_texto_y_link_str(val_cell)
                            h = headers[col_idx] if col_idx < len(headers) else f"Col {col_idx+1}"
                            if val != "" or link is not None:
                                specs[h] = {"val": val, "link": link}
                        componentes.append({"nombre": comp_nombre, "specs": specs})

            if tablero_encontrado:
                break

        return resumen_tablero, componentes
    except Exception as e:
        print(f"Error obteniendo tablero: {e}")
        return {}, []

def desplegar_menu_tablero(respond, nombre_tablero):
    resumen, componentes = obtener_datos_tablero(nombre_tablero)

    if not resumen and not componentes:
        respond(f"❌ No se encontraron datos para el *{nombre_tablero}* en Google Sheets.")
        return

    ult_mant = resumen.get("ULTIMA MANTENCION", resumen.get("ULTI ENGRAS", "N/D"))
    prox_mant = resumen.get("PROXIMA MANTENCION", resumen.get("PROX ENGRAS", "N/D"))
    venc_sello = resumen.get("VENC ACREDITACION", resumen.get("VENC SELLO", resumen.get("VENCIMIENTO SELLO", "N/D")))

    options = []
    for idx, comp in enumerate(componentes[:100]):
        marca_spec = comp["specs"].get("MAR", comp["specs"].get("MARCA", {}))
        marca = marca_spec.get("val", "") if isinstance(marca_spec, dict) else ""
        
        modelo_spec = comp["specs"].get("MODELO", {})
        modelo = modelo_spec.get("val", "") if isinstance(modelo_spec, dict) else ""

        etiqueta = f"{comp['nombre']} | {marca} {modelo}".strip()[:75]
        options.append({
            "text": {"type": "plain_text", "text": etiqueta},
            "value": f"{nombre_tablero}|{idx}"
        })

    blocks = [
        {
            "type": "header",
            "text": {"type": "plain_text", "text": f"🎛️ Estado Principal: {nombre_tablero.upper()}"}
        },
        {
            "type": "section",
            "fields": [
                {"type": "mrkdwn", "text": f"*🗓️ Última Mantención:*\n`{ult_mant}`"},
                {"type": "mrkdwn", "text": f"*🗓️ Próxima Mantención:*\n`{prox_mant}`"},
                {"type": "mrkdwn", "text": f"*🔒 Vencimiento Acreditación:*\n`{venc_sello}`"}
            ]
        },
        {"type": "divider"}
    ]

    if options:
        blocks.append({
            "type": "section",
            "text": {"type": "mrkdwn", "text": "*Selecciona un componente para desplegar especificaciones completas y link:*"},
            "accessory": {
                "type": "static_select",
                "placeholder": {"type": "plain_text", "text": "Ver componentes..."},
                "action_id": "seleccionar_componente_tablero",
                "options": options
            }
        })
    else:
        blocks.append({
            "type": "section",
            "text": {"type": "mrkdwn", "text": "_No hay componentes individuales registrados en este tablero._"}
        })

    respond(blocks=blocks, text=f"Resumen de {nombre_tablero}")

def obtener_datos_generadores():
    try:
        sh = obtener_sheet()
        hoja_gen = None
        for sheet in sh.worksheets():
            if "GENERADOR" in sheet.title.upper() or "ELECTROGENO" in sheet.title.upper():
                hoja_gen = sheet
                break
        
        if not hoja_gen:
            return []

        filas = hoja_gen.get_all_values()
        if len(filas) < 3:
            return []

        headers = [str(h or "").strip() for h in filas[2]]

        generadores = []
        for fila in filas[3:]:
            equipo = str(fila[0] if fila else "").strip()
            if not equipo or equipo.upper().startswith("TOTAL") or equipo.upper().startswith("EQUIPOS"):
                continue

            specs = {}
            for col_idx, val_cell in enumerate(fila):
                val, link = extraer_texto_y_link_str(val_cell)
                header = headers[col_idx] if col_idx < len(headers) and headers[col_idx] else f"Col {col_idx+1}"
                if val != "" or link is not None:
                    specs[header] = {"val": val, "link": link}

            generadores.append({"nombre": equipo, "specs": specs})

        return generadores
    except Exception as e:
        print(f"Error obteniendo generadores: {e}")
        return []

def desplegar_menu_generadores(respond):
    generadores = obtener_datos_generadores()
    if not generadores:
        respond("❌ No se encontraron datos en la pestaña *GENERADORES* de Google Sheets.")
        return

    options = []
    for idx, gen in enumerate(generadores[:100]):
        marca_spec = gen["specs"].get("MARCA", {})
        marca = marca_spec.get("val", "") if isinstance(marca_spec, dict) else ""
        kw_spec = gen["specs"].get("KW", {})
        kw = kw_spec.get("val", "") if isinstance(kw_spec, dict) else ""
        
        det = f"{marca} {kw}KW".strip()
        etiqueta = f"{gen['nombre']} | {det}".strip(" |")[:75]
        options.append({
            "text": {"type": "plain_text", "text": etiqueta},
            "value": f"GENERADORES|{idx}"
        })

    blocks = [
        {
            "type": "header",
            "text": {"type": "plain_text", "text": "⚡ EQUIPOS ELECTROGÉNOS Y GENERADORES"}
        },
        {
            "type": "section",
            "text": {"type": "mrkdwn", "text": f"Se encontraron *{len(generadores)}* equipos registrados."}
        },
        {"type": "divider"},
        {
            "type": "section",
            "text": {"type": "mrkdwn", "text": "*Selecciona un generador o motobomba para ver su ficha técnica:*"},
            "accessory": {
                "type": "static_select",
                "placeholder": {"type": "plain_text", "text": "Ver generador..."},
                "action_id": "seleccionar_generador",
                "options": options
            }
        }
    ]
    respond(blocks=blocks, text="Lista de Generadores")

    # -------------------------------------------------------------
# 4. BÚSQUEDA GLOBAL Y REPORTES PDF
# -------------------------------------------------------------
def buscar_componente_global(busqueda):
    try:
        sh = obtener_sheet()
        query_clean = busqueda.strip().upper().replace("!", "").replace("/", "")
        resultados = []

        for ws in sh.worksheets():
            filas = ws.get_all_values()
            if not filas:
                continue

            headers = filas[2] if len(filas) >= 3 else filas[0]

            for fila in filas[3 if len(filas) >= 3 else 1:]:
                fila_texto = " ".join([str(v or "") for v in fila]).upper()

                if query_clean in fila_texto:
                    item_info = [f"📁 *Sección / Pestaña:* `{ws.title}`"]
                    for col_idx, val_cell in enumerate(fila):
                        val, link = extraer_texto_y_link_str(val_cell)
                        header = headers[col_idx] if col_idx < len(headers) and headers[col_idx] else f"Columna {col_idx+1}"
                        
                        if val != "":
                            if link:
                                item_info.append(f"• *{header}*: <{link}|🔗 Abrir Link de Compra / Ficha>")
                            elif val.upper() != "ENLACE":
                                item_info.append(f"• *{header}*: {val}")
                    
                    resultados.append("\n".join(item_info))

        if not resultados:
            return f"❌ No se encontraron coincidencias para `'{busqueda}'` en Google Sheets."

        return f"🔍 *Resultados para '{busqueda}':*\n\n" + "\n\n---\n\n".join(resultados[:5])
    except Exception as e:
        return f"⚠️ Error buscando en Google Sheets: {e}"

def crear_pdf_resumen():
    pdf_path = os.path.join(BASE_DIR, "Resumen_Mantenciones.pdf")
    try:
        sh = obtener_sheet()
        hoy = date.today()
        doc = SimpleDocTemplate(pdf_path, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
        elements = []
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle(
            'TitleStyle', parent=styles['Heading1'], fontSize=16, textColor=colors.HexColor("#1D1C1D"), spaceAfter=10, alignment=1
        )
        
        elements.append(Paragraph("<b>Reporte General de Mantención, Engrase y Acreditaciones</b>", title_style))
        elements.append(Paragraph(f"<b>Fecha de emisión:</b> {hoy.strftime('%d/%m/%Y')}", styles['Normal']))
        elements.append(Spacer(1, 15))

        table_data = [["Sección", "Equipo / Componente", "Últ. Fecha", "Próx. Fecha", "Acreditación"]]

        for ws in sh.worksheets():
            filas = ws.get_all_values()
            if len(filas) < 4:
                continue

            col_ult, col_prox, col_sello = encontrar_columnas_fechas_filas(filas)

            for fila in filas[3:]:
                equipo = str(fila[0] if fila else "").strip()
                if not equipo or equipo.upper().startswith("TOTAL") or equipo.upper().startswith("EQUIPOS"):
                    continue

                ult = fila[col_ult - 1] if col_ult and (col_ult - 1) < len(fila) else "N/D"
                prox = fila[col_prox - 1] if col_prox and (col_prox - 1) < len(fila) else "N/D"
                sello = fila[col_sello - 1] if col_sello and (col_sello - 1) < len(fila) else "N/D"

                ult_str = str(ult) if ult else "N/D"
                prox_str = str(prox) if prox else "N/D"
                sello_str = str(sello) if sello else "N/D"

                if ult_str != "N/D" or prox_str != "N/D" or sello_str != "N/D":
                    table_data.append([
                        str(ws.title)[:15],
                        equipo[:25],
                        ult_str,
                        prox_str,
                        sello_str
                    ])

        if len(table_data) <= 1:
            return None

        t = Table(table_data, colWidths=[80, 150, 85, 85, 90])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#007A5A")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CCCCCC")),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))
        elements.append(t)
        doc.build(elements)
        return pdf_path
    except Exception as e:
        print(f"Error generando PDF Resumen: {e}")
        return None

def crear_pdf_parametros():
    pdf_path = os.path.join(BASE_DIR, "Resumen_Parametros_Operacion.pdf")
    try:
        sh = obtener_sheet()
        hoja_param = None
        for sheet in sh.worksheets():
            if "PARAMETRO" in sheet.title.upper():
                hoja_param = sheet
                break

        if not hoja_param:
            return None

        filas = hoja_param.get_all_values()
        if not filas:
            return None

        hoy = date.today()
        doc = SimpleDocTemplate(pdf_path, pagesize=landscape(letter), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
        elements = []
        styles = getSampleStyleSheet()

        title_style = ParagraphStyle(
            'TitleStyle', parent=styles['Heading1'], fontSize=16, textColor=colors.HexColor("#1D1C1D"), spaceAfter=10, alignment=1
        )

        elements.append(Paragraph("<b>Reporte de Parámetros de Operación</b>", title_style))
        elements.append(Paragraph(f"<b>Pestaña:</b> {hoja_param.title} | <b>Fecha de emisión:</b> {hoy.strftime('%d/%m/%Y')}", styles['Normal']))
        elements.append(Spacer(1, 15))

        header_row_idx = 2 if len(filas) >= 3 else 0
        headers = [str(v or "").strip() for v in filas[header_row_idx] if str(v or "").strip()]

        table_data = []
        cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=7, leading=8)
        cell_style_bold = ParagraphStyle('CellStyleBold', parent=styles['Normal'], fontSize=7, leading=8, textColor=colors.whitesmoke)

        if headers:
            table_data.append([Paragraph(f"<b>{h}</b>", cell_style_bold) for h in headers])

        for fila in filas[header_row_idx + 1:]:
            row_vals = []
            has_data = False
            for col_idx in range(len(headers)):
                val = fila[col_idx] if col_idx < len(fila) else ""
                val_str = str(val).strip() if val is not None else ""
                if val_str:
                    has_data = True
                row_vals.append(Paragraph(val_str if val_str else "-", cell_style))
            
            if has_data:
                table_data.append(row_vals)

        if len(table_data) <= 1:
            return None

        t = Table(table_data)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#007A5A")),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CCCCCC")),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ]))
        elements.append(t)
        doc.build(elements)
        return pdf_path
    except Exception as e:
        print(f"Error generando PDF Parametros: {e}")
        return None

    # -------------------------------------------------------------
# 5. COMANDOS DE SLACK
# -------------------------------------------------------------
@app.command("/resumen")
def cmd_resumen(ack, respond, client, body):
    ack()
    respond("📄 Generando informe PDF general de equipos y tableros desde Google Sheets, un momento...")
    pdf_path = crear_pdf_resumen()
    if pdf_path and os.path.exists(pdf_path):
        try:
            client.files_upload_v2(
                channel=body["channel_id"],
                file=pdf_path,
                title="Resumen General de Mantenimiento",
                initial_comment="📊 *Aquí tienes el reporte general en PDF:*"
            )
        except Exception as e:
            respond(f"❌ Error al enviar PDF: {e}")

@app.command("/parametros")
@app.command("/param")
@app.command("/parametrosoperacion")
def cmd_parametros(ack, respond, client, body):
    ack()
    respond("⚙️ Generando informe PDF de *Parámetros de Operación*, un momento...")
    pdf_path = crear_pdf_parametros()
    if pdf_path and os.path.exists(pdf_path):
        try:
            client.files_upload_v2(
                channel=body["channel_id"],
                file=pdf_path,
                title="Parámetros de Operación",
                initial_comment="📊 *Aquí tienes la planilla de Parámetros de Operación en PDF:*"
            )
        except Exception as e:
            respond(f"❌ Error al enviar el PDF de parámetros: {e}")
    else:
        respond("❌ No se encontraron datos o no se pudo generar la pestaña de Parámetros de Operación.")

@app.command("/tablerodec1")
@app.command("/decanter1")
def cmd_tablerodec1(ack, respond):
    ack()
    desplegar_menu_tablero(respond, "TABLERO DECANTER 1")

@app.command("/tablerodec2")
@app.command("/decanter2")
def cmd_tablerodec2(ack, respond):
    ack()
    desplegar_menu_tablero(respond, "TABLERO DECANTER 2-3")

@app.command("/tablerodec3")
@app.command("/decanter3")
def cmd_tablerodec3(ack, respond):
    ack()
    desplegar_menu_tablero(respond, "TABLERO DECANTER 2-3")

@app.command("/draga")
def cmd_draga(ack, respond):
    ack()
    desplegar_menu_tablero(respond, "TABLERO DRAGA")

@app.command("/generadores")
@app.command("/generador")
@app.command("/gen")
def cmd_generadores(ack, respond):
    ack()
    desplegar_menu_generadores(respond)

@app.command("/buscar")
@app.command("/equipo")
@app.command("/componente")
def cmd_buscar_global(ack, respond, command):
    ack()
    query = command.get("text", "").strip()
    if not query:
        respond("⚠️ Ingresa el parámetro a buscar. Ejemplo: `/buscar CFW-500`")
        return
    respond(buscar_componente_global(query))

@app.command("/alertas")
def cmd_forzar_alertas(ack, respond):
    ack()
    respond("🔍 Buscando mantenciones, engrases y acreditaciones vencidas...")
    enviadas = enviar_alertas()
    if enviadas == 0:
        respond("ℹ️ No se encontraron mantenciones o acreditaciones vencidas a la fecha de hoy.")

@app.message(r"^!")
def responder_codigo_admiracion(message, say):
    texto = message.get("text", "").strip()
    say(text=buscar_componente_global(texto))

# -------------------------------------------------------------
# 6. HANDLERS DE INTERACCIÓN (MENÚS Y BOTONES)
# -------------------------------------------------------------
@app.action("seleccionar_componente_tablero")
def responder_seleccion_componente(ack, body, client):
    ack()
    seleccion = body["actions"][0]["selected_option"]["value"]
    nombre_tablero, idx_str = seleccion.split("|")
    idx = int(idx_str)

    _, componentes = obtener_datos_tablero(nombre_tablero)
    if idx < len(componentes):
        comp = componentes[idx]
        
        lineas = [f"⚙️ *Ficha Técnica de Componente: {comp['nombre']}*"]
        lineas.append(f"📍 *Pertenece a:* `{nombre_tablero}`\n")

        for k, item in comp["specs"].items():
            if k.upper() not in ["TABLEROS", "EQUIPOS"]:
                val = item.get("val", "")
                link = item.get("link", None)
                if link:
                    lineas.append(f"• *{k}*: <{link}|🔗 Abrir Link de Compra / Ficha Técnica>")
                elif val and val.upper() != "ENLACE":
                    lineas.append(f"• *{k}*: {val}")

        client.chat_postEphemeral(
            channel=body["channel"]["id"],
            user=body["user"]["id"],
            text="\n".join(lineas)
        )

@app.action("seleccionar_generador")
def responder_seleccion_generador(ack, body, client):
    ack()
    seleccion = body["actions"][0]["selected_option"]["value"]
    _, idx_str = seleccion.split("|")
    idx = int(idx_str)

    generadores = obtener_datos_generadores()
    if idx < len(generadores):
        gen = generadores[idx]
        
        lineas = [f"⚡ *Ficha Técnica: {gen['nombre']}*"]
        lineas.append("📍 *Pestaña:* `GENERADORES`\n")

        for k, item in gen["specs"].items():
            val = item.get("val", "")
            link = item.get("link", None)
            if link:
                lineas.append(f"• *{k}*: <{link}|🔗 Abrir Link / Ficha>")
            elif val and val.upper() != "ENLACE":
                lineas.append(f"• *{k}*: {val}")

        client.chat_postEphemeral(
            channel=body["channel"]["id"],
            user=body["user"]["id"],
            text="\n".join(lineas)
        )

@app.action("btn_confirmar_mantenimiento")
def responder_boton(ack, body, client):
    ack()
    try:
        equipo = body["actions"][0]["value"]
        usuario = body["user"]["name"]
        proxima = actualizar_excel_equipo(equipo)

        client.chat_update(
            channel=body["channel"]["id"],
            ts=body["message"]["ts"],
            text=f"✅ Mantenimiento registrado para {equipo}",
            blocks=[{
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": f"✅ *¡Puesto al día en Google Sheets!* @{usuario} confirmó mantención/engrase para *{equipo}*.\n📅 *Última Fecha:* `{datetime.now().strftime('%d/%m/%Y')}`\n📅 *Próxima Fecha:* `{proxima}`"
                }
            }]
        )
    except Exception as e:
        print(f"Error procesando botón: {e}")

        # -------------------------------------------------------------
# 7. SISTEMA DE ALERTAS Y MOTOR DE VERIFICACIÓN
# -------------------------------------------------------------
def enviar_alertas():
    try:
        sh = obtener_sheet()
        hoy = date.today()
        alertas_enviadas = 0

        for ws in sh.worksheets():
            filas = ws.get_all_values()
            if len(filas) < 4:
                continue

            col_ult, col_prox, col_sello = encontrar_columnas_fechas_filas(filas)

            if not col_prox and not col_sello:
                continue

            for fila in filas[3:]:
                equipo_raw = str(fila[0] if fila else "").strip()
                if not equipo_raw or equipo_raw.upper().startswith("TOTAL") or equipo_raw.upper().startswith("EQUIPOS"):
                    continue

                if col_prox:
                    prox_val = fila[col_prox - 1] if (col_prox - 1) < len(fila) else None
                    fecha_dt = obtener_fecha_valida(prox_val)

                    if fecha_dt and fecha_dt <= hoy:
                        dias_diferencia = (fecha_dt - hoy).days
                        estado_texto = "🟢 *¡Programado para Hoy!*" if dias_diferencia == 0 else f"⚠️ *¡ATRASADO por {abs(dias_diferencia)} día(s)!*"
                        
                        bloques = [
                            {"type": "header", "text": {"type": "plain_text", "text": f"⚙️ Alerta Programada: {equipo_raw}"}},
                            {"type": "section", "fields": [
                                {"type": "mrkdwn", "text": f"*Estado:*\n{estado_texto}"},
                                {"type": "mrkdwn", "text": f"*Ubicación:*\nPestaña `{ws.title}`"}
                            ]},
                            {"type": "actions", "elements": [{
                                "type": "button",
                                "text": {"type": "plain_text", "text": "🟢 Confirmar Realizado"},
                                "style": "primary",
                                "value": str(equipo_raw),
                                "action_id": "btn_confirmar_mantenimiento"
                            }]}
                        ]
                        try:
                            app.client.chat_postMessage(channel=CANAL_SLACK, blocks=bloques, text=f"Alerta {equipo_raw}")
                            alertas_enviadas += 1
                        except Exception as e:
                            print(f"❌ Error enviando alerta a Slack para {equipo_raw}: {e}")

                if col_sello:
                    sello_val = fila[col_sello - 1] if (col_sello - 1) < len(fila) else None
                    fecha_sello = obtener_fecha_valida(sello_val)
                    if fecha_sello:
                        dias = (fecha_sello - hoy).days
                        if 0 <= dias <= 15:
                            msg = f"⚠️ *ALERTA DE ACREDITACIÓN:* El sello/acreditación de *{equipo_raw}* (`{ws.title}`) vence en {dias} días (`{fecha_sello.strftime('%d/%m/%Y')}`)."
                            try:
                                app.client.chat_postMessage(channel=CANAL_SLACK, text=msg)
                                me_enviadas += 1
                            except Exception as e:
                                print(f"❌ Error enviando alerta de sello a Slack: {e}")

        return alertas_enviadas
    except Exception as e:
        print(f"Error en motor de alertas: {e}")
        return 0

def bucle_verificacion_continua():
    ultima_fecha_ejecucion = None
    while True:
        ahora = datetime.now()
        hoy_date = ahora.date()
        if (ahora.hour > 7 or (ahora.hour == 7 and ahora.minute >= 40)) and ultima_fecha_ejecucion != hoy_date:
            enviar_alertas()
            ultima_fecha_ejecucion = hoy_date
        time.sleep(300)

# -------------------------------------------------------------
# 8. EJECUCIÓN PRINCIPAL
# -------------------------------------------------------------
if __name__ == "__main__":
    handler = SocketModeHandler(app, APP_TOKEN)
    threading.Thread(target=bucle_verificacion_continua, daemon=True).start()
    
    print("⚡ Bot en ejecución correctamente. Escuchando comandos y alertas...")
    
    while True:
        try:
            handler.start()
        except Exception as e:
            time.sleep(5)