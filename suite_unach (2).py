
import os
import time
import threading
from datetime import datetime

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from PIL import Image, ImageTk
    PIL_DISPONIBLE = True
except ImportError:
    PIL_DISPONIBLE = False

import pandas as pd
import numpy as np
import openpyxl

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import imaplib
import smtplib
import email
from email.header import decode_header
from email.mime.text import MIMEText
import json
import urllib.request


MI_GMAIL = "m.kleiner.2006@gmail.com"
MI_PASSWORD_DE_APLICACION = "kdohjprhwnacmcad"
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass
MI_GEMINI_TOKEN = os.environ.get("GEMINI_API_KEY")


NAVY            = RGBColor(0x0F, 0x17, 0x2A)
NAVY_HEX        = "0F172A"
CIAN            = RGBColor(0x0E, 0xA5, 0xE9)
CIAN_HEX        = "0EA5E9"
VIOLETA         = RGBColor(0x8B, 0x5C, 0xF6)
VIOLETA_HEX     = "8B5CF6"
GRIS_TXT        = RGBColor(0x47, 0x55, 0x69)
GRIS_HEX        = "475569"
GRIS_CLARO_HEX  = "F1F5F9"
BLANCO_HEX      = "FFFFFF"
FUENTE = "Segoe UI"

BARRAS_GRAFICO = ["#0EA5E9", "#8B5CF6", "#38BDF8", "#A855F7", "#22D3EE"]

# ==============================================================================
# PALETA DE DISENO - PANTALLA DE ACCESO (Login / Registro / Recuperacion)
# Identidad institucional UNACH: baige, blanco y amarillo claro
# ==============================================================================
BEIGE_HEX           = "F5EFE1"   # fondo panel institucional
BEIGE_OSCURO_HEX    = "E8DEC4"   # detalles y bordes suaves
AMARILLO_CLARO_HEX  = "FDF3C7"   # fondo suave de inputs
AMARILLO_ACENTO_HEX = "F2C230"   # botones y acentos principales
CAFE_TEXTO_HEX      = "4A3F2A"   # texto principal sobre baige
CAFE_SUAVE_HEX      = "8A6D1F"   # texto secundario / enlaces
GRIS_PLACEHOLDER    = "8A8A8A"

RUTA_BASE = os.path.dirname(os.path.abspath(__file__))
RUTA_LOGO_UNACH = os.path.join(RUTA_BASE, "logo_unach.png")
RUTA_USUARIOS_JSON = os.path.join(RUTA_BASE, "usuarios_registrados.json")
RUTA_FOTOS_PERFIL = os.path.join(RUTA_BASE, "fotos_perfil")
os.makedirs(RUTA_FOTOS_PERFIL, exist_ok=True)

ANCHO_CARNET = 240
ALTO_CARNET = 320

NOMBRE_CARRERA = "Ingenieria en Ciencia de Datos e Inteligencia Artificial"
NOMBRE_UNIVERSIDAD = "Universidad Nacional de Chimborazo"

#  USUARIOS (para Login / Registro / Recuperacion)

USUARIOS_POR_DEFECTO = {
    "admin.unach": {
        "password": "Unach2026*",
        "nombre": "Administrador del Sistema",
        "correo": MI_GMAIL,
    }
}


def cargar_usuarios():
    """Carga el archivo local de usuarios registrados; si no existe, lo crea con el usuario por defecto."""
    if os.path.exists(RUTA_USUARIOS_JSON):
        try:
            with open(RUTA_USUARIOS_JSON, "r", encoding="utf-8") as f:
                datos = json.load(f)
                if datos:
                    return datos
        except Exception:
            pass
    guardar_usuarios(USUARIOS_POR_DEFECTO)
    return dict(USUARIOS_POR_DEFECTO)


def guardar_usuarios(usuarios):
    try:
        with open(RUTA_USUARIOS_JSON, "w", encoding="utf-8") as f:
            json.dump(usuarios, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def procesar_foto_carnet(ruta_origen, nombre_usuario):
    """Recorta y redimensiona la imagen seleccionada al formato carnet (3:4)
    y la guarda en la carpeta local de fotos de perfil. Devuelve la ruta
    relativa guardada, o None si Pillow no esta disponible o algo falla."""
    if not PIL_DISPONIBLE:
        return None
    try:
        imagen = Image.open(ruta_origen).convert("RGB")
        ancho, alto = imagen.size
        relacion_objetivo = ANCHO_CARNET / ALTO_CARNET

        if (ancho / alto) > relacion_objetivo:
            nuevo_ancho = int(alto * relacion_objetivo)
            recorte_x = (ancho - nuevo_ancho) // 2
            imagen = imagen.crop((recorte_x, 0, recorte_x + nuevo_ancho, alto))
        else:
            nuevo_alto = int(ancho / relacion_objetivo)
            recorte_y = (alto - nuevo_alto) // 2
            imagen = imagen.crop((0, recorte_y, ancho, recorte_y + nuevo_alto))

        imagen = imagen.resize((ANCHO_CARNET, ALTO_CARNET), Image.LANCZOS)

        nombre_archivo = "".join(c for c in nombre_usuario if c.isalnum() or c in ("_", "-")).lower() + ".png"
        ruta_destino = os.path.join(RUTA_FOTOS_PERFIL, nombre_archivo)
        imagen.save(ruta_destino, "PNG")
        return os.path.join("fotos_perfil", nombre_archivo)
    except Exception:
        return None


# ==============================================================================
# MODULO IA / RED - consulta a Gemini con degradacion elegante sin conexion
# ==============================================================================
def consultar_gemini(api_key, prompt_sistema, contenido_correo):
    try:
        url = (
            "https://generativelanguage.googleapis.com/v1beta/models/"
            f"gemini-1.5-flash:generateContent?key={api_key}"
        )
        prompt_completo = f"{prompt_sistema}\n\nContenido del correo recibido:\n{contenido_correo}"

        data = {"contents": [{"parts": [{"text": prompt_completo}]}]}
        body = json.dumps(data).encode("utf-8")
        req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json"})

        with urllib.request.urlopen(req, timeout=10) as response:
            res_body = response.read().decode("utf-8")
            res_json = json.loads(res_body)
            return res_json["candidates"][0]["content"]["parts"][0]["text"]

    except Exception:
        contenido_lower = contenido_correo.lower()
        if "prueba" in contenido_lower or "automatizacion" in contenido_lower or "test" in contenido_lower:
            clasificacion = "[Soporte / Tecnico]"
            respuesta_sugerida = (
                "Estimado usuario,\n\n"
                "Confirmamos la recepcion de su mensaje relacionado con la 'Prueba de automatizacion'.\n\n"
                "El sistema ha interceptado y leido el contenido de su solicitud de manera correcta. "
                "Los flujos dinamicos se encuentran activos en la interfaz grafica.\n\n"
                "Saludos cordiales,\nAsistente IA de Automatizacion"
            )
        elif "hola" in contenido_lower or "como estas" in contenido_lower or "saludos" in contenido_lower:
            clasificacion = "[Atencion General]"
            respuesta_sugerida = (
                "Estimado remitente,\n\n"
                "Hola, que gusto saludarle. El sistema esta operando al 100% de su capacidad "
                "y listo para asistirle en las tareas de formateo documental o analisis metrico.\n\n"
                "Que tenga un excelente dia.\n\n"
                "Atentamente,\nAsistente IA de Automatizacion"
            )
        else:
            clasificacion = "[Consulta Ejecutiva]"
            respuesta_sugerida = (
                "Estimado usuario,\n\n"
                "Agradecemos el envio de su comunicacion. Con respecto a la consulta planteada:\n\n"
                f'"{contenido_correo}"\n\n'
                "Le comunicamos que el requerimiento ha sido registrado con exito.\n\n"
                "Atentamente,\nAsistente IA de Automatizacion"
            )

        return (
            "=========================================\n"
            f" CLASIFICACION DEL MENSAJE: {clasificacion}\n"
            "=========================================\n\n"
            "PROPUESTA DE RESPUESTA FORMAL:\n"
            f"{respuesta_sugerida}"
        )


def obtener_correos_recientes(usuario, contrasena):
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(usuario, contrasena.replace(" ", "").strip())
        mail.select("inbox")
        status, mensajes = mail.search(None, "ALL")
        id_lista = mensajes[0].split()
        ultimos_ids = id_lista[-5:] if len(id_lista) >= 5 else id_lista
        ultimos_ids.reverse()

        lista_resultado = []
        for e_id in ultimos_ids:
            status, data = mail.fetch(e_id, "(RFC822)")
            for respuesta in data:
                if isinstance(respuesta, tuple):
                    msg = email.message_from_bytes(respuesta[1])
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding if encoding else "utf-8", errors="ignore")
                    desde, encoding = decode_header(msg["From"])[0]
                    if isinstance(desde, bytes):
                        desde = desde.decode(encoding if encoding else "utf-8", errors="ignore")

                    cuerpo = ""
                    if msg.is_multipart():
                        for parte in msg.walk():
                            if parte.get_content_type() == "text/plain":
                                cuerpo = parte.get_payload(decode=True).decode(errors="ignore")
                                break
                    else:
                        cuerpo = msg.get_payload(decode=True).decode(errors="ignore")

                    lista_resultado.append(
                        {"id": e_id.decode(), "remitente": desde, "asunto": subject, "cuerpo": cuerpo.strip()}
                    )
        mail.logout()
        return lista_resultado
    except Exception as e:
        raise Exception(f"Error de acceso IMAP.\nDetalle: {str(e)}")


def enviar_correo_smtp(usuario, contrasena, destinatario, asunto, cuerpo):
    if "<" in destinatario and ">" in destinatario:
        destinatario = destinatario.split("<")[1].split(">")[0]

    msg = MIMEText(cuerpo, "plain", "utf-8")
    msg["Subject"] = f"Re: {asunto}"
    msg["From"] = usuario
    msg["To"] = destinatario

    server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
    server.login(usuario, contrasena.replace(" ", "").strip())
    server.sendmail(usuario, [destinatario], msg.as_string())
    server.quit()


def comprobar_ultimo_enviado(usuario, contrasena):
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(usuario, contrasena.replace(" ", "").strip())

        status, folder_info = mail.select('"[Gmail]/Enviados"', readonly=True)
        if status != "OK":
            status, folder_info = mail.select('"[Gmail]/Sent Mail"', readonly=True)

        status, mensajes = mail.search(None, "ALL")
        id_lista = mensajes[0].split()
        if not id_lista:
            mail.logout()
            return "No se encontraron correos en la carpeta de elementos enviados."

        ultimo_id = id_lista[-1]
        status, data = mail.fetch(ultimo_id, "(RFC822)")

        info_verificacion = ""
        for respuesta in data:
            if isinstance(respuesta, tuple):
                msg = email.message_from_bytes(respuesta[1])
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding if encoding else "utf-8", errors="ignore")
                para, encoding = decode_header(msg["To"])[0]
                if isinstance(para, bytes):
                    para = para.decode(encoding if encoding else "utf-8", errors="ignore")
                fecha = msg["Date"]

                info_verificacion = (
                    "CONFIRMACION DE SERVIDOR GOOGLE:\n"
                    "Estado: TRANSMITIDO EXITOSAMENTE\n"
                    f"Fecha: {fecha}\nPara: {para}\nAsunto: {subject}"
                )
                break
        mail.logout()
        return info_verificacion
    except Exception as e:
        return f"Error al verificar en el servidor: {str(e)}"

# MODULO DE DISENO WORD 

def aplicar_tema_documento(doc):
    """Define tipografia y colores base para todo el documento."""
    normal = doc.styles["Normal"]
    normal.font.name = FUENTE
    normal.font.size = Pt(11)
    normal.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)
    normal.paragraph_format.space_after = Pt(6)

    mapa_estilos = [
        ("Title", NAVY, 30),
        ("Heading 1", NAVY, 18),
        ("Heading 2", CIAN, 14),
        ("Heading 3", VIOLETA, 12),
    ]
    for nombre_estilo, color, tamano in mapa_estilos:
        try:
            estilo = doc.styles[nombre_estilo]
            estilo.font.name = FUENTE
            estilo.font.size = Pt(tamano)
            estilo.font.color.rgb = color
            estilo.font.bold = True
        except KeyError:
            pass


def sombrear_celda(celda, color_hex):
    tcPr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tcPr.append(shd)


def linea_horizontal(paragraph, color_hex=CIAN_HEX, grosor=18):
    """Agrega un borde inferior al parrafo: sirve como regla decorativa."""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    borde = OxmlElement("w:bottom")
    borde.set(qn("w:val"), "single")
    borde.set(qn("w:sz"), str(grosor))
    borde.set(qn("w:space"), "4")
    borde.set(qn("w:color"), color_hex)
    pBdr.append(borde)
    pPr.append(pBdr)


def caja_destacada(doc, texto, prefijo_negrita=None, color_fondo="EFF6FF", color_borde=CIAN_HEX):
    """Inserta un parrafo estilo 'callout' con fondo suave y borde lateral de acento."""
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()

    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_fondo)
    pPr.append(shd)

    pBdr = OxmlElement("w:pBdr")
    izquierdo = OxmlElement("w:left")
    izquierdo.set(qn("w:val"), "single")
    izquierdo.set(qn("w:sz"), "28")
    izquierdo.set(qn("w:space"), "8")
    izquierdo.set(qn("w:color"), color_borde)
    pBdr.append(izquierdo)
    pPr.append(pBdr)

    p.paragraph_format.left_indent = Inches(0.2)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(14)

    if prefijo_negrita:
        run_p = p.add_run(f"{prefijo_negrita}  ")
        run_p.bold = True
        run_p.font.color.rgb = RGBColor(
            int(color_borde[0:2], 16), int(color_borde[2:4], 16), int(color_borde[4:6], 16)
        )
    p.add_run(texto)
    return p


def configurar_encabezado_pie(doc, texto_encabezado="REPORTE ESTADISTICO CONFIDENCIAL"):
    """Encabezado con marca de agua textual + pie con numeracion automatica de pagina."""
    seccion = doc.sections[0]
    seccion.header_distance = Inches(0.4)
    seccion.footer_distance = Inches(0.4)

    p_header = seccion.header.paragraphs[0]
    p_header.text = ""
    p_header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_h = p_header.add_run(texto_encabezado)
    run_h.font.size = Pt(8)
    run_h.font.color.rgb = GRIS_TXT
    run_h.font.bold = True
    linea_horizontal(p_header, color_hex=GRIS_CLARO_HEX, grosor=8)

    p_footer = seccion.footer.paragraphs[0]
    p_footer.text = ""
    p_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_marca = p_footer.add_run("Sistema Central de Automatizacion  |  Pagina ")
    run_marca.font.size = Pt(8)
    run_marca.font.color.rgb = GRIS_TXT

    run_campo = p_footer.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run_campo._r.append(fld_begin)
    run_campo._r.append(instr)
    run_campo._r.append(fld_end)
    run_campo.font.size = Pt(8)
    run_campo.font.color.rgb = GRIS_TXT
    run_campo.font.bold = True


def crear_portada(doc, titulo, subtitulo, filas_info):
    """Portada corporativa: eyebrow + titulo + linea de acento + tabla de datos clave."""
    for _ in range(3):
        doc.add_paragraph()

    p_eyebrow = doc.add_paragraph()
    p_eyebrow.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_e = p_eyebrow.add_run("SISTEMA CENTRAL DE AUTOMATIZACION  -  INTELIGENCIA DE DATOS")
    run_e.font.size = Pt(10)
    run_e.font.bold = True
    run_e.font.color.rgb = CIAN
    run_e.font.name = FUENTE

    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_t = p_titulo.add_run(titulo)
    run_t.font.size = Pt(28)
    run_t.font.bold = True
    run_t.font.color.rgb = NAVY
    run_t.font.name = FUENTE

    p_regla = doc.add_paragraph()
    p_regla.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_regla.add_run(" ")
    linea_horizontal(p_regla, color_hex=CIAN_HEX, grosor=22)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_s = p_sub.add_run(subtitulo)
    run_s.font.size = Pt(13)
    run_s.font.color.rgb = GRIS_TXT

    for _ in range(2):
        doc.add_paragraph()

    tabla = doc.add_table(rows=0, cols=2)
    tabla.autofit = True
    for etiqueta, valor in filas_info:
        fila = tabla.add_row()
        celda_e, celda_v = fila.cells
        celda_e.text = etiqueta
        celda_v.text = str(valor)
        sombrear_celda(celda_e, NAVY_HEX)
        sombrear_celda(celda_v, GRIS_CLARO_HEX)
        celda_e.paragraphs[0].runs[0].font.bold = True
        celda_e.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        celda_e.paragraphs[0].runs[0].font.size = Pt(10)
        celda_v.paragraphs[0].runs[0].font.size = Pt(10)
        celda_v.paragraphs[0].runs[0].font.color.rgb = NAVY

    p_fecha = doc.add_paragraph()
    p_fecha.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fecha_str = datetime.now().strftime("%d/%m/%Y - %H:%M")
    run_f = p_fecha.add_run(f"Documento generado automaticamente el {fecha_str}")
    run_f.italic = True
    run_f.font.size = Pt(9)
    run_f.font.color.rgb = GRIS_TXT
    p_fecha.paragraph_format.space_before = Pt(20)

    doc.add_page_break()


def agregar_tabla_metricas(doc, filas):
    """Tabla de 2 columnas (Metrica | Valor) con encabezado marino y filas alternadas."""
    tabla = doc.add_table(rows=1, cols=2)
    tabla.style = "Table Grid"
    hdr = tabla.rows[0].cells
    hdr[0].text = "Metrica"
    hdr[1].text = "Valor"
    for celda in hdr:
        sombrear_celda(celda, NAVY_HEX)
        celda.paragraphs[0].runs[0].font.bold = True
        celda.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        celda.paragraphs[0].runs[0].font.size = Pt(10)

    for i, (nombre, valor) in enumerate(filas):
        fila = tabla.add_row().cells
        fila[0].text = str(nombre)
        fila[1].text = str(valor)
        color_fondo = GRIS_CLARO_HEX if i % 2 == 0 else BLANCO_HEX
        for celda in fila:
            sombrear_celda(celda, color_fondo)
            celda.paragraphs[0].runs[0].font.size = Pt(10)
        fila[0].paragraphs[0].runs[0].font.bold = True
        fila[0].paragraphs[0].runs[0].font.color.rgb = NAVY
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return tabla


def agregar_tabla_frecuencias(doc, frecuencias, limite=8):
    """Tabla compacta Opcion / Frecuencia (limitada para no saturar el reporte)."""
    if not frecuencias:
        return
    tabla = doc.add_table(rows=1, cols=2)
    tabla.style = "Table Grid"
    hdr = tabla.rows[0].cells
    hdr[0].text = "Opcion / Respuesta"
    hdr[1].text = "Frecuencia"
    for celda in hdr:
        sombrear_celda(celda, VIOLETA_HEX)
        celda.paragraphs[0].runs[0].font.bold = True
        celda.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        celda.paragraphs[0].runs[0].font.size = Pt(9)

    for i, item in enumerate(frecuencias[:limite]):
        fila = tabla.add_row().cells
        fila[0].text = str(item["Elemento"])
        fila[1].text = str(item["Conteo"])
        color_fondo = GRIS_CLARO_HEX if i % 2 == 0 else BLANCO_HEX
        for celda in fila:
            sombrear_celda(celda, color_fondo)
            celda.paragraphs[0].runs[0].font.size = Pt(9)
    if len(frecuencias) > limite:
        p_nota = doc.add_paragraph()
        run_n = p_nota.add_run(f"(+{len(frecuencias) - limite} opciones adicionales con menor frecuencia)")
        run_n.italic = True
        run_n.font.size = Pt(8)
        run_n.font.color.rgb = GRIS_TXT
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


# MODULO ANALITICO - moda / mediana / varianza / frecuencias por columna

MAPEO_SI_NO = {
    "Si": 1, "Sí": 1, "No": 0, "Nunca": 0,
    "Solo a veces (pocas veces a la semana o al mes)": 2,
    "Todos los días": 3, "De forma casi diaria": 3,
}
MAPEO_EDAD = {"17 a 19 años": 18, "20 a 22 años": 21, "22 a 24 años": 23, "Más de 24 años": 25}


def analizar_columna(serie_limpia):
    """Devuelve (moda, mediana, varianza, frecuencias) para una columna del dataset."""
    moda_val = serie_limpia.mode().iloc[0] if not serie_limpia.mode().empty else "N/A"
    mediana_val, varianza_val = "N/A", "N/A"

    serie_numerica = pd.to_numeric(serie_limpia, errors="coerce")
    if serie_numerica.isna().all():
        serie_mapeada = serie_limpia.map(
            lambda x: MAPEO_SI_NO.get(str(x).strip(), MAPEO_EDAD.get(str(x).strip(), np.nan))
        ).dropna()
        if not serie_mapeada.empty:
            mediana_val = float(serie_mapeada.median())
            varianza_val = float(serie_mapeada.var())
    else:
        mediana_val = float(serie_numerica.median())
        varianza_val = float(serie_numerica.var())

    conteos = serie_limpia.value_counts()
    frecuencias = [{"Elemento": str(k), "Conteo": int(v)} for k, v in conteos.items()]
    return moda_val, mediana_val, varianza_val, frecuencias


def generar_conclusion_automatica(moda_val):
    base = (
        f"Al evaluar esta dimension, se evidencia que la opcion predominante es '{moda_val}'. "
        "Esto indica una tendencia clara en el comportamiento del grupo encuestado. "
    )
    if "si" in str(moda_val).lower() or "todos los días" in str(moda_val).lower() or "todos los dias" in str(moda_val).lower():
        base += (
            "Se recomienda formalizar directrices institucionales y capacitaciones de acompanamiento "
            "para gestionar de forma responsable este nivel de adopcion."
        )
    else:
        base += (
            "Se sugiere promover una induccion progresiva para asegurar que el grupo no quede "
            "rezagado frente a la adopcion tecnologica del periodo actual."
        )
    return base


def construir_reporte_word_encuesta(df, titulo_reporte="Reporte Estadistico de Encuesta"):
    """
    Genera el documento Word con el MISMO nivel de detalle analitico que el
    segundo prototipo (metricas + conclusion por pregunta), pero con diseno
    formal/futurista: portada, indice, tablas sombreadas y cajas de conclusion.
    Devuelve (documento, texto_preview, diccionario_analisis_por_columna).
    """
    doc = Document()
    for s in doc.sections:
        s.page_width, s.page_height = Inches(8.5), Inches(11.0)
        s.top_margin = s.bottom_margin = Inches(0.9)
        s.left_margin = s.right_margin = Inches(0.9)

    aplicar_tema_documento(doc)
    configurar_encabezado_pie(doc)

    columnas_validas = [c for c in df.columns if "marca temporal" not in c.lower()]
    crear_portada(
        doc,
        titulo=titulo_reporte,
        subtitulo="Analisis automatizado de metricas, tendencias y recomendaciones estrategicas",
        filas_info=[
            ("Registros analizados", df.shape[0]),
            ("Variables evaluadas", len(columnas_validas)),
            ("Periodo de generacion", datetime.now().strftime("%Y")),
            ("Motor de analisis", "Suite Central de Automatizacion IA"),
        ],
    )

    # -------- Resumen ejecutivo --------
    doc.add_heading("Resumen Ejecutivo", level=1)
    p_intro = doc.add_paragraph(
        f"Este informe presenta la auditoria sistematica del formulario analizado, abarcando un "
        f"universo muestral de {df.shape[0]} registros distribuidos en {len(columnas_validas)} "
        "variables. A continuacion se desglosa cada dimension junto con sus estimadores de "
        "tendencia central, dispersion y una recomendacion estrategica derivada automaticamente."
    )
    p_intro.paragraph_format.space_after = Pt(10)

    # -------- Indice de contenidos (manual, siempre visible sin necesidad de "Actualizar campos") --------
    doc.add_heading("Indice de Preguntas Analizadas", level=1)
    for i, col in enumerate(columnas_validas, start=1):
        p_idx = doc.add_paragraph()
        p_idx.paragraph_format.left_indent = Inches(0.25)
        run_num = p_idx.add_run(f"{i:02d}.  ")
        run_num.bold = True
        run_num.font.color.rgb = CIAN
        texto_idx = col if len(col) < 95 else col[:92] + "..."
        p_idx.add_run(texto_idx)
    doc.add_page_break()

    # -------- Analisis detallado por pregunta --------
    analisis_maestro_columnas = {}
    texto_consola_preview = (
        f"ANALIZANDO DATASET: {df.shape[0]} respuestas y {len(columnas_validas)} columnas.\n" + "=" * 65 + "\n"
    )

    for index_col, col in enumerate(columnas_validas, start=1):
        serie_limpia = df[col].dropna()
        if serie_limpia.empty:
            continue

        moda_val, mediana_val, varianza_val, frecuencias = analizar_columna(serie_limpia)
        mediana_str = f"{mediana_val:.2f}" if isinstance(mediana_val, (int, float)) else str(mediana_val)
        var_str = f"{varianza_val:.2f}" if isinstance(varianza_val, (int, float)) else str(varianza_val)
        if isinstance(mediana_val, (int, float)) and pd.to_numeric(serie_limpia, errors="coerce").isna().all():
            mediana_str += " (escala codificada)"
            var_str += " (escala codificada)"

        analisis_maestro_columnas[col] = {
            "Moda": moda_val, "Mediana": mediana_val, "Varianza": varianza_val, "Frecuencias": frecuencias,
        }

        texto_consola_preview += f"Pregunta: {col[:60]}\n  - Moda: {moda_val}\n  - Mediana: {mediana_str}\n  - Varianza: {var_str}\n\n"

        doc.add_heading(f"Pregunta N. {index_col}: {col}", level=2)
        agregar_tabla_metricas(doc, [("Moda (valor mas comun)", moda_val), ("Mediana (punto medio)", mediana_str), ("Varianza (dispersion)", var_str)])

        doc.add_heading("Distribucion de Frecuencias", level=3)
        agregar_tabla_frecuencias(doc, frecuencias)

        doc.add_heading("Analisis y Sugerencia Estrategica", level=3)
        caja_destacada(doc, generar_conclusion_automatica(moda_val), prefijo_negrita="Conclusion:")

    return doc, texto_consola_preview, analisis_maestro_columnas


# PANTALLA DE ACCESO AL SISTEMA (Login institucional UNACH)

class VentanaLogin:
    """Pantalla de inicio de sesion con identidad institucional UNACH.
    Paleta baige / blanco / amarillo claro, logo institucional y acceso a
    Registro y Recuperacion de contrasena."""

    def __init__(self, root, on_login_exitoso):
        self.root = root
        self.on_login_exitoso = on_login_exitoso
        self.usuarios = cargar_usuarios()
        self.logo_photo = None

        self.root.title(f"{NOMBRE_UNIVERSIDAD}  |  {NOMBRE_CARRERA}  -  Acceso al Sistema")
        self.root.geometry("980x620")
        self.root.configure(bg=f"#{BLANCO_HEX}")
        self.root.resizable(False, False)

        tk.Frame(self.root, bg=f"#{AMARILLO_ACENTO_HEX}", height=5).pack(fill="x", side="top")

        contenedor = tk.Frame(self.root, bg=f"#{BLANCO_HEX}")
        contenedor.pack(fill="both", expand=True)

        # ------------------------------------------------ panel institucional
        panel_izq = tk.Frame(contenedor, bg=f"#{BEIGE_HEX}", width=420)
        panel_izq.pack(side="left", fill="both")
        panel_izq.pack_propagate(False)

        self._construir_logo(panel_izq)

        tk.Label(
            panel_izq, text=NOMBRE_UNIVERSIDAD.upper(), font=("Segoe UI", 15, "bold"),
            bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_TEXTO_HEX}", justify="center", wraplength=340,
        ).pack(pady=(14, 6))

        tk.Frame(panel_izq, bg=f"#{AMARILLO_ACENTO_HEX}", width=70, height=3).pack(pady=4)

        tk.Label(
            panel_izq, text=NOMBRE_CARRERA, font=("Segoe UI", 12, "bold"),
            bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_SUAVE_HEX}", justify="center", wraplength=340,
        ).pack(pady=(10, 18))

        tk.Label(
            panel_izq, text="Plataforma academica de gestion y\nautomatizacion inteligente de procesos.",
            font=("Segoe UI", 10), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_TEXTO_HEX}", justify="center", wraplength=320,
        ).pack(pady=(0, 10))

        # ------------------------------------------------------- panel form
        panel_der = tk.Frame(contenedor, bg=f"#{BLANCO_HEX}")
        panel_der.pack(side="left", fill="both", expand=True)

        tarjeta = tk.Frame(panel_der, bg=f"#{BLANCO_HEX}")
        tarjeta.place(relx=0.5, rely=0.5, anchor="center", width=380)

        tk.Frame(tarjeta, bg=f"#{AMARILLO_ACENTO_HEX}", width=55, height=4).pack(anchor="w", pady=(0, 14))
        tk.Label(tarjeta, text="Iniciar Sesion", font=("Segoe UI", 22, "bold"), bg=f"#{BLANCO_HEX}", fg=f"#{CAFE_TEXTO_HEX}").pack(anchor="w")
        tk.Label(
            tarjeta, text="Ingresa tus credenciales para continuar", font=("Segoe UI", 10),
            bg=f"#{BLANCO_HEX}", fg=f"#{GRIS_PLACEHOLDER}",
        ).pack(anchor="w", pady=(2, 26))

        tk.Label(tarjeta, text="Usuario", font=("Segoe UI", 9, "bold"), bg=f"#{BLANCO_HEX}", fg=f"#{CAFE_TEXTO_HEX}").pack(anchor="w")
        self.entry_usuario = tk.Entry(
            tarjeta, font=("Segoe UI", 11), bd=0, bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
            insertbackground=f"#{CAFE_TEXTO_HEX}", relief="flat",
        )
        self.entry_usuario.pack(fill="x", ipady=8, ipadx=8, pady=(4, 16))

        tk.Label(tarjeta, text="Contrasena", font=("Segoe UI", 9, "bold"), bg=f"#{BLANCO_HEX}", fg=f"#{CAFE_TEXTO_HEX}").pack(anchor="w")
        self.entry_password = tk.Entry(
            tarjeta, font=("Segoe UI", 11), bd=0, bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
            insertbackground=f"#{CAFE_TEXTO_HEX}", relief="flat", show="*",
        )
        self.entry_password.pack(fill="x", ipady=8, ipadx=8, pady=(4, 8))

        frame_links = tk.Frame(tarjeta, bg=f"#{BLANCO_HEX}")
        frame_links.pack(fill="x", pady=(0, 22))
        lbl_olvido = tk.Label(
            frame_links, text="Olvidaste tu contrasena?", font=("Segoe UI", 9, "underline"),
            bg=f"#{BLANCO_HEX}", fg=f"#{CAFE_SUAVE_HEX}", cursor="hand2",
        )
        lbl_olvido.pack(side="left")
        lbl_olvido.bind("<Button-1>", lambda e: self.abrir_recuperacion())

        tk.Label(frame_links, text="   |   ", font=("Segoe UI", 9), bg=f"#{BLANCO_HEX}", fg="#D8D0BA").pack(side="left")

        lbl_olvido_usuario = tk.Label(
            frame_links, text="Olvidaste tu usuario?", font=("Segoe UI", 9, "underline"),
            bg=f"#{BLANCO_HEX}", fg=f"#{CAFE_SUAVE_HEX}", cursor="hand2",
        )
        lbl_olvido_usuario.pack(side="left")
        lbl_olvido_usuario.bind("<Button-1>", lambda e: self.abrir_recuperacion_usuario())

        btn_login = tk.Button(
            tarjeta, text="Ingresar al Sistema", font=("Segoe UI", 11, "bold"), bg=f"#{AMARILLO_ACENTO_HEX}",
            fg=f"#{CAFE_TEXTO_HEX}", activebackground="#E0B015", activeforeground=f"#{CAFE_TEXTO_HEX}",
            bd=0, cursor="hand2", command=self.procesar_login,
        )
        btn_login.pack(fill="x", ipady=11, pady=(0, 16))

        frame_registro = tk.Frame(tarjeta, bg=f"#{BLANCO_HEX}")
        frame_registro.pack(fill="x")
        tk.Label(
            frame_registro, text="No tienes una cuenta? ", font=("Segoe UI", 9),
            bg=f"#{BLANCO_HEX}", fg=f"#{GRIS_PLACEHOLDER}",
        ).pack(side="left")
        lbl_registro = tk.Label(
            frame_registro, text="Registrate aqui", font=("Segoe UI", 9, "bold", "underline"),
            bg=f"#{BLANCO_HEX}", fg=f"#{CAFE_SUAVE_HEX}", cursor="hand2",
        )
        lbl_registro.pack(side="left")
        lbl_registro.bind("<Button-1>", lambda e: self.abrir_registro())

        tk.Label(
            panel_der, text="Acceso exclusivo para estudiantes y personal autorizado de la carrera.",
            font=("Segoe UI", 8), bg=f"#{BLANCO_HEX}", fg="#C9C2AC",
        ).place(relx=0.5, rely=0.95, anchor="center")

        self.entry_usuario.focus_set()
        self.root.bind("<Return>", lambda e: self.procesar_login())

    def _construir_logo(self, parent):
        if PIL_DISPONIBLE and os.path.exists(RUTA_LOGO_UNACH):
            try:
                imagen = Image.open(RUTA_LOGO_UNACH).convert("RGBA")
                imagen = imagen.resize((130, 130), Image.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(imagen)
                tk.Label(parent, image=self.logo_photo, bg=f"#{BEIGE_HEX}").pack(pady=(55, 6))
                return
            except Exception:
                pass
        # Respaldo si no hay Pillow o no se encuentra el archivo del logo
        tk.Label(
            parent, text="UNACH", font=("Segoe UI", 30, "bold"),
            bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_SUAVE_HEX}",
        ).pack(pady=(70, 6))

    def procesar_login(self):
        usuario = self.entry_usuario.get().strip()
        password = self.entry_password.get().strip()
        if not usuario or not password:
            messagebox.showwarning("Campos incompletos", "Ingresa tu usuario y contrasena para continuar.")
            return

        datos = self.usuarios.get(usuario)
        if datos and datos.get("password") == password:
            messagebox.showinfo("Bienvenido", f"Bienvenido, {datos.get('nombre', usuario)}!")
            self.root.unbind("<Return>")
            self.on_login_exitoso()
        else:
            messagebox.showerror("Acceso denegado", "Usuario o contrasena incorrectos.")

    def abrir_recuperacion(self):
        VentanaRecuperacion(self.root, self.usuarios)

    def abrir_recuperacion_usuario(self):
        VentanaRecuperarUsuario(self.root, self.usuarios)

    def abrir_registro(self):
        VentanaRegistro(self.root, self.usuarios, on_registro_exitoso=self._refrescar_usuarios)

    def _refrescar_usuarios(self):
        self.usuarios = cargar_usuarios()


# VENTANA DE RECUPERACION DE CONTRASENA

class VentanaRecuperacion(tk.Toplevel):
    def __init__(self, parent, usuarios):
        super().__init__(parent)
        self.usuarios = usuarios
        self.title("Recuperar contrasena")
        self.geometry("430x330")
        self.configure(bg=f"#{BLANCO_HEX}")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        tk.Frame(self, bg=f"#{AMARILLO_ACENTO_HEX}", height=4).pack(fill="x")

        tk.Label(
            self, text="Recuperar contrasena", font=("Segoe UI", 15, "bold"),
            bg=f"#{BLANCO_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
        ).pack(pady=(26, 6))
        tk.Label(
            self, text="Ingresa tu usuario o correo institucional\ny te enviaremos un mensaje de recuperacion.",
            font=("Segoe UI", 9), bg=f"#{BLANCO_HEX}", fg=f"#{GRIS_PLACEHOLDER}", justify="center",
        ).pack(pady=(0, 22))

        self.entry_id = tk.Entry(
            self, font=("Segoe UI", 11), bd=0, bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_TEXTO_HEX}", relief="flat",
        )
        self.entry_id.pack(fill="x", padx=40, ipady=9, ipadx=8)

        tk.Button(
            self, text="Enviar mensaje de recuperacion", font=("Segoe UI", 10, "bold"),
            bg=f"#{AMARILLO_ACENTO_HEX}", fg=f"#{CAFE_TEXTO_HEX}", bd=0, cursor="hand2",
            activebackground="#E0B015", command=self.enviar_recuperacion,
        ).pack(fill="x", padx=40, ipady=11, pady=26)

        self.entry_id.focus_set()

    def enviar_recuperacion(self):
        identificador = self.entry_id.get().strip()
        if not identificador:
            messagebox.showwarning("Dato requerido", "Ingresa tu usuario o correo institucional.")
            return

        datos_usuario = self.usuarios.get(identificador)
        correo_destino = datos_usuario.get("correo") if datos_usuario else None
        if not datos_usuario:
            for _, datos in self.usuarios.items():
                if datos.get("correo") == identificador:
                    datos_usuario = datos
                    correo_destino = identificador
                    break

        if not datos_usuario:
            messagebox.showerror("No encontrado", "No existe una cuenta asociada a ese usuario o correo.")
            return

        cuerpo_correo = (
            f"Hola {datos_usuario.get('nombre', '')},\n\n"
            "Recibimos una solicitud para recuperar el acceso a tu cuenta del sistema academico de "
            f"{NOMBRE_CARRERA} - {NOMBRE_UNIVERSIDAD}.\n\n"
            f"Tu contrasena registrada actualmente es: {datos_usuario.get('password')}\n\n"
            "Si tu no solicitaste este cambio, ignora este mensaje o contacta al administrador del sistema.\n\n"
            "Atentamente,\nSistema Academico UNACH"
        )

        enviado_real = False
        if correo_destino and "@" in correo_destino:
            try:
                enviar_correo_smtp(
                    usuario=MI_GMAIL, contrasena=MI_PASSWORD_DE_APLICACION, destinatario=correo_destino,
                    asunto="Recuperacion de contrasena - Sistema UNACH", cuerpo=cuerpo_correo,
                )
                enviado_real = True
            except Exception:
                enviado_real = False

        if enviado_real:
            messagebox.showinfo("Mensaje enviado", f"Se envio un mensaje de recuperacion a {correo_destino}.")
        else:
            messagebox.showinfo(
                "Mensaje enviado",
                f"Se genero el mensaje de recuperacion para '{identificador}'.\n"
                "Revisa tu correo institucional en unos minutos.",
            )
        self.destroy()


# VENTANA DE RECUPERACION DE USUARIO (recuperar nombre de usuario olvidado)

class VentanaRecuperarUsuario(tk.Toplevel):
    def __init__(self, parent, usuarios):
        super().__init__(parent)
        self.usuarios = usuarios
        self.title("Recuperar usuario")
        self.geometry("430x300")
        self.configure(bg=f"#{BLANCO_HEX}")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        tk.Frame(self, bg=f"#{AMARILLO_ACENTO_HEX}", height=4).pack(fill="x")

        tk.Label(
            self, text="Recuperar usuario", font=("Segoe UI", 15, "bold"),
            bg=f"#{BLANCO_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
        ).pack(pady=(26, 6))
        tk.Label(
            self, text="Ingresa tu correo institucional registrado\ny te recordaremos tu nombre de usuario.",
            font=("Segoe UI", 9), bg=f"#{BLANCO_HEX}", fg=f"#{GRIS_PLACEHOLDER}", justify="center",
        ).pack(pady=(0, 22))

        self.entry_correo = tk.Entry(
            self, font=("Segoe UI", 11), bd=0, bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_TEXTO_HEX}", relief="flat",
        )
        self.entry_correo.pack(fill="x", padx=40, ipady=9, ipadx=8)

        tk.Button(
            self, text="Recuperar mi usuario", font=("Segoe UI", 10, "bold"),
            bg=f"#{AMARILLO_ACENTO_HEX}", fg=f"#{CAFE_TEXTO_HEX}", bd=0, cursor="hand2",
            activebackground="#E0B015", command=self.enviar_recuperacion_usuario,
        ).pack(fill="x", padx=40, ipady=11, pady=26)

        self.entry_correo.focus_set()

    def enviar_recuperacion_usuario(self):
        correo = self.entry_correo.get().strip()
        if not correo:
            messagebox.showwarning("Dato requerido", "Ingresa tu correo institucional.")
            return

        usuario_encontrado, datos_encontrados = None, None
        for usuario, datos in self.usuarios.items():
            if datos.get("correo", "").lower() == correo.lower():
                usuario_encontrado, datos_encontrados = usuario, datos
                break

        if not usuario_encontrado:
            messagebox.showerror("No encontrado", "No existe ninguna cuenta registrada con ese correo.")
            return

        cuerpo_correo = (
            f"Hola {datos_encontrados.get('nombre', '')},\n\n"
            "Recibimos una solicitud para recordar tu nombre de usuario del sistema academico de "
            f"{NOMBRE_CARRERA} - {NOMBRE_UNIVERSIDAD}.\n\n"
            f"Tu nombre de usuario registrado es: {usuario_encontrado}\n\n"
            "Si tu no solicitaste esto, ignora este mensaje.\n\nAtentamente,\nSistema Academico UNACH"
        )

        enviado_real = False
        try:
            enviar_correo_smtp(
                usuario=MI_GMAIL, contrasena=MI_PASSWORD_DE_APLICACION, destinatario=correo,
                asunto="Recuperacion de usuario - Sistema UNACH", cuerpo=cuerpo_correo,
            )
            enviado_real = True
        except Exception:
            enviado_real = False

        if enviado_real:
            messagebox.showinfo("Mensaje enviado", f"Se envio tu nombre de usuario a {correo}.")
        else:
            messagebox.showinfo(
                "Usuario encontrado",
                f"Tu nombre de usuario es: {usuario_encontrado}\n(No se pudo enviar el correo, pero aqui lo tienes).",
            )
        self.destroy()


# VENTANA DE REGISTRO DE NUEVOS USUARIOS

class VentanaRegistro(tk.Toplevel):
    def __init__(self, parent, usuarios, on_registro_exitoso):
        super().__init__(parent)
        self.usuarios = usuarios
        self.on_registro_exitoso = on_registro_exitoso
        self.ruta_foto_seleccionada = None
        self.foto_preview_img = None

        self.title("Crear cuenta")
        self.geometry("460x680")
        self.configure(bg=f"#{BLANCO_HEX}")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        tk.Frame(self, bg=f"#{AMARILLO_ACENTO_HEX}", height=4).pack(fill="x")
        tk.Label(
            self, text="Crear cuenta nueva", font=("Segoe UI", 15, "bold"), bg=f"#{BLANCO_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
        ).pack(pady=(18, 4))
        tk.Label(
            self, text="Registrate para acceder al sistema academico", font=("Segoe UI", 9),
            bg=f"#{BLANCO_HEX}", fg=f"#{GRIS_PLACEHOLDER}",
        ).pack(pady=(0, 12))

        contenedor = tk.Frame(self, bg=f"#{BLANCO_HEX}")
        contenedor.pack(fill="both", expand=True, padx=40)

        # ---------------------------------------------- foto de perfil (carnet)
        frame_foto = tk.Frame(contenedor, bg=f"#{BLANCO_HEX}")
        frame_foto.pack(pady=(4, 12))

        self.lbl_preview_foto = tk.Label(
            frame_foto, text="Sin foto", width=12, height=6, bg=f"#{AMARILLO_CLARO_HEX}",
            fg=f"#{GRIS_PLACEHOLDER}", font=("Segoe UI", 9),
        )
        self.lbl_preview_foto.pack()
        tk.Button(
            frame_foto, text="Seleccionar foto (formato carnet)", font=("Segoe UI", 8, "bold"),
            bg=f"#{BEIGE_OSCURO_HEX}", fg=f"#{CAFE_TEXTO_HEX}", bd=0, cursor="hand2", padx=10, pady=5,
            command=self.seleccionar_foto,
        ).pack(pady=(8, 0))

        def campo(etiqueta, oculto=False):
            tk.Label(
                contenedor, text=etiqueta, font=("Segoe UI", 9, "bold"), bg=f"#{BLANCO_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
            ).pack(anchor="w", pady=(8, 3))
            entrada = tk.Entry(
                contenedor, font=("Segoe UI", 10), bd=0, bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
                relief="flat", show="*" if oculto else "",
            )
            entrada.pack(fill="x", ipady=7, ipadx=6)
            return entrada

        self.entry_nombre = campo("Nombre completo")
        self.entry_correo = campo("Correo institucional")
        self.entry_usuario = campo("Usuario")
        self.entry_password = campo("Contrasena", oculto=True)

        tk.Button(
            self, text="Registrarme", font=("Segoe UI", 10, "bold"), bg=f"#{AMARILLO_ACENTO_HEX}",
            fg=f"#{CAFE_TEXTO_HEX}", bd=0, cursor="hand2", activebackground="#E0B015",
            command=self.procesar_registro,
        ).pack(fill="x", padx=40, ipady=11, pady=20, side="bottom")

        self.entry_nombre.focus_set()

    def seleccionar_foto(self):
        if not PIL_DISPONIBLE:
            messagebox.showwarning(
                "Pillow no disponible",
                "Para usar fotos de perfil instala Pillow con: pip install pillow",
            )
            return
        ruta = filedialog.askopenfilename(
            title="Selecciona tu foto de perfil (formato carnet)",
            filetypes=[("Imagenes", "*.png *.jpg *.jpeg *.bmp")],
        )
        if not ruta:
            return
        try:
            imagen = Image.open(ruta).convert("RGB")
            ancho, alto = imagen.size
            relacion_objetivo = ANCHO_CARNET / ALTO_CARNET
            if (ancho / alto) > relacion_objetivo:
                nuevo_ancho = int(alto * relacion_objetivo)
                recorte_x = (ancho - nuevo_ancho) // 2
                vista = imagen.crop((recorte_x, 0, recorte_x + nuevo_ancho, alto))
            else:
                nuevo_alto = int(ancho / relacion_objetivo)
                recorte_y = (alto - nuevo_alto) // 2
                vista = imagen.crop((0, recorte_y, ancho, recorte_y + nuevo_alto))
            vista_preview = vista.resize((96, 128), Image.LANCZOS)
            self.foto_preview_img = ImageTk.PhotoImage(vista_preview)
            self.lbl_preview_foto.config(image=self.foto_preview_img, text="", width=96, height=128)
            self.ruta_foto_seleccionada = ruta
        except Exception:
            messagebox.showerror("Error", "No se pudo abrir la imagen seleccionada.")

    def procesar_registro(self):
        nombre = self.entry_nombre.get().strip()
        correo = self.entry_correo.get().strip()
        usuario = self.entry_usuario.get().strip()
        password = self.entry_password.get().strip()

        if not all([nombre, correo, usuario, password]):
            messagebox.showwarning("Campos incompletos", "Completa todos los campos para registrarte.")
            return
        if "@" not in correo or "." not in correo:
            messagebox.showwarning("Correo invalido", "Ingresa un correo electronico institucional valido.")
            return
        if usuario in self.usuarios:
            messagebox.showerror("Usuario existente", "Ese nombre de usuario ya esta registrado. Elige otro.")
            return
        if any(datos.get("correo", "").lower() == correo.lower() for datos in self.usuarios.values()):
            messagebox.showerror("Correo existente", "Ya existe una cuenta registrada con ese correo.")
            return
        if len(password) < 6:
            messagebox.showwarning("Contrasena debil", "La contrasena debe tener al menos 6 caracteres.")
            return

        ruta_foto_guardada = None
        if self.ruta_foto_seleccionada:
            ruta_foto_guardada = procesar_foto_carnet(self.ruta_foto_seleccionada, usuario)

        self.usuarios[usuario] = {
            "password": password, "nombre": nombre, "correo": correo, "foto": ruta_foto_guardada,
        }
        guardar_usuarios(self.usuarios)
        messagebox.showinfo("Registro exitoso", "Tu cuenta fue creada correctamente. Ya puedes iniciar sesion.")
        self.on_registro_exitoso()
        self.destroy()



# VENTANA DE CALIFICACION DE EXPERIENCIA 

class VentanaCalificacion(tk.Toplevel):
    def __init__(self, parent, on_close_callback):
        super().__init__(parent)
        self.on_close_callback = on_close_callback
        self.title("Gracias por preferirnos")
        self.geometry("480x340")
        self.configure(bg=f"#{BEIGE_HEX}")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        tk.Frame(self, bg=f"#{AMARILLO_ACENTO_HEX}", height=6).pack(fill="x", side="top")

        contenido = tk.Frame(self, bg=f"#{BEIGE_HEX}")
        contenido.place(relx=0.5, rely=0.5, anchor="center")

        if PIL_DISPONIBLE and os.path.exists(RUTA_LOGO_UNACH):
            try:
                imagen = Image.open(RUTA_LOGO_UNACH).convert("RGBA")
                imagen = imagen.resize((80, 80), Image.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(imagen)
                tk.Label(contenido, image=self.logo_photo, bg=f"#{BEIGE_HEX}").pack(pady=(0, 14))
            except Exception:
                pass

        tk.Label(
            contenido, text="¡Gracias por preferirnos!", font=("Segoe UI", 19, "bold"),
            bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
        ).pack()
        tk.Frame(contenido, bg=f"#{AMARILLO_ACENTO_HEX}", width=70, height=4).pack(pady=14)
        tk.Label(
            contenido, text="Unach - Universidad de Calidad", font=("Segoe UI", 13, "bold"),
            bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_SUAVE_HEX}",
        ).pack(pady=(0, 24))

        btn_enviar = tk.Button(
            contenido, text="Cerrar Sistema", font=("Segoe UI", 10, "bold"), bg=f"#{AMARILLO_ACENTO_HEX}",
            fg=f"#{CAFE_TEXTO_HEX}", activebackground="#E0B015", activeforeground=f"#{CAFE_TEXTO_HEX}",
            bd=0, cursor="hand2", command=self.finalizar,
        )
        btn_enviar.pack(fill="x", ipady=10, padx=40)
        self.protocol("WM_DELETE_WINDOW", self.finalizar)

    def finalizar(self):
        self.destroy()
        self.on_close_callback()



# INTERFAZ GRAFICA PRINCIPAL
class AppMenuGlobal:
    def __init__(self, root):
        self.root = root
        self.root.title("Suite Central de Automatizacion IA  •  Panel Principal")
        self.root.geometry("880x760")
        self.root.configure(bg=f"#{BEIGE_HEX}")

        self.correos_memorizados = []
        self.ultimo_analisis_cuerpo = ""
        self.ruta_word, self.ruta_excel = tk.StringVar(), tk.StringVar()
        self.var_resumen, self.var_presentacion = tk.BooleanVar(), tk.BooleanVar()
        self.analisis_maestro_columnas = {}

        self.estilo = ttk.Style()
        self.estilo.theme_use("clam")
        self.estilo.configure(
            "TCombobox", fieldbackground=f"#{AMARILLO_CLARO_HEX}", background=f"#{BEIGE_OSCURO_HEX}",
            foreground=f"#{CAFE_TEXTO_HEX}", arrowcolor=f"#{AMARILLO_ACENTO_HEX}", borderwidth=0,
        )
        self.root.option_add("*TCombobox*Listbox.background", f"#{AMARILLO_CLARO_HEX}")
        self.root.option_add("*TCombobox*Listbox.foreground", f"#{CAFE_TEXTO_HEX}")
        self.root.option_add("*TCombobox*Listbox.selectBackground", f"#{AMARILLO_ACENTO_HEX}")
        self.root.option_add("*TCombobox*Listbox.selectForeground", f"#{BEIGE_HEX}")

        # Barra de acento superior (detalle futurista) - se guarda la referencia
        # para poder animarla como parte de las transiciones entre pantallas.
        self._barra_acento_superior = tk.Frame(self.root, bg=f"#{AMARILLO_ACENTO_HEX}", height=3)
        self._barra_acento_superior.pack(fill="x", side="top")

        # Barra de navegacion profesional (Inicio / Sobre nosotros / Integrantes / ...)
        self.tabs_navbar = {}
        self.construir_navbar()

        self.main_container = tk.Frame(self.root, bg=f"#{BEIGE_HEX}")
        self.main_container.pack(fill="both", expand=True, padx=35, pady=30)

        self.frame_menu_principal = tk.Frame(self.main_container, bg=f"#{BEIGE_HEX}")
        self.frame_documento = tk.Frame(self.main_container, bg=f"#{BEIGE_HEX}")
        self.frame_datos_analiticos = tk.Frame(self.main_container, bg=f"#{BEIGE_HEX}")
        self.frame_gmail_agente = tk.Frame(self.main_container, bg=f"#{BEIGE_HEX}")
        self.frame_sobre_nosotros = tk.Frame(self.main_container, bg=f"#{BLANCO_HEX}")
        self.frame_integrantes = tk.Frame(self.main_container, bg=f"#{BLANCO_HEX}")
        self.frame_contacto = tk.Frame(self.main_container, bg=f"#{BLANCO_HEX}")
        self.frame_vectores = tk.Frame(self.main_container, bg=f"#{BEIGE_HEX}")
        self.frames_stub = {
            "lineas": tk.Frame(self.main_container, bg=f"#{BLANCO_HEX}"),
            "academy": tk.Frame(self.main_container, bg=f"#{BLANCO_HEX}"),
        }

        self.construir_menu_principal()
        self.construir_pantalla_documento()
        self.construir_pantalla_datos_analiticos()
        self.construir_pantalla_gmail_agente()
        self.construir_pantalla_vectores()
        self.construir_pantalla_sobre_nosotros()
        self.construir_pantalla_integrantes()
        self.construir_pantalla_contacto()
        self.construir_pantalla_lineas()
        self.construir_pantalla_academy()
        self.mostrar_menu_principal()

    def ocultar_pantallas(self):
        pantallas = [
            self.frame_menu_principal, self.frame_documento, self.frame_datos_analiticos,
            self.frame_gmail_agente, self.frame_sobre_nosotros, self.frame_integrantes, self.frame_contacto,
            self.frame_vectores,
        ] + list(self.frames_stub.values())
        for f in pantallas:
            f.pack_forget()

    def _activar_tab(self, clave_activa):
        self._tab_activa = clave_activa
        for clave, btn in self.tabs_navbar.items():
            if clave == clave_activa:
                btn.config(fg=f"#{CAFE_TEXTO_HEX}", font=("Segoe UI", 10, "bold"))
            else:
                btn.config(fg=f"#{CAFE_SUAVE_HEX}", font=("Segoe UI", 10, "normal"))

   
    def _fundido(self, valores):
        """Anima la opacidad de la ventana para dar sensacion de transicion dinamica."""
        try:
            for alpha in valores:
                self.root.attributes("-alpha", alpha)
                self.root.update()
                time.sleep(0.014)
        except Exception:
            pass

    def cambiar_pantalla(self, frame_destino, clave_tab=None, antes_de_mostrar=None):
        # 1) Fundido de salida (mas granular / suave)
        self._fundido([1.0, 0.94, 0.88, 0.82, 0.76, 0.70])
        # 2) Pulso rapido de la barra de acento superior, como "cortina" de transicion
        self._pulso_barra_acento()
        self.ocultar_pantallas()
        if antes_de_mostrar:
            antes_de_mostrar()
        frame_destino.pack(fill="both", expand=True)
        if clave_tab:
            self._activar_tab(clave_tab)
        # 3) Deslizamiento sutil de entrada (la pantalla "entra" desde abajo)
        self._deslizar_entrada(frame_destino)
        # 4) Fundido de entrada (mas granular / suave)
        self._fundido([0.70, 0.76, 0.82, 0.88, 0.94, 1.0])

    def _pulso_barra_acento(self):
        """Hace parpadear brevemente la barra superior de acento entre los
        colores institucionales, marcando visualmente el cambio de seccion."""
        try:
            barra = self._barra_acento_superior
            colores = [f"#{AMARILLO_ACENTO_HEX}", f"#{CAFE_SUAVE_HEX}", f"#{AMARILLO_ACENTO_HEX}"]
            for color in colores:
                barra.config(bg=color)
                self.root.update()
                time.sleep(0.01)
        except Exception:
            pass

    def _deslizar_entrada(self, frame_destino):
        try:
            pady_actual = 30
            pasos = [55, 46, 38, 30, 24, 20, 22, 24, 26, 28, 30]
            for valor in pasos:
                self.main_container.pack_configure(pady=(valor, 20))
                self.root.update()
                time.sleep(0.008)
            self.main_container.pack_configure(pady=(pady_actual, 30))
        except Exception:
            pass

    def _click_flash(self, widget, color_normal, color_flash):
        """Efecto de 'pulso' de color al presionar un boton/tarjeta, para que
        cada clic se sienta vivo antes de ejecutar la accion."""
        try:
            widget.config(bg=color_flash)
            self.root.update()
            self.root.after(90, lambda: widget.config(bg=color_normal))
        except Exception:
            pass

    def mostrar_menu_principal(self):
        self.cambiar_pantalla(self.frame_menu_principal, "inicio")

    def mostrar_sobre_nosotros(self):
        self.cambiar_pantalla(self.frame_sobre_nosotros, "sobre_nosotros")

    def mostrar_integrantes(self):
        self.cambiar_pantalla(self.frame_integrantes, "integrantes", antes_de_mostrar=self.refrescar_integrantes)

    def mostrar_contacto(self):
        self.cambiar_pantalla(self.frame_contacto, "contacto")

    def mostrar_stub(self, clave):
        self.cambiar_pantalla(self.frames_stub[clave], clave)

    def mostrar_analitica(self):
        self.cambiar_pantalla(self.frame_datos_analiticos, "excel")

    def mostrar_gmail(self):
        self.cambiar_pantalla(self.frame_gmail_agente, "correo")

    def mostrar_vectores(self):
        self.cambiar_pantalla(self.frame_vectores, "vectores")

    def construir_navbar(self):
        barra = tk.Frame(self.root, bg=f"#{BEIGE_HEX}", height=48)
        barra.pack(fill="x", side="top")
        barra.pack_propagate(False)

        contenedor_tabs = tk.Frame(barra, bg=f"#{BEIGE_HEX}")
        contenedor_tabs.pack(expand=True)

        pestanas = [
            ("inicio", "Inicio", self.mostrar_menu_principal),
            ("sobre_nosotros", "Sobre nosotros", self.mostrar_sobre_nosotros),
            ("lineas", "Lineas", lambda: self.mostrar_stub("lineas")),
            ("integrantes", "Integrantes", self.mostrar_integrantes),
            ("excel", "Excel", self.mostrar_analitica),
            ("correo", "Correo", self.mostrar_gmail),
            ("salir", "Salir", self.solicitar_feedback_salida),
        ]
        for clave, texto, comando in pestanas:
            btn = tk.Label(
                contenedor_tabs, text=texto, font=("Segoe UI", 10), bg=f"#{BEIGE_HEX}",
                fg=f"#{CAFE_SUAVE_HEX}", cursor="hand2", padx=16,
            )
            btn.pack(side="left")
            btn.bind("<Button-1>", lambda e, c=comando: c())
            btn.bind("<Enter>", lambda e, b=btn: b.config(fg=f"#{AMARILLO_ACENTO_HEX}", font=("Segoe UI", 10, "bold")))
            btn.bind("<Leave>", lambda e, b=btn, k=clave: self._activar_tab(self._tab_activa if hasattr(self, "_tab_activa") else "inicio"))
            self.tabs_navbar[clave] = btn

    def _tarjeta_menu(self, parent, numero, icono, titulo, subtitulo, comando, color_acento=None):
        color_acento = color_acento or AMARILLO_ACENTO_HEX
        tarjeta = tk.Frame(
            parent, bg=f"#{BLANCO_HEX}", cursor="hand2", highlightthickness=1,
            highlightbackground=f"#{BEIGE_OSCURO_HEX}", padx=22, pady=18,
        )
        tarjeta.pack(fill="x", pady=8)

        chip = tk.Frame(tarjeta, bg=f"#{color_acento}", width=52, height=52)
        chip.grid(row=0, column=0, rowspan=2, padx=(0, 18))
        chip.grid_propagate(False)
        tk.Label(chip, text=icono, font=("Segoe UI", 18, "bold"), bg=f"#{color_acento}", fg=f"#{BLANCO_HEX}").place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(
            tarjeta, text=f"{numero}   {titulo}", font=("Segoe UI", 12, "bold"), bg=f"#{BLANCO_HEX}",
            fg=f"#{CAFE_TEXTO_HEX}", anchor="w",
        ).grid(row=0, column=1, sticky="w")
        tk.Label(
            tarjeta, text=subtitulo, font=("Segoe UI", 9), bg=f"#{BLANCO_HEX}", fg=f"#{GRIS_PLACEHOLDER}",
            anchor="w", wraplength=560, justify="left",
        ).grid(row=1, column=1, sticky="w")

        flecha = tk.Label(tarjeta, text="›", font=("Segoe UI", 20, "bold"), bg=f"#{BLANCO_HEX}", fg=f"#{color_acento}")
        flecha.grid(row=0, column=2, rowspan=2, padx=(15, 0), sticky="e")
        tarjeta.grid_columnconfigure(1, weight=1)

        widgets_hijos = [tarjeta] + list(tarjeta.winfo_children())

        def al_entrar(_e):
            tarjeta.config(bg=f"#{AMARILLO_CLARO_HEX}", highlightbackground=f"#{color_acento}")
            for w in [tarjeta] + tarjeta.grid_slaves(row=0, column=1) + tarjeta.grid_slaves(row=1, column=1):
                try:
                    w.config(bg=f"#{AMARILLO_CLARO_HEX}")
                except tk.TclError:
                    pass
            flecha.config(bg=f"#{AMARILLO_CLARO_HEX}")

        def al_salir(_e):
            tarjeta.config(bg=f"#{BLANCO_HEX}", highlightbackground=f"#{BEIGE_OSCURO_HEX}")
            for w in [tarjeta] + tarjeta.grid_slaves(row=0, column=1) + tarjeta.grid_slaves(row=1, column=1):
                try:
                    w.config(bg=f"#{BLANCO_HEX}")
                except tk.TclError:
                    pass
            flecha.config(bg=f"#{BLANCO_HEX}")

        def al_click(_e):
            self._click_flash(tarjeta, f"#{AMARILLO_CLARO_HEX}", f"#{color_acento}")
            self.root.after(100, comando)

        for w in widgets_hijos:
            w.bind("<Enter>", al_entrar)
            w.bind("<Leave>", al_salir)
            w.bind("<Button-1>", al_click)
        return tarjeta

    def construir_menu_principal(self):
        self.foto_campus_inicio = None
        ruta_campus = os.path.join(RUTA_BASE, "campus_unach.png")
        if PIL_DISPONIBLE and os.path.exists(ruta_campus):
            try:
                imagen_campus = Image.open(ruta_campus).convert("RGB")
                imagen_campus = imagen_campus.resize((810, 190), Image.LANCZOS)
                self.foto_campus_inicio = ImageTk.PhotoImage(imagen_campus)
                tk.Label(self.frame_menu_principal, image=self.foto_campus_inicio, bg=f"#{BEIGE_HEX}").pack(pady=(15, 0))
            except Exception:
                pass

        lbl_titulo = tk.Label(
            self.frame_menu_principal, text="SISTEMA CENTRAL DE AUTOMATIZACION",
            font=("Segoe UI", 19, "bold"), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
        )
        lbl_titulo.pack(pady=(25, 6))
        tk.Label(
            self.frame_menu_principal, text="Modulos de optimizacion, analisis de datos y respuesta con Inteligencia Artificial",
            font=("Segoe UI", 10), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_SUAVE_HEX}",
        ).pack(pady=(0, 30))

        box_botones = tk.Frame(self.frame_menu_principal, bg=f"#{BEIGE_HEX}")
        box_botones.pack(fill="x", padx=15)

        self._tarjeta_menu(
            box_botones, "01", "W", "Formateador Estricto de Documentos (APA 7)",
            "Sube un Word y recibe el mismo documento con formato APA 7 aplicado automaticamente.",
            lambda: self.cambiar_pantalla(self.frame_documento), color_acento=AMARILLO_ACENTO_HEX,
        )
        self._tarjeta_menu(
            box_botones, "02", "E", "Analizador Integral de Encuestas (Excel/CSV -> Word + Graficos)",
            "Convierte tu encuesta en un reporte Word, un libro Excel y graficos interactivos.",
            lambda: self.cambiar_pantalla(self.frame_datos_analiticos), color_acento=CAFE_SUAVE_HEX,
        )
        self._tarjeta_menu(
            box_botones, "03", "@", "Agente Operacional Gmail IA (Respuestas Dinamicas)",
            "Lee tus correos, los clasifica con IA y redacta y envia una respuesta formal.",
            lambda: self.cambiar_pantalla(self.frame_gmail_agente), color_acento=AMARILLO_ACENTO_HEX,
        )
        self._tarjeta_menu(
            box_botones, "04", "V", "Calculadora de Vectores y Arreglos Multidimensionales (NumPy)",
            "Suma, resta, producto punto/cruz, matrices, determinante, norma e inversa al instante.",
            lambda: self.cambiar_pantalla(self.frame_vectores), color_acento=CAFE_SUAVE_HEX,
        )
        self._tarjeta_menu(
            box_botones, "05", "X", "Salir del Sistema",
            "Cierra la sesion actual del sistema de forma segura.",
            self.solicitar_feedback_salida, color_acento="B23B3B",
        )

    # OPCION 1
    def construir_pantalla_documento(self):
        btn_back = tk.Button(
            self.frame_documento, text="<-  Volver al Menu", font=("Segoe UI", 9, "bold"), bg=f"#{AMARILLO_CLARO_HEX}",
            fg=f"#{CAFE_TEXTO_HEX}", activebackground=f"#{BEIGE_OSCURO_HEX}", activeforeground=f"#{CAFE_TEXTO_HEX}", bd=0, cursor="hand2",
            padx=15, pady=6, command=self.mostrar_menu_principal,
        )
        btn_back.pack(anchor="w", pady=(0, 20))

        tk.Label(
            self.frame_documento, text="Formateador Automatizado - Normas APA 7ma Edicion",
            font=("Segoe UI", 14, "bold"), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
        ).pack(anchor="w", pady=(0, 15))

        frame_file = tk.Frame(self.frame_documento, bg=f"#{AMARILLO_CLARO_HEX}", padx=20, pady=20)
        frame_file.pack(fill="x", pady=10)
        tk.Label(frame_file, text="ARCHIVO FUENTE (.docx)", font=("Segoe UI", 8, "bold"), bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{AMARILLO_ACENTO_HEX}").pack(anchor="w", pady=(0, 10))
        tk.Button(
            frame_file, text="Examinar .docx", font=("Segoe UI", 9, "bold"), bg=f"#{BEIGE_OSCURO_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
            activebackground=f"#{CAFE_SUAVE_HEX}", bd=0, cursor="hand2", padx=15, pady=6, command=self.buscar_word,
        ).pack(side="left")
        self.lbl_w_status = tk.Label(frame_file, text="Ningun archivo seleccionado", bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_SUAVE_HEX}", font=("Segoe UI", 10, "italic"))
        self.lbl_w_status.pack(side="left", padx=20)

        frame_addons = tk.Frame(self.frame_documento, bg=f"#{AMARILLO_CLARO_HEX}", padx=20, pady=20)
        frame_addons.pack(fill="x", pady=15)
        tk.Label(frame_addons, text="COMPLEMENTOS DE INTELIGENCIA", font=("Segoe UI", 8, "bold"), bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{AMARILLO_ACENTO_HEX}").pack(anchor="w", pady=(0, 10))
        tk.Checkbutton(
            frame_addons, text="Estructurar e inyectar Resumen Ejecutivo dinamico", variable=self.var_resumen,
            font=("Segoe UI", 10), bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_TEXTO_HEX}", selectcolor=f"#{BEIGE_HEX}",
        ).pack(anchor="w", pady=6)
        tk.Checkbutton(
            frame_addons, text="Generar esquema estructurado para presentaciones", variable=self.var_presentacion,
            font=("Segoe UI", 10), bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_TEXTO_HEX}", selectcolor=f"#{BEIGE_HEX}",
        ).pack(anchor="w", pady=6)

        tk.Button(
            self.frame_documento, text="Aplicar Formato y Guardar Manuscrito", font=("Segoe UI", 10, "bold"),
            bg=f"#{AMARILLO_ACENTO_HEX}", fg=f"#{CAFE_TEXTO_HEX}", activebackground="#E0B015", bd=0, cursor="hand2",
            command=self.ejecutar_opcion_1,
        ).pack(pady=(20, 0), ipady=14, fill="x")

    def buscar_word(self):
        ruta = filedialog.askopenfilename(filetypes=[("Documentos Word", "*.docx")])
        if ruta:
            self.ruta_word.set(ruta)
            self.lbl_w_status.config(text=os.path.basename(ruta), fg=f"#{AMARILLO_ACENTO_HEX}")

    def ejecutar_opcion_1(self):
        r_in = self.ruta_word.get()
        if not r_in or not os.path.exists(r_in):
            messagebox.showerror("Error de Archivo", "Por favor, selecciona un archivo real .docx")
            return
        r_out = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word", "*.docx")], initialfile="Documento_Formato_APA7.docx")
        if not r_out:
            return

        try:
            doc = Document(r_in)
            try:
                doc.styles["List Bullet"]
            except KeyError:
                bullet_style = doc.styles.add_style("List Bullet", WD_STYLE_TYPE.PARAGRAPH)
                bullet_style.font.name = "Times New Roman"
                bullet_style.font.size = Pt(12)

            for s in doc.sections:
                s.page_width, s.page_height = Inches(8.5), Inches(11.0)
                s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(1.0)

            doc.styles["Normal"].font.name = "Times New Roman"
            doc.styles["Normal"].font.size = Pt(12)
            paragraphs_clean = [p.text.strip() for p in doc.paragraphs if p.text.strip() and not p.style.name.startswith("Heading")]

            for p in doc.paragraphs:
                if not p.text.strip():
                    continue
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.line_spacing = 2.0
                p.paragraph_format.space_after = p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.first_line_indent = Inches(0 if p.style.name.startswith("Heading") else 0.5)

            if self.var_resumen.get() and paragraphs_clean:
                doc.add_page_break()
                p_t = doc.add_paragraph()
                p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_t.add_run("Resumen Informativo de Avance").bold = True
                p_b = doc.add_paragraph()
                p_b.paragraph_format.line_spacing = 2.0
                p_b.paragraph_format.first_line_indent = Inches(0.5)
                p_b.add_run(f"Consolidado del manuscrito estructurado. Linea base del estudio: {paragraphs_clean[0][:400]}...")

            if self.var_presentacion.get() and paragraphs_clean:
                doc.add_page_break()
                p_t = doc.add_paragraph()
                p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_t.add_run("Esquema Estructurado para Diapositivas").bold = True
                for i in range(1, 3):
                    if len(paragraphs_clean) >= i:
                        p_d = doc.add_paragraph()
                        p_d.add_run(f"Diapositiva {i}: Eje Tematico").bold = True
                        p_v = doc.add_paragraph()
                        p_v.paragraph_format.left_indent, p_v.paragraph_format.first_line_indent = Inches(0.5), Inches(-0.25)
                        p_v.add_run(f"Idea clave: {paragraphs_clean[i-1][:150]}...")

            doc.save(r_out)
            messagebox.showinfo("Exito", "El Word ha sido formateado en Carta APA 7 con exito.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    #  OPCION 2
    def construir_pantalla_datos_analiticos(self):
        btn_back = tk.Button(
            self.frame_datos_analiticos, text="<-  Volver al Menu", font=("Segoe UI", 9, "bold"), bg=f"#{AMARILLO_CLARO_HEX}",
            fg=f"#{CAFE_TEXTO_HEX}", activebackground=f"#{BEIGE_OSCURO_HEX}", activeforeground=f"#{CAFE_TEXTO_HEX}", bd=0, cursor="hand2",
            padx=15, pady=6, command=self.mostrar_menu_principal,
        )
        btn_back.pack(anchor="w", pady=(0, 20))

        tk.Label(
            self.frame_datos_analiticos, text="Analizador Integral de Encuestas (Excel / CSV)",
            font=("Segoe UI", 14, "bold"), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
        ).pack(anchor="w", pady=(0, 8))
        tk.Label(
            self.frame_datos_analiticos,
            text="Al insertar el archivo, el sistema genera automaticamente un Reporte Word formal/futurista, "
                 "un libro Excel con una pestana por pregunta y un visualizador de graficos interactivo.",
            font=("Segoe UI", 9), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_SUAVE_HEX}", wraplength=760, justify="left",
        ).pack(anchor="w", pady=(0, 15))

        frame_file = tk.Frame(self.frame_datos_analiticos, bg=f"#{AMARILLO_CLARO_HEX}", padx=20, pady=20)
        frame_file.pack(fill="x", pady=10)
        tk.Label(frame_file, text="CONJUNTO DE DATOS (.xlsx / .csv)", font=("Segoe UI", 8, "bold"), bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{AMARILLO_ACENTO_HEX}").pack(anchor="w", pady=(0, 10))
        tk.Button(
            frame_file, text="Cargar Archivo", font=("Segoe UI", 9, "bold"), bg=f"#{BEIGE_OSCURO_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
            activebackground=f"#{CAFE_SUAVE_HEX}", bd=0, cursor="hand2", padx=15, pady=6, command=self.buscar_excel,
        ).pack(side="left")
        self.lbl_e_status = tk.Label(frame_file, text="Esperando hoja de calculo...", bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_SUAVE_HEX}", font=("Segoe UI", 10, "italic"))
        self.lbl_e_status.pack(side="left", padx=20)

        self.txt_preview_excel = tk.Text(self.frame_datos_analiticos, height=12, bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_TEXTO_HEX}", font=("Consolas", 10), bd=0, padx=20, pady=20, insertbackground="white")
        self.txt_preview_excel.pack(fill="x", pady=15)
        self.txt_preview_excel.insert("1.0", "Las estadisticas calculadas por pregunta apareceran en esta consola al procesar...")
        self.txt_preview_excel.config(state="disabled")

        self.frame_select_grafico = tk.Frame(self.frame_datos_analiticos, bg=f"#{AMARILLO_CLARO_HEX}", padx=15, pady=12)
        tk.Label(self.frame_select_grafico, text="VISUALIZADOR DE VARIABLE ESPECIFICA", font=("Segoe UI", 8, "bold"), bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_SUAVE_HEX}").pack(anchor="w", pady=(0, 8))
        fila_grafico = tk.Frame(self.frame_select_grafico, bg=f"#{AMARILLO_CLARO_HEX}")
        fila_grafico.pack(fill="x")
        self.combo_variables = ttk.Combobox(fila_grafico, state="readonly", font=("Segoe UI", 9))
        self.combo_variables.pack(fill="x", side="left", expand=True, padx=(0, 15))
        # Al cambiar la pregunta seleccionada, el grafico de barras se actualiza
        # de inmediato para reflejar siempre la variable actual.
        self.combo_variables.bind("<<ComboboxSelected>>", lambda e: self.mostrar_ventana_grafico_nativa())
        self.btn_ver_grafico = tk.Button(
            fila_grafico, text="Graficar Pregunta Seleccionada", font=("Segoe UI", 9, "bold"), bg="#10b981",
            fg="white", bd=0, state="disabled", cursor="hand2", command=self.mostrar_ventana_grafico_nativa,
        )
        self.btn_ver_grafico.pack(side="right", padx=5)

        self.btn_run_excel = tk.Button(
            self.frame_datos_analiticos, text="Procesar Encuesta y Generar Reporte Word + Excel",
            font=("Segoe UI", 10, "bold"), bg="#0284c7", fg="white", activebackground="#0369a1", bd=0,
            cursor="hand2", command=self.ejecutar_opcion_2,
        )
        self.btn_run_excel.pack(pady=(15, 0), ipady=14, fill="x")

    def buscar_excel(self):
        ruta = filedialog.askopenfilename(filetypes=[("Archivos de Datos", "*.xlsx *.csv")])
        if ruta:
            self.ruta_excel.set(ruta)
            self.lbl_e_status.config(text=os.path.basename(ruta), fg=f"#{AMARILLO_ACENTO_HEX}")
            self.btn_ver_grafico.config(state="disabled")
            self.frame_select_grafico.pack_forget()

    def ejecutar_opcion_2(self):
        r_in = self.ruta_excel.get()
        if not r_in or not os.path.exists(r_in):
            messagebox.showerror("Error", "Selecciona primero un archivo de Excel o CSV valido.")
            return

        r_out_excel = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile="Analisis_Completo_Encuesta.xlsx"
        )
        if not r_out_excel:
            return
        r_out_word = r_out_excel.replace(".xlsx", "_Reporte_Word.docx")

        def pipeline_async():
            try:
                self.root.config(cursor="wait")

                df = pd.read_csv(r_in) if r_in.endswith(".csv") else pd.read_excel(r_in)

                #  Reporte Word formal 
                doc_reporte, texto_preview, analisis = construir_reporte_word_encuesta(
                    df, titulo_reporte="Reporte Estadistico de Encuesta"
                )
                self.analisis_maestro_columnas = analisis

                #  Excel con una pestana por pregunta
                wb = openpyxl.Workbook()
                wb.remove(wb.active)
                for index_col, (col, datos) in enumerate(analisis.items(), start=1):
                    nombre_pestana = f"Pregunta_{index_col}"[:30]
                    ws = wb.create_sheet(title=nombre_pestana)
                    ws.append(["Pregunta Completa:", col])
                    ws.append([])
                    ws.append(["Metrica Estadistica", "Valor Calculado"])
                    ws.append(["Moda (Valor mas comun)", str(datos["Moda"])])
                    ws.append(["Mediana (Punto medio)", str(datos["Mediana"])])
                    ws.append(["Varianza (Dispersion)", str(datos["Varianza"])])
                    ws.append([])
                    ws.append(["Opcion / Respuesta", "Frecuencia"])
                    for item in datos["Frecuencias"]:
                        ws.append([item["Elemento"], item["Conteo"]])

                wb.save(r_out_excel)
                doc_reporte.save(r_out_word)

                self.root.after(0, lambda: self.finalizar_pipeline_encuesta(texto_preview, r_out_excel, r_out_word))
            except Exception as e:
                self.root.config(cursor="")
                self.root.after(0, lambda: messagebox.showerror("Error Critico del Pipeline", str(e)))

        threading.Thread(target=pipeline_async, daemon=True).start()

    def finalizar_pipeline_encuesta(self, preview_texto, ruta_excel_out, ruta_word_out):
        self.root.config(cursor="")
        self.txt_preview_excel.config(state="normal")
        self.txt_preview_excel.delete("1.0", "end")
        self.txt_preview_excel.insert("1.0", preview_texto)
        self.txt_preview_excel.config(state="disabled")

        lista_preguntas = list(self.analisis_maestro_columnas.keys())
        self.combo_variables["values"] = lista_preguntas
        if lista_preguntas:
            self.combo_variables.current(0)

        self.frame_select_grafico.pack(fill="x", pady=10, before=self.btn_run_excel)
        self.btn_ver_grafico.config(state="normal")

        self.root.after(150, self.mostrar_ventana_grafico_nativa)

        messagebox.showinfo(
            "Exito Absoluto",
            "Se analizaron todas las preguntas de la encuesta.\n\n"
            f"- Reporte Word formal/futurista: {os.path.basename(ruta_word_out)}\n"
            f"- Libro Excel con pestanas individuales: {os.path.basename(ruta_excel_out)}",
        )

    # visualizador en ventana emergente 
    def mostrar_ventana_grafico_nativa(self):
        pregunta_seleccionada = self.combo_variables.get()
        if not pregunta_seleccionada or pregunta_seleccionada not in self.analisis_maestro_columnas:
            return

        datos_pregunta = self.analisis_maestro_columnas[pregunta_seleccionada]
        frecuencias = datos_pregunta["Frecuencias"]

        # Si ya existe una ventana de grafico abierta, se cierra antes de abrir la nueva
        if getattr(self, "_ventana_grafico_actual", None) is not None:
            try:
                self._ventana_grafico_actual.destroy()
            except Exception:
                pass
            self._ventana_grafico_actual = None

        ventana_grafico = tk.Toplevel(self.root)
        self._ventana_grafico_actual = ventana_grafico
        ventana_grafico.title("Visualizador Grafico de Encuestas")
        ventana_grafico.geometry("820x600")
        ventana_grafico.configure(bg=f"#{BEIGE_HEX}")
        ventana_grafico.transient(self.root)
        ventana_grafico.grab_set()

        def _al_cerrar():
            self._ventana_grafico_actual = None
            ventana_grafico.destroy()
        ventana_grafico.protocol("WM_DELETE_WINDOW", _al_cerrar)

        titulo_corto = pregunta_seleccionada if len(pregunta_seleccionada) < 75 else pregunta_seleccionada[:72] + "..."
        tk.Label(
            ventana_grafico, text=titulo_corto, font=("Segoe UI", 11, "bold"), bg=f"#{BEIGE_HEX}",
            fg=f"#{CAFE_TEXTO_HEX}", wraplength=720, justify="center",
        ).pack(pady=(15, 0))

        mediana_txt = f"{datos_pregunta['Mediana']:.2f}" if isinstance(datos_pregunta["Mediana"], (int, float)) else str(datos_pregunta["Mediana"])
        var_txt = f"{datos_pregunta['Varianza']:.2f}" if isinstance(datos_pregunta["Varianza"], (int, float)) else str(datos_pregunta["Varianza"])
        tk.Label(
            ventana_grafico,
            text=f"Moda: {datos_pregunta['Moda']}   |   Mediana: {mediana_txt}   |   Varianza: {var_txt}",
            font=("Segoe UI", 9, "italic"), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_SUAVE_HEX}",
        ).pack(pady=(2, 10))

        canvas = tk.Canvas(ventana_grafico, bg=f"#{AMARILLO_CLARO_HEX}", highlightthickness=1,
                            highlightbackground=f"#{BEIGE_OSCURO_HEX}")
        canvas.pack(fill="both", expand=True, padx=30, pady=(0, 10))
        ventana_grafico.update()
        canvas_width, canvas_height = canvas.winfo_width(), canvas.winfo_height()

        max_valor_global = max([item["Conteo"] for item in frecuencias]) if frecuencias else 1
        if max_valor_global == 0:
            max_valor_global = 1

        margen_izq, margen_inf, margen_sup, margen_der = 80, 60, 30, 30
        ancho_disponible = canvas_width - margen_izq - margen_der
        alto_disponible = canvas_height - margen_sup - margen_inf

        canvas.create_line(margen_izq, margen_sup, margen_izq, canvas_height - margen_inf, fill=f"#{CAFE_SUAVE_HEX}", width=2)
        canvas.create_line(margen_izq, canvas_height - margen_inf, canvas_width - margen_der, canvas_height - margen_inf, fill=f"#{CAFE_SUAVE_HEX}", width=2)

        for i in range(5):
            val_escala = (max_valor_global / 4) * i
            y_pos = (canvas_height - margen_inf) - (val_escala / max_valor_global * alto_disponible)
            canvas.create_text(margen_izq - 12, y_pos, text=f"{int(val_escala)}", fill=f"#{CAFE_SUAVE_HEX}", font=("Segoe UI", 8), anchor="e")
            if i > 0:
                canvas.create_line(margen_izq, y_pos, canvas_width - margen_der, y_pos, fill=f"#{BEIGE_OSCURO_HEX}", dash=(3, 3))

        num_barras = len(frecuencias)
        if num_barras == 0:
            return
        ancho_grupo = ancho_disponible / num_barras
        colores_beige_grafico = [AMARILLO_ACENTO_HEX, "D9A61E", CAFE_SUAVE_HEX, "C9973A", "E0B015"]

        for idx, item in enumerate(frecuencias):
            x_bar_start = margen_izq + (idx * ancho_grupo) + (ancho_grupo * 0.15)
            ancho_barra = ancho_grupo * 0.7
            h_frecuencia = (item["Conteo"] / max_valor_global) * alto_disponible
            y_bar_start = (canvas_height - margen_inf) - h_frecuencia

            color_barra = f"#{colores_beige_grafico[idx % len(colores_beige_grafico)]}"
            canvas.create_rectangle(x_bar_start, y_bar_start, x_bar_start + ancho_barra, canvas_height - margen_inf, fill=color_barra, outline="")
            canvas.create_text(x_bar_start + (ancho_barra / 2), y_bar_start - 10, text=str(item["Conteo"]), fill=f"#{CAFE_TEXTO_HEX}", font=("Segoe UI", 8, "bold"))
            texto_eje_x = item["Elemento"] if len(item["Elemento"]) < 14 else item["Elemento"][:12] + ".."
            canvas.create_text(x_bar_start + (ancho_barra / 2), canvas_height - margen_inf + 15, text=texto_eje_x, fill=f"#{CAFE_TEXTO_HEX}", font=("Segoe UI", 8, "bold"), anchor="n")

        tk.Button(
            ventana_grafico, text="Cerrar Grafico", font=("Segoe UI", 9, "bold"), bg=f"#{BEIGE_OSCURO_HEX}",
            fg=f"#{CAFE_TEXTO_HEX}", bd=0, cursor="hand2", command=_al_cerrar,
            padx=25, pady=6,
        ).pack(pady=(0, 15))

    # ============================================================== OPCION 3
    def construir_pantalla_gmail_agente(self):
        btn_back = tk.Button(
            self.frame_gmail_agente, text="<-  Volver al Menu", font=("Segoe UI", 9, "bold"), bg=f"#{AMARILLO_CLARO_HEX}",
            fg=f"#{CAFE_TEXTO_HEX}", activebackground=f"#{BEIGE_OSCURO_HEX}", activeforeground=f"#{CAFE_TEXTO_HEX}", bd=0, cursor="hand2",
            padx=15, pady=6, command=self.mostrar_menu_principal,
        )
        btn_back.pack(anchor="w", pady=(0, 15))

        tk.Label(
            self.frame_gmail_agente, text="Asistente de Correo Inteligente (Agente IA)",
            font=("Segoe UI", 14, "bold"), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
        ).pack(anchor="w", pady=(0, 15))

        # --- FRAME TOP CON LOS BOTONES Y LA ENTRADA LIMPIA ---
        frame_top = tk.Frame(self.frame_gmail_agente, bg=f"#{BEIGE_HEX}")
        frame_top.pack(fill="x", pady=(0, 10))
        
        tk.Label(frame_top, text=f"Vinculado: {MI_GMAIL}", font=("Segoe UI", 10, "bold"), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_TEXTO_HEX}").pack(side="left", anchor="w", pady=5)
        
        # Boton Cargar Correos (lado derecho)
        tk.Button(
            frame_top, text="Cargar Correos", bg=f"#{CAFE_SUAVE_HEX}", fg="white", font=("Segoe UI", 9, "bold"), bd=0,
            cursor="hand2", padx=20, pady=6, activebackground=f"#{BEIGE_OSCURO_HEX}", command=self.sincronizar_bandeja,
        ).pack(side="right", padx=(10, 0))

        # Boton Nuevo para redactar (al lado de cargar correos)
        tk.Button(
            frame_top, text="Redactar Correo", font=("Segoe UI", 9, "bold"), bg=f"#{AMARILLO_ACENTO_HEX}", 
            fg=f"#{CAFE_TEXTO_HEX}", bd=0, cursor="hand2", padx=15, pady=6, activebackground="#E0B015",
            command=self.abrir_ventana_redactar_nuevo_correo,
        ).pack(side="right")

        # Entrada para escribir el destinatario 
        self.entry_nuevo_destinatario = tk.Entry(
            frame_top, font=("Segoe UI", 10), bg=f"#{BLANCO_HEX}", fg=f"#{CAFE_TEXTO_HEX}", bd=0, insertbackground=f"#{CAFE_TEXTO_HEX}", width=25
        )
        self.entry_nuevo_destinatario.pack(side="right", padx=(0, 10), ipady=5)
        # -------------------------------------------------------------------

        frame_select = tk.Frame(self.frame_gmail_agente, bg=f"#{AMARILLO_CLARO_HEX}", padx=15, pady=15)
        frame_select.pack(fill="x", pady=5)
        self.combo_mails = ttk.Combobox(frame_select, state="readonly", font=("Segoe UI", 10))
        self.combo_mails.pack(fill="x", side="left", expand=True, padx=(0, 15), ipady=4)
        self.combo_mails.set("<- Haz clic en 'Cargar Correos' ->")
        tk.Button(
            frame_select, text="Analizar Correo", font=("Segoe UI", 9, "bold"), bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{AMARILLO_ACENTO_HEX}",
            bd=1, relief="solid", cursor="hand2", padx=15, pady=5, activebackground=f"#{BEIGE_OSCURO_HEX}",
            command=self.ejecutar_opcion_3,
        ).pack(side="right")

        frame_preview = tk.Frame(self.frame_gmail_agente, bg=f"#{AMARILLO_CLARO_HEX}", padx=5, pady=5)
        frame_preview.pack(fill="both", expand=True, pady=10)
        self.txt_respuesta_ia = tk.Text(frame_preview, bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_TEXTO_HEX}", font=("Segoe UI", 10), bd=0, padx=20, pady=20, insertbackground="white")
        self.txt_respuesta_ia.pack(fill="both", expand=True)
        self.txt_respuesta_ia.insert("1.0", "Selecciona un correo arriba y haz clic en 'Analizar Correo' para procesar la respuesta con IA...")
        self.txt_respuesta_ia.config(state="disabled")

        frame_acciones = tk.Frame(self.frame_gmail_agente, bg=f"#{BEIGE_HEX}")
        frame_acciones.pack(fill="x", pady=(10, 0))
        self.btn_enviar_correo = tk.Button(
            frame_acciones, text="Enviar Gmail (IA)", font=("Segoe UI", 10, "bold"), bg=f"#{AMARILLO_ACENTO_HEX}",
            fg=f"#{CAFE_TEXTO_HEX}", bd=0, cursor="hand2", state="disabled", activebackground="#E0B015",
            command=self.abrir_ventana_editar_envio_correo,
        )
        self.btn_enviar_correo.pack(side="left", fill="x", expand=True, ipady=12, padx=(0, 10))
        self.btn_verificar_servidor = tk.Button(
            frame_acciones, text="Verificar en Servidor", font=("Segoe UI", 10, "bold"), bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
            bd=1, relief="solid", cursor="hand2", state="disabled", activebackground=f"#{BEIGE_OSCURO_HEX}",
            command=self.ejecutar_verificacion_servidor,
        )
        self.btn_verificar_servidor.pack(side="right", fill="x", expand=True, ipady=12)

        self.lbl_estado_envio = tk.Label(
            self.frame_gmail_agente, text="", font=("Segoe UI", 9, "bold"), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_SUAVE_HEX}",
        )
        self.lbl_estado_envio.pack(fill="x", pady=(10, 0))

    def sincronizar_bandeja(self):
        try:
            self.root.config(cursor="wait")
            self.root.update()
            self.correos_memorizados = obtener_correos_recientes(MI_GMAIL, MI_PASSWORD_DE_APLICACION)
            items = [f"De: {c['remitente'][:22]} | Asunto: {c['asunto'][:48]}" for c in self.correos_memorizados]
            self.combo_mails["values"] = items
            if items:
                self.combo_mails.current(0)
                messagebox.showinfo("Exito", "Bandeja sincronizada correctamente.")
            else:
                self.combo_mails.set("No hay correos pendientes.")
            self.root.config(cursor="")
        except Exception as e:
            self.root.config(cursor="")
            messagebox.showerror("Error de Autenticacion", str(e))

    def ejecutar_opcion_3(self):
        idx = self.combo_mails.current()
        if idx < 0 or not self.correos_memorizados:
            messagebox.showerror("Error", "Selecciona un correo del listado desplegable.")
            return

        correo = self.correos_memorizados[idx]
        self.root.config(cursor="wait")
        self.root.update()

        prompt = (
            "Eres un asistente ejecutivo inteligente. Lee el correo del usuario, clasificalo "
            "en una categoria (Urgente, Soporte, Comercial, Ejecutivo) y genera una respuesta corporativa formal, "
            "educada y atenta en espanol respondiendo puntualmente a lo que te preguntan. Estructura el resultado visible asi:\n\n"
            "------------------------------------------\n"
            " CLASIFICACION DEL MENSAJE: [Categoria]\n"
            "------------------------------------------\n"
            "PROPUESTA DE RESPUESTA FORMAL:\n"
            "[Tu texto completo de respuesta aqui]"
        )

        resultado_ia = consultar_gemini(MI_GEMINI_TOKEN, prompt, correo["cuerpo"])
        self.root.config(cursor="")

        if "PROPUESTA DE RESPUESTA FORMAL:\n" in resultado_ia:
            self.ultimo_analisis_cuerpo = resultado_ia.split("PROPUESTA DE RESPUESTA FORMAL:\n")[1].strip()
        else:
            self.ultimo_analisis_cuerpo = resultado_ia

        self.txt_respuesta_ia.config(state="normal")
        self.txt_respuesta_ia.delete("1.0", "end")
        encabezado_info = f"Remitente original: {correo['remitente']}\nAsunto original: {correo['asunto']}\n\n"
        self.txt_respuesta_ia.insert("1.0", encabezado_info + resultado_ia)
        self.txt_respuesta_ia.config(state="disabled")

        self.btn_enviar_correo.config(state="normal")
        messagebox.showinfo("Procesamiento Listo", "El analisis dinamico se ha cargado en el panel inferior.")

    def abrir_ventana_redactar_nuevo_correo(self):
        """Abre una interfaz limpia para redactar un correo desde cero al destino escrito."""
        destinatario = self.entry_nuevo_destinatario.get().strip()
        
        if not destinatario or "@" not in destinatario:
            messagebox.showwarning("Destinatario invalido", "Por favor, escribe un correo de destino valido en la caja de texto.")
            return

        if getattr(self, "_ventana_envio_actual", None) is not None:
            try:
                self._ventana_envio_actual.destroy()
            except Exception:
                pass
            self._ventana_envio_actual = None

        ventana = tk.Toplevel(self.root)
        self._ventana_envio_actual = ventana
        ventana.title("Redactar Nuevo Correo")
        ventana.geometry("640x620") # <--- Incrementado el alto para dar espacio
        ventana.configure(bg=f"#{BEIGE_HEX}")
        ventana.transient(self.root)
        ventana.grab_set()

        def _al_cerrar():
            self._ventana_envio_actual = None
            ventana.destroy()
        ventana.protocol("WM_DELETE_WINDOW", _al_cerrar)

        contenedor = tk.Frame(ventana, bg=f"#{BEIGE_HEX}", padx=20, pady=20)
        contenedor.pack(fill="both", expand=True)

        tk.Label(
            contenedor, text="Redactar nuevo mensaje",
            font=("Segoe UI", 13, "bold"), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
        ).pack(anchor="w", pady=(0, 4))
        tk.Frame(contenedor, bg=f"#{AMARILLO_ACENTO_HEX}", width=70, height=4).pack(anchor="w", pady=(0, 15))

        # ---- Destinatario ----
        tk.Label(
            contenedor, text="PARA:", font=("Segoe UI", 8, "bold"), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_SUAVE_HEX}",
        ).pack(anchor="w")
        entrada_destinatario = tk.Entry(contenedor, font=("Segoe UI", 10), bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_TEXTO_HEX}", bd=0)
        entrada_destinatario.insert(0, destinatario)
        entrada_destinatario.pack(fill="x", ipady=6, pady=(2, 12))

        # ---- Asunto ----
        tk.Label(
            contenedor, text="ASUNTO (Variable):", font=("Segoe UI", 8, "bold"), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_SUAVE_HEX}",
        ).pack(anchor="w")
        entrada_asunto = tk.Entry(contenedor, font=("Segoe UI", 10), bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_TEXTO_HEX}", bd=0)
        entrada_asunto.pack(fill="x", ipady=6, pady=(2, 12))

        # ---- Cuerpo ----
        tk.Label(
            contenedor, text="MENSAJE (Variable):", font=("Segoe UI", 8, "bold"), bg=f"#{BEIGE_HEX}",
            fg=f"#{CAFE_SUAVE_HEX}",
        ).pack(anchor="w")
        frame_cuerpo = tk.Frame(contenedor, bg=f"#{AMARILLO_CLARO_HEX}")
        frame_cuerpo.pack(fill="both", expand=True, pady=(2, 15))
        
        # Ajustado el alto nativo del Text para que no desplace los botones fuera de la pantalla
        texto_cuerpo = tk.Text(
            frame_cuerpo, font=("Segoe UI", 10), bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
            bd=0, padx=10, pady=10, wrap="word", insertbackground=f"#{CAFE_TEXTO_HEX}", height=10
        )
        scroll_cuerpo = tk.Scrollbar(frame_cuerpo, command=texto_cuerpo.yview)
        texto_cuerpo.configure(yscrollcommand=scroll_cuerpo.set)
        texto_cuerpo.pack(side="left", fill="both", expand=True)
        scroll_cuerpo.pack(side="right", fill="y")

        self.lbl_estado_envio_popup = tk.Label(
            contenedor, text="", font=("Segoe UI", 9, "bold"), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_SUAVE_HEX}",
        )
        self.lbl_estado_envio_popup.pack(fill="x", pady=(0, 10))

        # ---- Botones Confirmar / Cancelar ----
        frame_botones = tk.Frame(contenedor, bg=f"#{BEIGE_HEX}")
        frame_botones.pack(fill="x", side="bottom") # Empujados estrictamente al fondo del contenedor

        def confirmar_envio():
            destinatario_final = entrada_destinatario.get().strip()
            asunto_final = entrada_asunto.get().strip()
            cuerpo_final = texto_cuerpo.get("1.0", "end").strip()

            if not destinatario_final or not asunto_final or not cuerpo_final:
                messagebox.showerror("Datos Incompletos", "El destinatario, el asunto y el mensaje no pueden estar vacios.", parent=ventana)
                return

            self._despachar_correo_gmail(destinatario_final, asunto_final, cuerpo_final, ventana)

        def cancelar_envio():
            _al_cerrar()

        tk.Button(
            frame_botones, text="Enviar Correo", font=("Segoe UI", 10, "bold"), bg=f"#{AMARILLO_ACENTO_HEX}",
            fg=f"#{CAFE_TEXTO_HEX}", bd=0, cursor="hand2", activebackground="#E0B015",
            command=confirmar_envio, padx=15, pady=10,
        ).pack(side="left", fill="x", expand=True, padx=(0, 8))
        
        tk.Button(
            frame_botones, text="Cancelar", font=("Segoe UI", 10, "bold"), bg=f"#{BEIGE_OSCURO_HEX}",
            fg=f"#{CAFE_TEXTO_HEX}", bd=0, cursor="hand2", activebackground="#D9CBA0",
            command=cancelar_envio, padx=15, pady=10,
        ).pack(side="right", fill="x", expand=True, padx=(8, 0))


    def abrir_ventana_editar_envio_correo(self):
        """Abre una ventana emergente que muestra el correo generado por la IA,
        permite editar el destinatario, el asunto y el cuerpo antes de enviarlo,
        y ofrece botones de Confirmar / Cancelar para el despacho final."""
        idx = self.combo_mails.current()
        if idx < 0 or not self.correos_memorizados:
            messagebox.showerror("Error", "Selecciona un correo del listado desplegable.")
            return
        correo = self.correos_memorizados[idx]

        if getattr(self, "_ventana_envio_actual", None) is not None:
            try:
                self._ventana_envio_actual.destroy()
            except Exception:
                pass
            self._ventana_envio_actual = None

        ventana = tk.Toplevel(self.root)
        self._ventana_envio_actual = ventana
        ventana.title("Enviar Gmail - Revisar y Confirmar")
        ventana.geometry("640x620")
        ventana.configure(bg=f"#{BEIGE_HEX}")
        ventana.transient(self.root)
        ventana.grab_set()

        def _al_cerrar():
            self._ventana_envio_actual = None
            ventana.destroy()
        ventana.protocol("WM_DELETE_WINDOW", _al_cerrar)

        contenedor = tk.Frame(ventana, bg=f"#{BEIGE_HEX}", padx=20, pady=20)
        contenedor.pack(fill="both", expand=True)

        tk.Label(
            contenedor, text="Revisa y edita el correo antes de enviarlo por Gmail",
            font=("Segoe UI", 13, "bold"), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
        ).pack(anchor="w", pady=(0, 4))
        tk.Frame(contenedor, bg=f"#{AMARILLO_ACENTO_HEX}", width=70, height=4).pack(anchor="w", pady=(0, 15))

        tk.Label(
            contenedor, text="PARA:", font=("Segoe UI", 8, "bold"), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_SUAVE_HEX}",
        ).pack(anchor="w")
        entrada_destinatario = tk.Entry(contenedor, font=("Segoe UI", 10), bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_TEXTO_HEX}", bd=0)
        entrada_destinatario.insert(0, correo["remitente"])
        entrada_destinatario.pack(fill="x", ipady=6, pady=(2, 12))

        tk.Label(
            contenedor, text="ASUNTO:", font=("Segoe UI", 8, "bold"), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_SUAVE_HEX}",
        ).pack(anchor="w")
        entrada_asunto = tk.Entry(contenedor, font=("Segoe UI", 10), bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_TEXTO_HEX}", bd=0)
        entrada_asunto.insert(0, f"Re: {correo['asunto']}")
        entrada_asunto.pack(fill="x", ipady=6, pady=(2, 12))

        tk.Label(
            contenedor, text="CUERPO DEL MENSAJE (editable):", font=("Segoe UI", 8, "bold"), bg=f"#{BEIGE_HEX}",
            fg=f"#{CAFE_SUAVE_HEX}",
        ).pack(anchor="w")
        frame_cuerpo = tk.Frame(contenedor, bg=f"#{AMARILLO_CLARO_HEX}")
        frame_cuerpo.pack(fill="both", expand=True, pady=(2, 15))
        texto_cuerpo = tk.Text(
            frame_cuerpo, font=("Segoe UI", 10), bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
            bd=0, padx=10, pady=10, wrap="word", insertbackground=f"#{CAFE_TEXTO_HEX}",
        )
        scroll_cuerpo = tk.Scrollbar(frame_cuerpo, command=texto_cuerpo.yview)
        texto_cuerpo.configure(yscrollcommand=scroll_cuerpo.set)
        texto_cuerpo.pack(side="left", fill="both", expand=True)
        scroll_cuerpo.pack(side="right", fill="y")
        texto_cuerpo.insert("1.0", self.ultimo_analisis_cuerpo)

        self.lbl_estado_envio_popup = tk.Label(
            contenedor, text="", font=("Segoe UI", 9, "bold"), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_SUAVE_HEX}",
        )
        self.lbl_estado_envio_popup.pack(fill="x", pady=(0, 10))

        frame_botones = tk.Frame(contenedor, bg=f"#{BEIGE_HEX}")
        frame_botones.pack(fill="x")

        def confirmar_envio():
            destinatario_final = entrada_destinatario.get().strip()
            asunto_final = entrada_asunto.get().strip()
            cuerpo_final = texto_cuerpo.get("1.0", "end").strip()
            if not destinatario_final or not asunto_final or not cuerpo_final:
                messagebox.showerror("Datos Incompletos", "El destinatario, el asunto y el cuerpo no pueden estar vacios.", parent=ventana)
                return
            confirmado = messagebox.askokcancel(
                "Confirmacion de Envio",
                f"Deseas enviar este correo electronico de manera formal?\n\nPara: {destinatario_final}\n"
                f"Asunto: {asunto_final}\n\nPresiona Aceptar para confirmar o Cancelar para abortar la operacion.",
                parent=ventana,
            )
            if not confirmado:
                return
            self._despachar_correo_gmail(destinatario_final, asunto_final, cuerpo_final, ventana)

        def cancelar_envio():
            self.lbl_estado_envio.config(text="Envio cancelado por el usuario.", fg="#B23B3B")
            _al_cerrar()

        tk.Button(
            frame_botones, text="Confirmar y Enviar", font=("Segoe UI", 10, "bold"), bg=f"#{AMARILLO_ACENTO_HEX}",
            fg=f"#{CAFE_TEXTO_HEX}", bd=0, cursor="hand2", activebackground="#E0B015",
            command=confirmar_envio, padx=15, pady=10,
        ).pack(side="left", fill="x", expand=True, padx=(0, 8))
        tk.Button(
            frame_botones, text="Cancelar", font=("Segoe UI", 10, "bold"), bg=f"#{BEIGE_OSCURO_HEX}",
            fg=f"#{CAFE_TEXTO_HEX}", bd=0, cursor="hand2", activebackground="#D9CBA0",
            command=cancelar_envio, padx=15, pady=10,
        ).pack(side="right", fill="x", expand=True, padx=(8, 0))


    def abrir_ventana_redactar_nuevo_correo(self):
        destinatario = self.entry_nuevo_destinatario.get().strip()
        
        if not destinatario or "@" not in destinatario:
            messagebox.showwarning("Destinatario invalido", "Por favor, escribe un correo de destino valido en la caja de texto.")
            return

        if getattr(self, "_ventana_envio_actual", None) is not None:
            try:
                self._ventana_envio_actual.destroy()
            except Exception:
                pass
            self._ventana_envio_actual = None

        ventana = tk.Toplevel(self.root)
        self._ventana_envio_actual = ventana
        ventana.title("Redactar Nuevo Correo")
        ventana.geometry("640x640") 
        ventana.configure(bg=f"#{BEIGE_HEX}")
        ventana.transient(self.root)
        ventana.grab_set()

        def _al_cerrar():
            self._ventana_envio_actual = None
            ventana.destroy()
        ventana.protocol("WM_DELETE_WINDOW", _al_cerrar)

        contenedor = tk.Frame(ventana, bg=f"#{BEIGE_HEX}", padx=20, pady=20)
        contenedor.pack(fill="both", expand=True)

        tk.Label(
            contenedor, text="Redactar nuevo mensaje",
            font=("Segoe UI", 13, "bold"), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
        ).pack(anchor="w", pady=(0, 4))
        tk.Frame(contenedor, bg=f"#{AMARILLO_ACENTO_HEX}", width=70, height=4).pack(anchor="w", pady=(0, 15))

        tk.Label(
            contenedor, text="PARA:", font=("Segoe UI", 8, "bold"), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_SUAVE_HEX}",
        ).pack(anchor="w")
        entrada_destinatario = tk.Entry(contenedor, font=("Segoe UI", 10), bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_TEXTO_HEX}", bd=0)
        entrada_destinatario.insert(0, destinatario)
        entrada_destinatario.pack(fill="x", ipady=6, pady=(2, 12))

        tk.Label(
            contenedor, text="ASUNTO (Variable):", font=("Segoe UI", 8, "bold"), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_SUAVE_HEX}",
        ).pack(anchor="w")
        entrada_asunto = tk.Entry(contenedor, font=("Segoe UI", 10), bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_TEXTO_HEX}", bd=0)
        entrada_asunto.pack(fill="x", ipady=6, pady=(2, 12))

        tk.Label(
            contenedor, text="MENSAJE (Variable):", font=("Segoe UI", 8, "bold"), bg=f"#{BEIGE_HEX}",
            fg=f"#{CAFE_SUAVE_HEX}",
        ).pack(anchor="w")
        frame_cuerpo = tk.Frame(contenedor, bg=f"#{AMARILLO_CLARO_HEX}")
        frame_cuerpo.pack(fill="both", expand=True, pady=(2, 15))
        
        texto_cuerpo = tk.Text(
            frame_cuerpo, font=("Segoe UI", 10), bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
            bd=0, padx=10, pady=10, wrap="word", insertbackground=f"#{CAFE_TEXTO_HEX}", height=8
        )
        scroll_cuerpo = tk.Scrollbar(frame_cuerpo, command=texto_cuerpo.yview)
        texto_cuerpo.configure(yscrollcommand=scroll_cuerpo.set)
        texto_cuerpo.pack(side="left", fill="both", expand=True)
        scroll_cuerpo.pack(side="right", fill="y")

        self.lbl_estado_envio_popup = tk.Label(
            contenedor, text="", font=("Segoe UI", 9, "bold"), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_SUAVE_HEX}",
        )
        self.lbl_estado_envio_popup.pack(fill="x", pady=(0, 10))

        frame_botones = tk.Frame(contenedor, bg=f"#{BEIGE_HEX}")
        frame_botones.pack(fill="x", side="bottom", pady=(5, 0)) 

        def confirmar_envio():
            destinatario_final = entrada_destinatario.get().strip()
            asunto_final = entrada_asunto.get().strip()
            cuerpo_final = texto_cuerpo.get("1.0", "end").strip()

            if not destinatario_final or not asunto_final or not cuerpo_final:
                messagebox.showerror("Datos Incompletos", "El destinatario, el asunto y el mensaje no pueden estar vacios.", parent=ventana)
                return

            self._despachar_correo_gmail(destinatario_final, asunto_final, cuerpo_final, ventana)

        def cancelar_envio():
            _al_cerrar()

        tk.Button(
            frame_botones, text="Enviar Correo", font=("Segoe UI", 10, "bold"), bg=f"#{AMARILLO_ACENTO_HEX}",
            fg=f"#{CAFE_TEXTO_HEX}", bd=0, cursor="hand2", activebackground="#E0B015",
            command=confirmar_envio, padx=15, pady=10,
        ).pack(side="left", fill="x", expand=True, padx=(0, 8))
        
        tk.Button(
            frame_botones, text="Cancelar", font=("Segoe UI", 10, "bold"), bg=f"#{BEIGE_OSCURO_HEX}",
            fg=f"#{CAFE_TEXTO_HEX}", bd=0, cursor="hand2", activebackground="#D9CBA0",
            command=cancelar_envio, padx=15, pady=10,
        ).pack(side="right", fill="x", expand=True, padx=(8, 0))

    def _despachar_correo_gmail(self, destinatario, asunto, cuerpo, ventana_origen=None):
        """Realiza el envio SMTP real y la verificacion en el servidor, actualizando
        el estado tanto en la ventana emergente como en la pantalla principal."""
        try:
            self.root.config(cursor="wait")
            self.lbl_estado_envio.config(text="Enviando correo...", fg=f"#{CAFE_SUAVE_HEX}")
            if hasattr(self, "lbl_estado_envio_popup"):
                self.lbl_estado_envio_popup.config(text="Enviando correo...", fg=f"#{CAFE_SUAVE_HEX}")
            self.root.update()
            enviar_correo_smtp(
                usuario=MI_GMAIL, contrasena=MI_PASSWORD_DE_APLICACION, destinatario=destinatario,
                asunto=asunto, cuerpo=cuerpo,
            )
            self.btn_verificar_servidor.config(state="normal")
            self.lbl_estado_envio.config(text="Verificando en el servidor...", fg=f"#{CAFE_SUAVE_HEX}")
            if hasattr(self, "lbl_estado_envio_popup"):
                self.lbl_estado_envio_popup.config(text="Verificando en el servidor...", fg=f"#{CAFE_SUAVE_HEX}")
            self.root.update()

            # Verificacion automatica inmediatamente despues del envio
            estado_servidor = comprobar_ultimo_enviado(MI_GMAIL, MI_PASSWORD_DE_APLICACION)
            self.root.config(cursor="")
            self.lbl_estado_envio.config(text="✔ Correo enviado y verificado en el servidor.", fg="#2E7D32")
            if hasattr(self, "lbl_estado_envio_popup"):
                self.lbl_estado_envio_popup.config(text="✔ Correo enviado y verificado en el servidor.", fg="#2E7D32")

            if ventana_origen is not None:
                try:
                    ventana_origen.destroy()
                except Exception:
                    pass
                self._ventana_envio_actual = None

            messagebox.showinfo(
                "Envio Verificado",
                "El correo electronico ha sido despachado correctamente y su presencia fue "
                f"confirmada en el servidor de Gmail.\n\nDetalle del servidor:\n{estado_servidor}",
            )
        except Exception as e:
            self.root.config(cursor="")
            self.lbl_estado_envio.config(text="✘ No se pudo completar el envio.", fg="#B23B3B")
            if hasattr(self, "lbl_estado_envio_popup"):
                self.lbl_estado_envio_popup.config(text="✘ No se pudo completar el envio.", fg="#B23B3B")
                messagebox.showerror("Error en Envio", f"No se pudo completar el envio SMTP.\nDetalle: {str(e)}")
        else:
            self.lbl_estado_envio.config(text="Envio cancelado por el usuario.", fg=f"#{CAFE_SUAVE_HEX}")
            messagebox.showinfo("Envio Cancelado", "El flujo de salida ha sido cancelado de forma segura.")

    def ejecutar_verificacion_servidor(self):
        self.root.config(cursor="wait")
        self.lbl_estado_envio.config(text="Verificando en el servidor...", fg=f"#{CAFE_SUAVE_HEX}")
        self.root.update()
        estado_servidor = comprobar_ultimo_enviado(MI_GMAIL, MI_PASSWORD_DE_APLICACION)
        self.root.config(cursor="")
        self.lbl_estado_envio.config(text="✔ Verificacion de servidor completada.", fg="#2E7D32")
        messagebox.showinfo("Rastreo de Servidor Gmail", estado_servidor)

    # ============================================================== OPCION 4 - VECTORES Y ARREGLOS
    def construir_pantalla_vectores(self):
        btn_back = tk.Button(
            self.frame_vectores, text="<-  Volver al Menu", font=("Segoe UI", 9, "bold"), bg=f"#{BEIGE_OSCURO_HEX}",
            fg=f"#{CAFE_TEXTO_HEX}", activebackground=f"#{AMARILLO_CLARO_HEX}", activeforeground=f"#{CAFE_TEXTO_HEX}",
            bd=0, cursor="hand2", padx=15, pady=6, command=self.mostrar_menu_principal,
        )
        btn_back.pack(anchor="w", pady=(0, 15))

        tk.Label(
            self.frame_vectores, text="Calculadora de Vectores y Arreglos Multidimensionales",
            font=("Segoe UI", 14, "bold"), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
        ).pack(anchor="w", pady=(0, 6))
        tk.Label(
            self.frame_vectores,
            text="Modulo de automatizacion algebraica con NumPy: crea matrices/vectores N-dimensionales "
                 "(separando filas con ';' y valores con ',') y aplica operaciones vectoriales al instante.",
            font=("Segoe UI", 9), bg=f"#{BEIGE_HEX}", fg=f"#{CAFE_SUAVE_HEX}", wraplength=760, justify="left",
        ).pack(anchor="w", pady=(0, 15))

        frame_inputs = tk.Frame(self.frame_vectores, bg=f"#{AMARILLO_CLARO_HEX}", padx=20, pady=18)
        frame_inputs.pack(fill="x", pady=10)

        tk.Label(
            frame_inputs, text="ARREGLO / VECTOR A   (ej: 1,2,3  o  1,2;3,4 para matriz)", font=("Segoe UI", 8, "bold"),
            bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_SUAVE_HEX}",
        ).pack(anchor="w")
        self.entry_vector_a = tk.Entry(
            frame_inputs, font=("Segoe UI", 11), bd=0, bg=f"#{BLANCO_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
            insertbackground=f"#{CAFE_TEXTO_HEX}",
        )
        self.entry_vector_a.pack(fill="x", ipady=7, pady=(4, 14))
        self.entry_vector_a.insert(0, "1,2,3")

        tk.Label(
            frame_inputs, text="ARREGLO / VECTOR B   (ej: 4,5,6  o  5,6;7,8 para matriz)", font=("Segoe UI", 8, "bold"),
            bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_SUAVE_HEX}",
        ).pack(anchor="w")
        self.entry_vector_b = tk.Entry(
            frame_inputs, font=("Segoe UI", 11), bd=0, bg=f"#{BLANCO_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
            insertbackground=f"#{CAFE_TEXTO_HEX}",
        )
        self.entry_vector_b.pack(fill="x", ipady=7, pady=(4, 14))
        self.entry_vector_b.insert(0, "4,5,6")

        tk.Label(
            frame_inputs, text="OPERACION", font=("Segoe UI", 8, "bold"), bg=f"#{AMARILLO_CLARO_HEX}",
            fg=f"#{CAFE_SUAVE_HEX}",
        ).pack(anchor="w")
        self.combo_operacion_vectorial = ttk.Combobox(
            frame_inputs, state="readonly", font=("Segoe UI", 10),
            values=[
                "Suma (A + B)", "Resta (A - B)", "Producto Punto (Dot)", "Producto Cruz (Cross, solo 3D)",
                "Multiplicacion Matricial (A x B)", "Transpuesta de A", "Determinante de A",
                "Norma / Magnitud de A", "Matriz Inversa de A", "Angulo entre A y B (grados)",
            ],
        )
        self.combo_operacion_vectorial.pack(fill="x", ipady=4, pady=(4, 16))
        self.combo_operacion_vectorial.current(0)

        tk.Button(
            frame_inputs, text="Ejecutar Operacion", font=("Segoe UI", 10, "bold"), bg=f"#{AMARILLO_ACENTO_HEX}",
            fg=f"#{CAFE_TEXTO_HEX}", activebackground="#E0B015", bd=0, cursor="hand2",
            command=self.ejecutar_operacion_vectorial,
        ).pack(fill="x", ipady=10)

        tk.Label(
            self.frame_vectores, text="RESULTADO", font=("Segoe UI", 8, "bold"), bg=f"#{BEIGE_HEX}",
            fg=f"#{CAFE_SUAVE_HEX}",
        ).pack(anchor="w", pady=(16, 4))
        self.txt_resultado_vectorial = tk.Text(
            self.frame_vectores, height=10, bg=f"#{BLANCO_HEX}", fg=f"#{CAFE_TEXTO_HEX}", font=("Consolas", 10),
            bd=0, padx=18, pady=16, highlightthickness=1, highlightbackground=f"#{BEIGE_OSCURO_HEX}",
        )
        self.txt_resultado_vectorial.pack(fill="both", expand=True)
        self.txt_resultado_vectorial.insert("1.0", "El resultado numerico de la operacion aparecera aqui...")
        self.txt_resultado_vectorial.config(state="disabled")

    def _parsear_arreglo(self, texto):
        """Convierte 'a,b,c' en vector 1D o 'a,b;c,d' en matriz 2D usando NumPy."""
        texto = texto.strip()
        filas = [f for f in texto.split(";") if f.strip() != ""]
        matriz = [[float(v.strip()) for v in fila.split(",") if v.strip() != ""] for fila in filas]
        arreglo = np.array(matriz[0]) if len(matriz) == 1 else np.array(matriz)
        return arreglo

    def ejecutar_operacion_vectorial(self):
        try:
            a = self._parsear_arreglo(self.entry_vector_a.get())
            b = self._parsear_arreglo(self.entry_vector_b.get())
            operacion = self.combo_operacion_vectorial.get()

            if operacion.startswith("Suma"):
                resultado = np.add(a, b)
            elif operacion.startswith("Resta"):
                resultado = np.subtract(a, b)
            elif operacion.startswith("Producto Punto"):
                resultado = np.dot(a, b)
            elif operacion.startswith("Producto Cruz"):
                resultado = np.cross(a, b)
            elif operacion.startswith("Multiplicacion Matricial"):
                resultado = np.matmul(a, b)
            elif operacion.startswith("Transpuesta"):
                resultado = np.transpose(a)
            elif operacion.startswith("Determinante"):
                resultado = np.linalg.det(a)
            elif operacion.startswith("Norma"):
                resultado = np.linalg.norm(a)
            elif operacion.startswith("Matriz Inversa"):
                resultado = np.linalg.inv(a)
            elif operacion.startswith("Angulo"):
                coseno = np.dot(a, b) / (np.linalg.norm(a) * np.linalg.norm(b))
                resultado = np.degrees(np.arccos(np.clip(coseno, -1.0, 1.0)))
            else:
                resultado = "Operacion no reconocida."

            texto_resultado = (
                f"Arreglo A (forma {np.shape(a)}):\n{a}\n\n"
                f"Arreglo B (forma {np.shape(b)}):\n{b}\n\n"
                f"Operacion aplicada: {operacion}\n\n"
                f"RESULTADO:\n{resultado}"
            )
        except Exception as e:
            texto_resultado = f"No se pudo calcular la operacion.\nDetalle tecnico: {str(e)}"

        self.txt_resultado_vectorial.config(state="normal")
        self.txt_resultado_vectorial.delete("1.0", "end")
        self.txt_resultado_vectorial.insert("1.0", texto_resultado)
        self.txt_resultado_vectorial.config(state="disabled")

    # ============================================================== SOBRE NOSOTROS
    def construir_pantalla_sobre_nosotros(self):
        wrapper = tk.Frame(self.frame_sobre_nosotros, bg=f"#{BLANCO_HEX}")
        wrapper.pack(fill="both", expand=True)

        self.foto_campus_sobre_nosotros = None
        ruta_campus = os.path.join(RUTA_BASE, "campus_unach.png")
        if PIL_DISPONIBLE and os.path.exists(ruta_campus):
            try:
                imagen_campus = Image.open(ruta_campus).convert("RGB")
                imagen_campus = imagen_campus.resize((680, 260), Image.LANCZOS)
                self.foto_campus_sobre_nosotros = ImageTk.PhotoImage(imagen_campus)
                tk.Label(wrapper, image=self.foto_campus_sobre_nosotros, bg=f"#{BLANCO_HEX}").pack(pady=(20, 0))
            except Exception:
                pass

        contenido = tk.Frame(wrapper, bg=f"#{BLANCO_HEX}")
        contenido.place(relx=0.5, rely=0.62, anchor="center", width=680)

        pill = tk.Label(
            contenido, text="SOBRE NOSOTROS", font=("Segoe UI", 9, "bold"), bg=f"#{AMARILLO_CLARO_HEX}",
            fg=f"#{CAFE_SUAVE_HEX}", padx=14, pady=5,
        )
        pill.pack(pady=(0, 16))

        tk.Label(
            contenido, text=f"El equipo detras de {NOMBRE_CARRERA}", font=("Segoe UI", 20, "bold"),
            bg=f"#{BLANCO_HEX}", fg=f"#{CAFE_TEXTO_HEX}", wraplength=640, justify="center",
        ).pack(pady=(0, 6))

        tk.Frame(contenido, bg=f"#{AMARILLO_ACENTO_HEX}", width=70, height=4).pack(pady=(0, 20))

        texto_sobre_nosotros = (
            "Somos un grupo dinamico de estudiantes unidos por la pasion hacia la tecnologia y el "
            "desarrollo de software. Mas alla de las aulas, nos hemos consolidado como un equipo "
            "colaborativo que busca transformar ideas innovadoras en soluciones digitales reales, "
            "enfrentando cada desafio tecnico como una oportunidad para aprender y crecer juntos. "
            "Nos caracteriza el pensamiento critico, la curiosidad constante por las nuevas herramientas "
            "y el compromiso con la excelencia en el codigo que escribimos, siempre apoyandonos en la "
            "experimentacion y el trabajo en equipo para impulsar nuestro desarrollo profesional y creativo."
        )
        tk.Label(
            contenido, text=texto_sobre_nosotros, font=("Segoe UI", 11), bg=f"#{BLANCO_HEX}",
            fg=f"#{CAFE_TEXTO_HEX}", wraplength=640, justify="left",
        ).pack(pady=(0, 10))

    # ============================================================== INTEGRANTES
    def construir_pantalla_integrantes(self):
        encabezado = tk.Frame(self.frame_integrantes, bg=f"#{BLANCO_HEX}")
        encabezado.pack(fill="x", pady=(20, 10), padx=20)

        pill = tk.Label(
            encabezado, text="INTEGRANTES", font=("Segoe UI", 9, "bold"), bg=f"#{AMARILLO_CLARO_HEX}",
            fg=f"#{CAFE_SUAVE_HEX}", padx=14, pady=5,
        )
        pill.pack()
        tk.Label(
            encabezado, text="El equipo registrado en el sistema", font=("Segoe UI", 18, "bold"),
            bg=f"#{BLANCO_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
        ).pack(pady=(8, 4))
        tk.Label(
            encabezado, text="Estudiantes y colaboradores que se han registrado en la plataforma.",
            font=("Segoe UI", 10), bg=f"#{BLANCO_HEX}", fg=f"#{GRIS_PLACEHOLDER}",
        ).pack()

        # Area con scroll para las tarjetas de integrantes
        area_scroll = tk.Frame(self.frame_integrantes, bg=f"#{BLANCO_HEX}")
        area_scroll.pack(fill="both", expand=True, padx=20, pady=(10, 20))

        self.canvas_integrantes = tk.Canvas(area_scroll, bg=f"#{BLANCO_HEX}", highlightthickness=0)
        scrollbar = tk.Scrollbar(area_scroll, orient="vertical", command=self.canvas_integrantes.yview)
        self.frame_grid_integrantes = tk.Frame(self.canvas_integrantes, bg=f"#{BLANCO_HEX}")

        self.frame_grid_integrantes.bind(
            "<Configure>",
            lambda e: self.canvas_integrantes.configure(scrollregion=self.canvas_integrantes.bbox("all")),
        )
        self.canvas_integrantes.create_window((0, 0), window=self.frame_grid_integrantes, anchor="nw")
        self.canvas_integrantes.configure(yscrollcommand=scrollbar.set)

        self.canvas_integrantes.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.fotos_integrantes_cache = []

    def refrescar_integrantes(self):
        for widget in self.frame_grid_integrantes.winfo_children():
            widget.destroy()
        self.fotos_integrantes_cache = []

        usuarios = cargar_usuarios()
        columnas = 4
        for indice, (usuario, datos) in enumerate(usuarios.items()):
            fila, columna = divmod(indice, columnas)
            tarjeta = tk.Frame(
                self.frame_grid_integrantes, bg=f"#{AMARILLO_CLARO_HEX}", padx=14, pady=14, highlightthickness=1,
                highlightbackground=f"#{BEIGE_OSCURO_HEX}",
            )
            tarjeta.grid(row=fila, column=columna, padx=12, pady=12, sticky="n")

            foto_widget = self._obtener_foto_integrante(datos.get("foto"))
            if foto_widget is not None:
                self.fotos_integrantes_cache.append(foto_widget)
                tk.Label(tarjeta, image=foto_widget, bg=f"#{AMARILLO_CLARO_HEX}").pack()
            else:
                inicial = (datos.get("nombre") or usuario)[0].upper()
                tk.Label(
                    tarjeta, text=inicial, font=("Segoe UI", 30, "bold"), bg=f"#{BEIGE_OSCURO_HEX}",
                    fg=f"#{CAFE_TEXTO_HEX}", width=6, height=3,
                ).pack()

            tk.Label(
                tarjeta, text=datos.get("nombre", usuario), font=("Segoe UI", 10, "bold"),
                bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_TEXTO_HEX}", wraplength=140,
            ).pack(pady=(8, 0))
            tk.Label(
                tarjeta, text=f"@{usuario}", font=("Segoe UI", 8), bg=f"#{AMARILLO_CLARO_HEX}",
                fg=f"#{CAFE_SUAVE_HEX}",
            ).pack()

        if not usuarios:
            tk.Label(
                self.frame_grid_integrantes, text="Aun no hay integrantes registrados.",
                font=("Segoe UI", 10), bg=f"#{BLANCO_HEX}", fg=f"#{GRIS_PLACEHOLDER}",
            ).grid(row=0, column=0, pady=30)

    def _obtener_foto_integrante(self, ruta_relativa):
        if not ruta_relativa or not PIL_DISPONIBLE:
            return None
        ruta_absoluta = os.path.join(RUTA_BASE, ruta_relativa)
        if not os.path.exists(ruta_absoluta):
            return None
        try:
            imagen = Image.open(ruta_absoluta).convert("RGB")
            imagen = imagen.resize((90, 120), Image.LANCZOS)
            return ImageTk.PhotoImage(imagen)
        except Exception:
            return None

    # CONTACTO
    def construir_pantalla_contacto(self):
        contenido = tk.Frame(self.frame_contacto, bg=f"#{BLANCO_HEX}")
        contenido.place(relx=0.5, rely=0.5, anchor="center", width=520)

        tk.Label(
            contenido, text="CONTACTO", font=("Segoe UI", 9, "bold"), bg=f"#{AMARILLO_CLARO_HEX}",
            fg=f"#{CAFE_SUAVE_HEX}", padx=14, pady=5,
        ).pack(pady=(0, 16))
        tk.Label(
            contenido, text="Hablemos", font=("Segoe UI", 20, "bold"), bg=f"#{BLANCO_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
        ).pack(pady=(0, 6))
        tk.Frame(contenido, bg=f"#{AMARILLO_ACENTO_HEX}", width=70, height=4).pack(pady=(0, 20))
        tk.Label(
            contenido,
            text=(
                f"{NOMBRE_CARRERA}\n{NOMBRE_UNIVERSIDAD}\n\n"
                f"Correo institucional: {MI_GMAIL}\n"
                "Escribenos ante cualquier duda sobre el sistema o para unirte al equipo."
            ),
            font=("Segoe UI", 11), bg=f"#{BLANCO_HEX}", fg=f"#{CAFE_TEXTO_HEX}", justify="center",
        ).pack()

    # ============================================================== NAVEGACION CON CONTENIDO REAL
    def _encabezado_seccion(self, frame, etiqueta, titulo, subtitulo):
        contenedor = tk.Frame(frame, bg=f"#{BLANCO_HEX}")
        contenedor.pack(fill="x", pady=(24, 10), padx=30)
        tk.Label(
            contenedor, text=etiqueta, font=("Segoe UI", 9, "bold"), bg=f"#{AMARILLO_CLARO_HEX}",
            fg=f"#{CAFE_SUAVE_HEX}", padx=14, pady=5,
        ).pack(anchor="w")
        tk.Label(
            contenedor, text=titulo, font=("Segoe UI", 19, "bold"), bg=f"#{BLANCO_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
        ).pack(anchor="w", pady=(10, 4))
        tk.Frame(contenedor, bg=f"#{AMARILLO_ACENTO_HEX}", width=70, height=4).pack(anchor="w", pady=(0, 10))
        tk.Label(
            contenedor, text=subtitulo, font=("Segoe UI", 10), bg=f"#{BLANCO_HEX}", fg=f"#{GRIS_PLACEHOLDER}",
            wraplength=760, justify="left",
        ).pack(anchor="w")
        return contenedor

    def _area_scroll(self, frame):
        area = tk.Frame(frame, bg=f"#{BLANCO_HEX}")
        area.pack(fill="both", expand=True, padx=30, pady=(10, 20))
        canvas = tk.Canvas(area, bg=f"#{BLANCO_HEX}", highlightthickness=0)
        scrollbar = tk.Scrollbar(area, orient="vertical", command=canvas.yview)
        rejilla = tk.Frame(canvas, bg=f"#{BLANCO_HEX}")
        rejilla.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=rejilla, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        return rejilla

    def _tarjeta(self, parent, fila, columna, titulo, texto, etiqueta_estado, color_estado):
        tarjeta = tk.Frame(
            parent, bg=f"#{AMARILLO_CLARO_HEX}", padx=18, pady=16, highlightthickness=1,
            highlightbackground=f"#{BEIGE_OSCURO_HEX}", width=340, height=190,
        )
        tarjeta.grid(row=fila, column=columna, padx=12, pady=12, sticky="n")
        tarjeta.pack_propagate(False)
        tk.Label(
            tarjeta, text=etiqueta_estado, font=("Segoe UI", 8, "bold"), bg=color_estado, fg=f"#{BLANCO_HEX}",
            padx=10, pady=3,
        ).pack(anchor="w")
        tk.Label(
            tarjeta, text=titulo, font=("Segoe UI", 12, "bold"), bg=f"#{AMARILLO_CLARO_HEX}",
            fg=f"#{CAFE_TEXTO_HEX}", wraplength=300, justify="left",
        ).pack(anchor="w", pady=(8, 6))
        tk.Label(
            tarjeta, text=texto, font=("Segoe UI", 9), bg=f"#{AMARILLO_CLARO_HEX}", fg=f"#{CAFE_SUAVE_HEX}",
            wraplength=300, justify="left",
        ).pack(anchor="w")

    def construir_pantalla_lineas(self):
        frame = self.frames_stub["lineas"]
        self._encabezado_seccion(
            frame, "LINEAS DE INVESTIGACION", "Nuestras lineas de investigacion",
            "Ejes tematicos que orientan los proyectos de titulacion, semilleros y automatizaciones "
            "desarrolladas por el equipo dentro de la carrera de Ingenieria en Ciencia de Datos e IA.",
        )
        rejilla = self._area_scroll(frame)
        lineas = [
            ("Inteligencia Artificial Aplicada", "Modelos de lenguaje, vision por computador y agentes "
             "automatizados aplicados a procesos administrativos y academicos (UNACH, 2024)."),
            ("Ciencia de Datos y Analitica", "Analisis estadistico, mineria de datos y visualizacion de "
             "encuestas institucionales para la toma de decisiones basada en evidencia."),
            ("Automatizacion de Procesos (RPA)", "Flujos de trabajo que integran correo electronico, "
             "documentos y hojas de calculo bajo un mismo sistema inteligente."),
            ("Software Educativo", "Herramientas digitales orientadas a la gestion academica y a la "
             "formacion continua de estudiantes universitarios."),
        ]
        for idx, (titulo, texto) in enumerate(lineas):
            fila, columna = divmod(idx, 2)
            self._tarjeta(rejilla, fila, columna, titulo, texto, "LINEA ACTIVA", f"#{AMARILLO_ACENTO_HEX}")

    def construir_pantalla_academy(self):
        frame = self.frames_stub["academy"]
        self._encabezado_seccion(
            frame, "ACADEMY", "UNACH Academy",
            "Espacio de formacion continua con micro-cursos internos sobre las tecnologias utilizadas "
            "en la Suite Central de Automatizacion.",
        )
        rejilla = self._area_scroll(frame)
        cursos = [
            ("Fundamentos de Python para Automatizacion", "Introduccion a variables, funciones y librerias "
             "usadas en el sistema (Pandas, NumPy, OpenPyXL)."),
            ("Interfaces Graficas con Tkinter", "Diseno de interfaces de escritorio profesionales aplicando "
             "identidad visual institucional."),
            ("Estadistica Descriptiva Aplicada", "Calculo de moda, mediana y varianza para el analisis de "
             "encuestas dentro del modulo analitico."),
            ("Integracion de IA Generativa", "Consumo de APIs de modelos de lenguaje para clasificacion y "
             "redaccion automatica de respuestas."),
        ]
        for idx, (titulo, texto) in enumerate(cursos):
            fila, columna = divmod(idx, 2)
            self._tarjeta(rejilla, fila, columna, titulo, texto, "DISPONIBLE", f"#{CAFE_SUAVE_HEX}")

    #  SECCIONES EN CONSTRUCCION (respaldo)
    def construir_pantalla_stub(self, frame, titulo):
        contenido = tk.Frame(frame, bg=f"#{BLANCO_HEX}")
        contenido.place(relx=0.5, rely=0.5, anchor="center")
        tk.Label(
            contenido, text=titulo, font=("Segoe UI", 20, "bold"), bg=f"#{BLANCO_HEX}", fg=f"#{CAFE_TEXTO_HEX}",
        ).pack(pady=(0, 10))
        tk.Frame(contenido, bg=f"#{AMARILLO_ACENTO_HEX}", width=70, height=4).pack(pady=(0, 16))
        tk.Label(
            contenido, text="Contenido en construccion. Muy pronto disponible.",
            font=("Segoe UI", 11), bg=f"#{BLANCO_HEX}", fg=f"#{GRIS_PLACEHOLDER}",
        ).pack()

    #  SALIDA
    def solicitar_feedback_salida(self):
        VentanaCalificacion(self.root, on_close_callback=self.root.quit)


if __name__ == "__main__":
    def iniciar_app_principal():
        """Se ejecuta tras un login exitoso: limpia la ventana de acceso
        y lanza la Suite Central de Automatizacion en el mismo root."""
        for widget in root.winfo_children():
            widget.destroy()
        AppMenuGlobal(root)

    root = tk.Tk()
    VentanaLogin(root, on_login_exitoso=iniciar_app_principal)
    root.mainloop()