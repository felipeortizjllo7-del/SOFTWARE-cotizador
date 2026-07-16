# -*- coding: utf-8 -*-
"""
Cotizador INNOBA Colombia DMC
- Itinerario combinado: hasta 5 destinos en una sola cotizacion.
- Tarifas 2026. Precios en USD. TRM en vivo (dolar-colombia.com, -100), NO se muestra.
- Descripciones de tours (TARIFARIO - TOURS 2026).
- Al generar el PDF, se envia por correo (Office 365) al email del cliente.
"""
import os
import sys
import re
import ssl
import json
import shutil
import smtplib
import difflib
import tempfile
import datetime
import calendar
import threading
import unicodedata
import urllib.request
from email.message import EmailMessage
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
from fpdf import FPDF
from PIL import Image


# ============================================================================
# Rutas
# ============================================================================
def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def recurso(nombre):
    base = getattr(sys, "_MEIPASS", app_dir())
    return os.path.join(base, nombre)

def datos_dir():
    """Carpeta ESCRIBIBLE para datos del usuario (config, clientes).
       Necesaria cuando el programa se instala en Program Files."""
    base = os.path.join(os.environ.get("APPDATA") or app_dir(), "CotizadorInnoba")
    try:
        os.makedirs(base, exist_ok=True)
    except Exception:
        base = app_dir()
    return base

CONFIG_PATH = os.path.join(datos_dir(), "config_empresa.json")


# ============================================================================
# Version del software y actualizaciones automaticas
# ============================================================================
# IMPORTANTE: este numero se incrementa en cada ajuste (lo hace publicar_version.py).
# Esquema resumido de 2 digitos: 1.0 -> 1.1 -> ... -> 1.9 -> 2.0
VERSION = "3.0"
GITHUB_OWNER = "felipeortizjllo7-del"
GITHUB_REPO = "SOFTWARE-cotizador"
# Webhook (Google Apps Script /exec) por donde el HTML de los clientes envia sus
# cotizaciones; el .exe las importa aqui.
WEBHOOK_URL = "https://script.google.com/macros/s/AKfycby0tbuYavMW7dl5cah7qfIsJVM3hOmt6Sh6h4M2ZQD4l7ncIGxyCdHK2w3ogny0o3oWVQ/exec"
# Archivo con la ultima version publicada (rama main del repositorio)
UPDATE_URL = (f"https://raw.githubusercontent.com/{GITHUB_OWNER}/{GITHUB_REPO}"
              f"/main/version.json")

def _ver_tuple(s):
    nums = re.findall(r"\d+", str(s or "0"))
    return tuple(int(x) for x in nums[:3]) if nums else (0,)

def obtener_version_remota():
    """Lee la ultima version publicada. Primero la API de contenidos (sin cache del
       CDN, deteccion instantanea); si falla, el archivo raw."""
    ctx = ssl.create_default_context()
    fuentes = [
        (f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}"
         f"/contents/version.json?ref=main", "application/vnd.github.raw"),
        (UPDATE_URL, "*/*"),
    ]
    for url, accept in fuentes:
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": "CotizadorInnoba", "Accept": accept,
                "Cache-Control": "no-cache"})
            with urllib.request.urlopen(req, context=ctx, timeout=12) as r:
                return json.loads(r.read().decode("utf-8"))
        except Exception:
            continue
    return None

def hay_actualizacion():
    """Devuelve el dict de la version remota si es MAYOR a la instalada, si no None."""
    info = obtener_version_remota()
    if info and _ver_tuple(info.get("version", "0")) > _ver_tuple(VERSION):
        return info
    return None


# ============================================================================
# Paleta de marca
# ============================================================================
NAVY = "#013984"; NAVY2 = "#00285F"; BLUE = "#1466C7"; BLUE_H = "#0F4FA0"
CYAN = "#2E8BE6"; BG = "#EEF3FA"; CARD = "#FFFFFF"; CARD2 = "#F4F8FD"
TEXT = "#16233D"; MUTED = "#64748B"; GREEN = "#1E9E5A"; GREEN_H = "#178049"
LINE = "#D7E1EF"; RED = "#C0392B"


# ============================================================================
# Tasa de cambio (TRM en vivo)
# ============================================================================
TRM_URL = "https://www.dolar-colombia.com/"
# Fuente OFICIAL de la TRM (Superfinanciera via datos.gov.co): estable y con CORS.
TRM_API = ("https://www.datos.gov.co/resource/32sa-8pi3.json"
           "?$order=vigenciadesde%20DESC&$limit=1")
DESCUENTO_DOLAR = 100.0   # descuento por defecto

def _trm_valida(v):
    """Acepta solo valores de TRM con sentido (evita precios negativos por un
       parseo malo). La TRM COP/USD ronda 3000-5000; damos margen amplio."""
    try:
        v = float(str(v).replace(",", ""))
    except (TypeError, ValueError):
        return None
    return v if 1000.0 <= v <= 20000.0 else None

# Quien realiza la cotizacion (nombre, cargo) -> firma del PDF
COTIZADORES = [
    ("Felipe Ortiz Jaramillo", "Gerente - Innoba DMC"),
    ("Carlos Ortiz Jaramillo", "Gerente Comercial - Innoba DMC"),
]

def periodo_por_fecha(fecha):
    """Segun la fecha de IDA -> (descuento_pesos, margen_hotel, margen_terrestre).
       margen None = usar el margen normal de cada hoja. Reglas:
       - 2027 (todo el ano): -300, margenes 0.82 / 0.69
       - sep-dic 2026: -200, margenes normales
       - resto (por defecto): -100, margenes normales."""
    if fecha:
        y, m = fecha.year, fecha.month
        if y >= 2027:
            return 300.0, 0.82, 0.69
        if y == 2026 and 9 <= m <= 12:
            return 200.0, None, None
    return 100.0, None, None

def obtener_trm(timeout=12):
    """Devuelve la TRM del dia (COP por USD) o None. Primero la fuente oficial;
       si falla, dolar-colombia.com. Siempre validando el rango."""
    # 1) Fuente oficial (datos.gov.co) - estable y diaria
    for ctx in (ssl.create_default_context(), ssl._create_unverified_context()):
        try:
            req = urllib.request.Request(TRM_API,
                                         headers={"User-Agent": "CotizadorInnoba"})
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as r:
                data = json.loads(r.read().decode("utf-8"))
            v = _trm_valida(data[0].get("valor")) if data else None
            if v:
                return v
        except Exception:
            continue
    # 2) Respaldo: dolar-colombia.com
    req = urllib.request.Request(
        TRM_URL, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
    html = None
    for ctx in (ssl.create_default_context(), ssl._create_unverified_context()):
        try:
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as r:
                html = r.read().decode("utf-8", "replace")
            break
        except Exception:
            continue
    if not html:
        return None
    # tomar el primer numero con formato de miles (ej. 3,252.11) y validarlo
    for m in re.finditer(r"(\d{1,2}[.,]\d{3}[.,]\d{2})", html):
        v = _trm_valida(m.group(1).replace(",", ""))
        if v:
            return v
    return None


# ============================================================================
# Configuracion de la empresa
# ============================================================================
DEFAULT_CONFIG = {
    "empresa": "INNOBA Colombia DMC",
    "nit": "", "direccion": "", "telefono": "", "email": "", "web": "",
    "logo": "",
    "firma_nombre": "Felipe Ortiz",
    "firma_cargo": "Gerente - INNOBA Colombia DMC",
    "correo_remitente": "",
    "smtp_servidor": "smtp.office365.com",
    "smtp_puerto": "587",
    "smtp_password": "",
    "ultima_trm": "", "ultima_trm_fecha": "",
    "notas": ("Tarifas sujetas a disponibilidad al momento de la reserva. "
              "Precios en dolares americanos (USD) por el total indicado."),
}

def cargar_config():
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception:
            pass
    return cfg

def guardar_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ---- Clientes / Empresas (con sus vendedores) ----
CLIENTES_PATH = os.path.join(datos_dir(), "clientes.json")

def _sembrar_clientes():
    """Primera apertura tras instalar: copia la base de clientes incluida
       en el paquete hacia la carpeta escribible del usuario."""
    if os.path.exists(CLIENTES_PATH):
        return
    for semilla in (recurso("clientes.json"),
                    os.path.join(app_dir(), "clientes.json")):
        if os.path.exists(semilla):
            try:
                shutil.copyfile(semilla, CLIENTES_PATH)
            except Exception:
                pass
            return

def cargar_clientes():
    _sembrar_clientes()
    if os.path.exists(CLIENTES_PATH):
        try:
            with open(CLIENTES_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return []

def guardar_clientes(lst):
    with open(CLIENTES_PATH, "w", encoding="utf-8") as f:
        json.dump(lst, f, ensure_ascii=False, indent=2)


# ---- Historial de cotizaciones (con consecutivo) ----
COTIZACIONES_PATH = os.path.join(datos_dir(), "cotizaciones.json")

def cargar_cotizaciones():
    if os.path.exists(COTIZACIONES_PATH):
        try:
            with open(COTIZACIONES_PATH, "r", encoding="utf-8") as f:
                d = json.load(f)
            if isinstance(d, dict) and "items" in d:
                return d
        except Exception:
            pass
    return {"seq": 0, "items": []}

def guardar_cotizaciones(data):
    with open(COTIZACIONES_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def respaldar_datos(cfg):
    """Guarda copias de seguridad de cotizaciones/clientes en %APPDATA%\\...\\respaldos.
       Siempre deja un respaldo '_ultimo'; y si la version cambio (hubo una
       actualizacion) deja una copia permanente por version, para que las
       cotizaciones NUNCA se pierdan al actualizar."""
    try:
        carpeta = os.path.join(datos_dir(), "respaldos")
        os.makedirs(carpeta, exist_ok=True)
        cambio_version = cfg.get("ultima_version_vista", "") != VERSION
        hoy = datetime.date.today().strftime("%Y%m%d")
        for nombre in ("cotizaciones.json", "clientes.json", "config_empresa.json"):
            src = os.path.join(datos_dir(), nombre)
            if not os.path.exists(src):
                continue
            base = nombre.rsplit(".", 1)[0]
            shutil.copyfile(src, os.path.join(carpeta, base + "_ultimo.json"))
            if cambio_version:
                shutil.copyfile(
                    src, os.path.join(carpeta, f"{base}_v{VERSION}_{hoy}.json"))
        if cambio_version:
            cfg["ultima_version_vista"] = VERSION
            try:
                guardar_config(cfg)
            except Exception:
                pass
    except Exception:
        pass

def peek_numero_cotizacion():
    """Devuelve el proximo consecutivo (sin reservarlo aun), ej. COT-00001."""
    return f"COT-{cargar_cotizaciones().get('seq', 0) + 1:05d}"

def registrar_cotizacion(rec):
    """Asigna el consecutivo, guarda el registro y devuelve el numero asignado."""
    data = cargar_cotizaciones()
    data["seq"] = int(data.get("seq", 0)) + 1
    numero = f"COT-{data['seq']:05d}"
    rec = dict(rec); rec["numero"] = numero
    data["items"].append(rec)
    guardar_cotizaciones(data)
    return numero

def importar_cotizaciones_html():
    """Trae las cotizaciones creadas por clientes en el HTML (via WEBHOOK_URL) y
       las agrega al historial local sin duplicar. Devuelve cuantas nuevas."""
    if not WEBHOOK_URL:
        return 0
    try:
        ctx = ssl.create_default_context()
        req = urllib.request.Request(WEBHOOK_URL, headers={"User-Agent": "CotizadorInnoba"})
        with urllib.request.urlopen(req, context=ctx, timeout=20) as r:
            remotas = json.loads(r.read().decode("utf-8"))
    except Exception:
        return 0
    if isinstance(remotas, dict):
        remotas = remotas.get("items", [])
    if not isinstance(remotas, list):
        return 0
    data = cargar_cotizaciones()
    existentes = {str(it.get("web_id")) for it in data["items"] if it.get("web_id")}
    nuevas = 0
    for rc in remotas:
        wid = str(rc.get("id") or rc.get("web_id") or "")
        if not wid or wid in existentes:
            continue
        dests = rc.get("destinos", [])
        if isinstance(dests, str):
            dests = [d.strip() for d in dests.split(",") if d.strip()]
        try:
            total = float(rc.get("total", 0) or 0)
        except (TypeError, ValueError):
            total = 0.0
        data["seq"] = int(data.get("seq", 0)) + 1
        data["items"].append({
            "numero": f"COT-{data['seq']:05d}", "web_id": wid, "origen": "HTML (cliente)",
            "cliente": rc.get("cliente", ""), "asesor": rc.get("asesor", ""),
            "asesor_tel": rc.get("asesor_tel", ""), "email": rc.get("email", ""),
            "fecha": rc.get("fecha", ""), "fechas_viaje": rc.get("fechas_viaje", ""),
            "destinos": dests, "total": total, "estado": "Pendiente", "pdf": ""})
        existentes.add(wid); nuevas += 1
    if nuevas:
        guardar_cotizaciones(data)
    return nuevas

def parse_fecha(s):
    """Convierte 'dd/mm/aaaa' (o similares) a date; None si no se puede."""
    s = (s or "").strip()
    if not s:
        return None
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None

def seguimientos_pendientes(data=None):
    """Cotizaciones (no cerradas) cuya fecha de seguimiento o alguna tarea ya
       vencio (<= hoy). Devuelve lista de (item, motivo)."""
    data = data or cargar_cotizaciones()
    hoy = datetime.date.today()
    res = []
    for it in data.get("items", []):
        if it.get("estado") in ("Ganada", "Perdida"):
            continue
        motivos = []
        fs = parse_fecha(it.get("fecha_seg", ""))
        if fs and fs <= hoy:
            motivos.append(f"seguimiento {it.get('fecha_seg')}")
        for t in it.get("tareas", []):
            if t.get("hecha"):
                continue
            ft = parse_fecha(t.get("fecha", ""))
            if ft and ft <= hoy:
                motivos.append(f"tarea: {t.get('texto','')[:30]} ({t.get('fecha')})")
        if motivos:
            res.append((it, "; ".join(motivos)))
    return res

def exportar_clientes_excel(lst, ruta):
    import openpyxl
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Clientes"
    ws.append(["Empresa", "NIT/Documento", "Telefono", "Email", "Sitio web", "Pais",
               "Vendedor", "Vendedor telefono", "Vendedor email", "Vendedor cargo"])
    for c in lst:
        vends = c.get("vendedores") or [{}]
        for v in vends:
            ws.append([c.get("empresa", ""), c.get("nit", ""), c.get("telefono", ""),
                       c.get("email", ""), c.get("web", ""), c.get("pais", ""),
                       v.get("nombre", ""), v.get("telefono", ""), v.get("email", ""),
                       v.get("cargo", "")])
    wb.save(ruta)

def _nz(s):
    s = "".join(ch for ch in unicodedata.normalize("NFD", str(s or ""))
                if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", s).lower().strip()

def _leer_filas(ruta):
    """Lee filas de un .xlsx o .csv (detecta codificacion y delimitador)."""
    ext = os.path.splitext(ruta)[1].lower()
    if ext in (".csv", ".txt"):
        import csv, io
        raw = open(ruta, "rb").read()
        texto = None
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                texto = raw.decode(enc); break
            except UnicodeDecodeError:
                continue
        muestra = texto[:2000]
        delim = ";" if muestra.count(";") >= muestra.count(",") else ","
        return list(csv.reader(io.StringIO(texto), delimiter=delim))
    import openpyxl
    wb = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
    return [row for row in wb.active.iter_rows(values_only=True)]

def importar_clientes_excel(ruta):
    """Importa clientes de un Excel/CSV. Reconoce nuestro formato y exportaciones
       tipo Bitrix (Nombre + Apellido + Compania)."""
    filas = _leer_filas(ruta)
    if not filas:
        return []
    header = [_nz(x) for x in filas[0]]
    def col(keys, exclude=()):
        for i, h in enumerate(header):
            if any(k in h for k in keys) and not any(e in h for e in exclude):
                return i
        return None
    i_comp = col(["compania", "empresa", "company", "razon social"])
    i_cli = col(["cliente"], exclude=["tipo"])
    i_emp = i_comp if i_comp is not None else i_cli
    i_nom = col(["nombre"], exclude=["compania", "empresa", "segundo"])
    i_ape = col(["apellido"])
    i_nit = col(["nit", "documento", "ruc", "id fiscal", "identificacion"], exclude=["tipo"])
    i_tel = col(["telefono del trabajo", "telefono trabajo", "telefono", "phone", "celular", "movil"],
                exclude=["vendedor", "sms", "localizador", "fax", "casa"])
    i_email = col(["e-mail del trabajo", "email del trabajo", "email", "correo", "e-mail"],
                  exclude=["vendedor", "boletines", "casa"])
    i_web = col(["sitio web", "web", "url"], exclude=["facebook", "vk"])
    i_pais = col(["pais", "country"], exclude=["codigo"])
    i_cargo = col(["cargo", "rol", "puesto"])
    i_vnom = col(["vendedor", "asesor"], exclude=["telefono", "email", "correo", "cargo"])
    i_vtel = col(["vendedor telefono", "telefono vendedor"])
    i_vemail = col(["vendedor email", "vendedor correo", "email vendedor"])
    if i_emp is None and i_nom is None:
        i_emp = 0
    def val(row, i):
        return str(row[i]).strip() if (i is not None and i < len(row) and row[i] is not None) else ""

    def es_junk(t):
        t = t.lower()
        return any(k in t for k in ("no-reply", "noreply", "no_reply", "bitrix24",
                                    "google-noreply", "@google.com", "imol|facebook"))

    empresas = {}
    orden = []
    for row in filas[1:]:
        if not any(row):
            continue
        nombre = (val(row, i_nom) + " " + val(row, i_ape)).strip().lstrip("?").strip()
        email = val(row, i_email)
        tel = val(row, i_tel)
        if tel and "e+" in tel.lower():
            tel = ""
        cargo = val(row, i_cargo)
        emp = val(row, i_emp) or val(row, i_vnom)
        if not emp:
            emp = nombre or email
        if not emp or es_junk(nombre + " " + email + " " + emp):
            continue
        key = _nz(emp)
        if key not in empresas:
            empresas[key] = {"empresa": emp, "nit": val(row, i_nit), "telefono": tel,
                             "email": email, "web": val(row, i_web),
                             "pais": val(row, i_pais), "vendedores": []}
            orden.append(key)
        e = empresas[key]
        if not e["email"] and email:
            e["email"] = email
        if not e["telefono"] and tel:
            e["telefono"] = tel
        if not e["nit"]:
            e["nit"] = val(row, i_nit)
        vnom = val(row, i_vnom) or (nombre if _nz(nombre) != key else "")
        if vnom and not any(_nz(v["nombre"]) == _nz(vnom) for v in e["vendedores"]):
            e["vendedores"].append({"nombre": vnom, "telefono": val(row, i_vtel) or tel,
                                    "email": val(row, i_vemail) or email, "cargo": cargo})
    return [empresas[k] for k in orden]


def add_months(fecha, meses):
    m = fecha.month - 1 + meses
    y = fecha.year + m // 12
    m = m % 12 + 1
    d = min(fecha.day, calendar.monthrange(y, m)[1])
    return datetime.date(y, m, d)


# ============================================================================
# Correo (Office 365 / SMTP)
# ============================================================================
def enviar_correo(cfg, destinatario, asunto, cuerpo, adjunto):
    remit = cfg.get("correo_remitente", "").strip()
    servidor = (cfg.get("smtp_servidor", "") or "smtp.office365.com").strip()
    try:
        puerto = int(cfg.get("smtp_puerto", "587") or 587)
    except Exception:
        puerto = 587
    pw = cfg.get("smtp_password", "")
    if not remit or not pw:
        raise ValueError("Falta configurar el correo remitente y su contrasena "
                         "en 'Datos de mi empresa'.")
    if not destinatario:
        raise ValueError("El cliente no tiene email.")
    msg = EmailMessage()
    msg["From"] = remit
    msg["To"] = destinatario
    msg["Subject"] = asunto
    msg.set_content(cuerpo)
    with open(adjunto, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="pdf",
                           filename=os.path.basename(adjunto))
    with smtplib.SMTP(servidor, puerto, timeout=30) as s:
        s.ehlo()
        s.starttls(context=ssl.create_default_context())
        s.ehlo()
        s.login(remit, pw)
        s.send_message(msg)


def _ics_recordatorio(item, fecha_dt):
    """Construye un evento de calendario (todo el dia) con recordatorio para el
       seguimiento de una cotizacion."""
    uid = (item.get("numero", "COT") + "-seg@innobadmc.com")
    ymd = fecha_dt.strftime("%Y%m%d")
    fin = (fecha_dt + datetime.timedelta(days=1)).strftime("%Y%m%d")
    stamp = datetime.datetime.now().strftime("%Y%m%dT%H%M%S")
    cli = (item.get("cliente", "") or "").replace(",", " ")
    ase = (item.get("asesor", "") or "")
    summ = f"Seguimiento cotizacion {item.get('numero','')} - {cli}"
    desc = (f"Dar seguimiento a la cotizacion {item.get('numero','')} de {cli}. "
            f"Asesor: {ase}. Destinos: {', '.join(item.get('destinos', []))}. "
            f"Total: {usd(item.get('total', 0))}.")
    return ("BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//INNOBA//Cotizador//ES\r\n"
            "CALSCALE:GREGORIAN\r\nMETHOD:PUBLISH\r\nBEGIN:VEVENT\r\n"
            f"UID:{uid}\r\nDTSTAMP:{stamp}\r\n"
            f"DTSTART;VALUE=DATE:{ymd}\r\nDTEND;VALUE=DATE:{fin}\r\n"
            f"SUMMARY:{summ}\r\nDESCRIPTION:{desc}\r\n"
            "BEGIN:VALARM\r\nTRIGGER:-PT9H\r\nACTION:DISPLAY\r\n"
            "DESCRIPTION:Recordatorio de seguimiento\r\nEND:VALARM\r\n"
            "END:VEVENT\r\nEND:VCALENDAR\r\n")

def enviar_recordatorio_ics(cfg, destinatarios, item, fecha_dt):
    """Envia por correo una invitacion de calendario (.ics) con el recordatorio."""
    remit = cfg.get("correo_remitente", "").strip()
    pw = cfg.get("smtp_password", "")
    if not remit or not pw:
        raise ValueError("Falta configurar el correo remitente y su contrasena.")
    dest = [d for d in destinatarios if d]
    if not dest:
        raise ValueError("No hay destinatario para el recordatorio.")
    servidor = (cfg.get("smtp_servidor", "") or "smtp.office365.com").strip()
    try:
        puerto = int(cfg.get("smtp_puerto", "587") or 587)
    except Exception:
        puerto = 587
    ics = _ics_recordatorio(item, fecha_dt)
    msg = EmailMessage()
    msg["From"] = remit
    msg["To"] = ", ".join(dest)
    msg["Subject"] = (f"Recordatorio seguimiento {item.get('numero','')} - "
                      f"{item.get('cliente','')} ({fecha_dt.strftime('%d/%m/%Y')})")
    msg.set_content(
        f"Recordatorio automatico: dar seguimiento a la cotizacion "
        f"{item.get('numero','')} de {item.get('cliente','')} el "
        f"{fecha_dt.strftime('%d/%m/%Y')}.\n\nAsesor: {item.get('asesor','')}\n"
        f"Adjuntamos una invitacion de calendario con recordatorio.")
    msg.add_attachment(ics.encode("utf-8"), maintype="text", subtype="calendar",
                       filename="seguimiento.ics", params={"method": "PUBLISH"})
    with smtplib.SMTP(servidor, puerto, timeout=30) as s:
        s.ehlo(); s.starttls(context=ssl.create_default_context()); s.ehlo()
        s.login(remit, pw); s.send_message(msg)


# ============================================================================
# Base de precios y logica
# ============================================================================
def cargar_precios():
    for ruta in (os.path.join(app_dir(), "precios_2026.json"), recurso("precios_2026.json")):
        if os.path.exists(ruta):
            with open(ruta, "r", encoding="utf-8") as f:
                return json.load(f)
    raise FileNotFoundError("No se encontro 'precios_2026.json'.")

def es_transporte(nombre):
    n = nombre.lower()
    return any(k in n for k in ("traslado", "asistencia", "asitencia", "transporte"))

def precio_terrestre_usd(servicio, pax, tasa, margen):
    precios = servicio["precios"]
    disp = sorted(int(k) for k in precios.keys())
    if not disp or not tasa:
        return 0.0
    col = pax if pax in disp else min(disp, key=lambda x: abs(x - pax))
    if pax > max(disp):
        col = max(disp)
    ppc = precios.get(str(col)) or precios.get(col)
    total_cop = ppc * pax
    venta_cop = total_cop / margen if margen else total_cop
    return venta_cop / tasa

def precio_hotel_usd_noche(valor_cop, tasa, margen):
    if not valor_cop or not tasa:
        return 0.0
    venta_cop = valor_cop / margen if margen else valor_cop
    return venta_cop / tasa


# ---- Reglas de ninos ----
# Edad en anos cumplidos. 0 = "0-11 meses" (menor de 1 ano).
EDAD_OPCIONES = ["0-11 meses", "1 ano", "2 anos", "3 anos", "4 anos", "5 anos",
                 "6 anos", "7 anos", "8 anos", "9 anos"]
CHILD_PRIVADO_USD = 10.0     # 12 meses a 2 anos, servicio PRIVADO: 10 USD
CHILD_HOTEL_COP = 70000.0    # 3 a 9 anos: 70.000 COP/noche

def es_privado(nombre):
    n = _norm_txt(nombre)
    return "privado" in n or "privada" in n   # si no dice -> compartido

def precio_servicio_grupo(serv, adultos, ninos_ages, tasa, margen, privado):
    """Total USD de un terrestre/tour para el grupo, aplicando reglas de ninos.
       0-11m: cortesia | 1-2 anos: privado 10USD / compartido 100% | 3-9: 50%."""
    N = adultos + sum(1 for a in ninos_ages if a >= 1)   # bebes 0-11m no ocupan cupo
    N = max(N, 1)
    total_N = precio_terrestre_usd(serv, N, tasa, margen)
    pp = total_N / N                                     # precio por persona (100%)
    total = adultos * pp
    for a in ninos_ages:
        if a == 0:
            continue                                     # cortesia
        elif a <= 2:
            total += CHILD_PRIVADO_USD if privado else pp
        else:
            total += 0.5 * pp
    return total

def precio_hotel_nino_noche(tasa, margen):
    """USD por noche por nino de 3-9 anos (70.000 / margen hotel / TRM)."""
    if not tasa:
        return 0.0
    return (CHILD_HOTEL_COP / margen) / tasa if margen else CHILD_HOTEL_COP / tasa


# ---- Descripciones de tours ----
_STOP = {"de", "del", "la", "el", "en", "y", "a", "por", "con", "los", "las",
         "tour", "tours", "visita", "dia", "full", "the", "compartido", "compartida",
         "privado", "privada", "especial", "sencilla", "sencillo", "doble", "grupo",
         "pax", "round", "trip", "in", "out", "ok", "o", "u", "para", "desde", "hacia"}

def _norm_txt(s):
    s = "".join(c for c in unicodedata.normalize("NFD", str(s))
                if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).lower().strip()

def _toks(s):
    return set(t for t in _norm_txt(s).split() if t not in _STOP and len(t) > 2)

def _match_score(a, b):
    ta, tb = _toks(a), _toks(b)
    if not ta or not tb:
        return 0.0, 0
    inter = len(ta & tb)
    if inter == 0:
        return 0.0, 0          # sin palabra clave en comun -> no forzar match
    cont = inter / len(ta)     # fraccion de palabras clave de A presentes en B
    jacc = inter / len(ta | tb)
    return max(cont, jacc), inter

def cargar_descripciones():
    for ruta in (os.path.join(app_dir(), "descripciones_tours.json"),
                 recurso("descripciones_tours.json")):
        if os.path.exists(ruta):
            try:
                with open(ruta, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
    return {}

def buscar_descripcion(nombre, destino, data, umbral=0.34):
    """Devuelve la MEJOR descripcion (misma ciudad) que comparta palabra clave.
       Prioriza no omitir descripciones correctas, sin poner una equivocada."""
    best = None; bs = 0.0
    for c in data.get(destino, []):
        sc, inter = _match_score(nombre, c["nombre"])
        if sc > bs:
            bs, best = sc, c
    if best and bs >= umbral:
        return best
    return None

def texto_descripcion(reg, maxlen=1200):
    d = (reg.get("descripcion", "") or "").strip()
    dur = (reg.get("duracion", "") or "").strip()
    inc = (reg.get("incluye", "") or "").strip()
    extra = []
    if dur and len(dur) < 40:
        extra.append("Duracion: " + dur)
    if inc:
        extra.append("Incluye: " + inc)
    txt = d
    if extra:
        txt = (d + "  " if d else "") + " | ".join(extra)
    txt = re.sub(r"\s+", " ", txt).strip()
    if len(txt) > maxlen:
        txt = txt[:maxlen - 1].rsplit(" ", 1)[0] + "..."
    return txt


# ============================================================================
# PDF
# ============================================================================
PDF_PRIM = (1, 57, 132); PDF_BLUE = (20, 102, 199)
PDF_CLARO = (233, 240, 250); PDF_TXT = (30, 40, 60)

def usd(v):
    try:
        return f"USD {v:,.2f}"
    except Exception:
        return f"USD {v}"


class CotizacionPDF(FPDF):
    def __init__(self, cfg):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.cfg = cfg
        self.set_auto_page_break(auto=True, margin=18)
        self.set_margins(15, 15, 15)

    def header(self):
        cfg = self.cfg; y0 = 12; logo = cfg.get("logo", ""); text_x = 15
        if logo and os.path.exists(logo):
            try:
                with Image.open(logo) as im:
                    w_px, h_px = im.size
                max_w, max_h = 42, 24
                ratio = min(max_w / w_px, max_h / h_px)
                w_mm = w_px * ratio * 0.2645833; h_mm = h_px * ratio * 0.2645833
                if h_mm > max_h:
                    s = max_h / h_mm; w_mm *= s; h_mm *= s
                self.image(logo, x=15, y=y0, h=h_mm)
                text_x = 15 + w_mm + 6
            except Exception:
                text_x = 15
        self.set_xy(text_x, y0)
        self.set_text_color(*PDF_PRIM); self.set_font("Helvetica", "B", 16)
        self.cell(0, 7, self._t(cfg.get("empresa", "")), ln=1)
        self.set_x(text_x); self.set_text_color(*PDF_TXT); self.set_font("Helvetica", "", 9)
        l2 = []
        if cfg.get("nit"): l2.append("NIT/RUC: " + cfg["nit"])
        if cfg.get("direccion"): l2.append(cfg["direccion"])
        if l2:
            self.set_x(text_x); self.cell(0, 5, self._t("  |  ".join(l2)), ln=1)
        l3 = []
        if cfg.get("telefono"): l3.append("Tel: " + cfg["telefono"])
        if cfg.get("email"): l3.append(cfg["email"])
        if cfg.get("web"): l3.append(cfg["web"])
        if l3:
            self.set_x(text_x); self.cell(0, 5, self._t("  |  ".join(l3)), ln=1)
        self.set_draw_color(*PDF_PRIM); self.set_line_width(0.6)
        self.line(15, 40, 195, 40); self.set_y(45)

    def footer(self):
        self.set_y(-14); self.set_font("Helvetica", "I", 8)
        self.set_text_color(130, 130, 130)
        self.cell(0, 10, self._t(f"{self.cfg.get('empresa','')}  -  Pagina {self.page_no()}"),
                  align="C")

    def _t(self, texto):
        if texto is None:
            return ""
        return str(texto).encode("latin-1", "replace").decode("latin-1")


def _seccion_tabla(pdf, titulo, filas, total_seccion):
    if not filas:
        return
    T = pdf._t
    pdf.ln(1)
    pdf.set_font("Helvetica", "B", 10); pdf.set_text_color(*PDF_BLUE)
    pdf.cell(0, 6, T(titulo), ln=1)
    w_desc, w_det, w_pp, w_val = 90, 30, 30, 30
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_fill_color(*PDF_PRIM); pdf.set_text_color(255, 255, 255)
    pdf.cell(w_desc, 7, T("  Concepto"), fill=True)
    pdf.cell(w_det, 7, T("Detalle"), fill=True, align="C")
    pdf.cell(w_pp, 7, T("Por pasajero"), fill=True, align="C")
    pdf.cell(w_val, 7, T("Total (USD)"), fill=True, align="C", ln=1)
    pdf.set_text_color(*PDF_TXT)
    f = 0
    for fila in filas:
        desc, det, val = fila[0], fila[1], fila[2]
        pp = fila[3] if len(fila) > 3 else None
        descripcion = fila[4] if len(fila) > 4 else ""
        relleno = (f % 2 == 1)
        pdf.set_font("Helvetica", "B", 9)
        lin_c = pdf.multi_cell(w_desc - 4, 4.6, T("  " + desc), border=0,
                               align="L", split_only=True)
        alto_c = len(lin_c) * 4.6
        lin_d = []
        if descripcion:
            pdf.set_font("Helvetica", "I", 7.5)
            lin_d = pdf.multi_cell(w_desc - 6, 3.6, T("   " + descripcion), border=0,
                                   align="L", split_only=True)
        alto = max(7, alto_c + len(lin_d) * 3.6 + 3)
        x0 = pdf.get_x(); y0 = pdf.get_y()
        if y0 + alto > (297 - 18):
            pdf.add_page(); x0 = pdf.get_x(); y0 = pdf.get_y()
        if relleno:
            pdf.set_fill_color(*PDF_CLARO)
        pdf.multi_cell(w_desc, alto, "", border=0, fill=relleno)
        pdf.set_xy(x0, y0 + 1)
        pdf.set_text_color(*PDF_TXT); pdf.set_font("Helvetica", "B", 9)
        pdf.multi_cell(w_desc - 4, 4.6, T("  " + desc), border=0, align="L")
        if descripcion:
            pdf.set_xy(x0, y0 + 1 + alto_c)
            pdf.set_text_color(110, 120, 135); pdf.set_font("Helvetica", "I", 7.5)
            pdf.multi_cell(w_desc - 6, 3.6, T("   " + descripcion), border=0, align="L")
            pdf.set_text_color(*PDF_TXT)
        pdf.set_xy(x0 + w_desc, y0); pdf.set_font("Helvetica", "", 9)
        pdf.cell(w_det, alto, T(det), align="C", fill=relleno)
        pdf.cell(w_pp, alto, T(usd(pp) if pp else "-"), align="R", fill=relleno)
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(w_val, alto, T(usd(val)), align="R", fill=relleno, ln=1)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_y(y0 + alto)
        f += 1


def _banda_destino(pdf, texto):
    T = pdf._t
    pdf.ln(3)
    if pdf.get_y() > 250:
        pdf.add_page()
    pdf.set_fill_color(*PDF_BLUE); pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, T("  " + texto), ln=1, fill=True)
    pdf.set_text_color(*PDF_TXT)


def _tabla_hoteles_combinada(pdf, con_op):
    """Multidestino: una sola tabla que empareja el hotel de cada destino (por
       orden) y SUMA los precios por persona (Sencilla/Doble/Triple)."""
    T = pdf._t
    dests = [b["destino"] for b in con_op]
    pdf.ln(2)
    if pdf.get_y() > 230:
        pdf.add_page()
    pdf.set_font("Helvetica", "B", 10); pdf.set_text_color(*PDF_BLUE)
    pdf.cell(0, 6, T("OPCIONES DE HOTEL - precio por persona SUMANDO los "
                     f"{len(dests)} destinos (el cliente elige)"), ln=1)
    n = len(con_op)
    w_price, w_cat = 22.0, 24.0
    w_h = max(24.0, (180.0 - w_cat - 3 * w_price) / n)
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_fill_color(*PDF_PRIM); pdf.set_text_color(255, 255, 255)
    for b in con_op:
        pdf.cell(w_h, 7, T("  Hotel " + b["destino"][:13]), fill=True)
    pdf.cell(w_cat, 7, T("Categoria"), fill=True, align="C")
    pdf.cell(w_price, 7, T("Sencilla"), fill=True, align="C")
    pdf.cell(w_price, 7, T("Doble"), fill=True, align="C")
    pdf.cell(w_price, 7, T("Triple"), fill=True, align="C", ln=1)
    pdf.set_text_color(*PDF_TXT)

    def money(v):
        return f"{v:,.2f}" if v else "-"

    maxn = max(len(b["opciones"]) for b in con_op)
    for i in range(maxn):
        fila = [b["opciones"][min(i, len(b["opciones"]) - 1)] for b in con_op]
        relleno = (i % 2 == 1)
        if relleno:
            pdf.set_fill_color(*PDF_CLARO)
        if pdf.get_y() > 275:
            pdf.add_page()
        pdf.set_font("Helvetica", "B", 8)
        for o in fila:
            pdf.cell(w_h, 8, T("  " + o["nombre"][:24]), fill=relleno)
        cats = []
        for o in fila:
            c = (o.get("categoria") or "").strip()
            if c and c not in cats:
                cats.append(c)
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(w_cat, 8, T(("/".join(cats))[:14] or "-"), align="C", fill=relleno)

        def suma(acc):
            vals = [o.get(acc) for o in fila]
            return sum(vals) if all(v for v in vals) else None
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(w_price, 8, T(money(suma("sencilla"))), align="R", fill=relleno)
        pdf.cell(w_price, 8, T(money(suma("doble"))), align="R", fill=relleno)
        pdf.cell(w_price, 8, T(money(suma("triple"))), align="R", fill=relleno, ln=1)
    pdf.set_font("Helvetica", "I", 8); pdf.set_text_color(110, 120, 135)
    pdf.cell(0, 5, T("  Valores en USD POR PERSONA (adulto), SUMANDO "
                     + " + ".join(dests) + ", por todo el viaje. Incluye traslados "
                     "y actividades."), ln=1)
    pdf.set_text_color(*PDF_TXT)


def generar_pdf(cfg, datos, bloques, total, ruta_salida):
    """bloques: lista de dicts {destino, subtitulo, secciones:[(t,filas,sub)], subtotal}."""
    pdf = CotizacionPDF(cfg)
    pdf.add_page()
    T = pdf._t
    multi = len(bloques) > 1
    titulo = "  COTIZACION" + (" - ITINERARIO" if multi
                               else (" - " + bloques[0]["destino"] if bloques else ""))
    pdf.set_fill_color(*PDF_PRIM); pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 13)
    pdf.cell(0, 9, T(titulo), ln=1, fill=True)
    pdf.ln(2)

    pdf.set_text_color(*PDF_TXT)
    y_b = pdf.get_y()
    pdf.set_font("Helvetica", "B", 10); pdf.set_text_color(*PDF_PRIM)
    pdf.cell(90, 6, T("CLIENTE"), ln=1)
    pdf.set_text_color(*PDF_TXT)
    for etq, clave in [("Nombre", "cliente"), ("Email", "cli_email"),
                       ("Telefono", "cli_tel"), ("Asesor", "asesor"),
                       ("Tel. asesor", "asesor_tel")]:
        val = datos.get(clave, "")
        if val:
            pdf.set_font("Helvetica", "B", 9); pdf.cell(22, 5, T(etq + ":"))
            pdf.set_font("Helvetica", "", 9); pdf.cell(68, 5, T(val), ln=1)
    y_izq = pdf.get_y()
    pdf.set_xy(110, y_b)
    pdf.set_font("Helvetica", "B", 10); pdf.set_text_color(*PDF_PRIM)
    pdf.cell(85, 6, T("DETALLES"), ln=1)
    pdf.set_text_color(*PDF_TXT)
    for etq, val in [("No. Cotizacion", datos.get("numero", "")),
                     ("Fecha", datos.get("fecha", "")),
                     ("Valida hasta", datos.get("valida_hasta", "")),
                     ("Fechas viaje", datos.get("fechas_viaje", "")),
                     ("Pasajeros", datos.get("pax_txt", ""))]:
        if val:
            pdf.set_x(110)
            pdf.set_font("Helvetica", "B", 9); pdf.cell(30, 5, T(etq + ":"))
            pdf.set_font("Helvetica", "", 9); pdf.cell(55, 5, T(str(val)), ln=1)
    pdf.set_y(max(y_izq, pdf.get_y()) + 2)

    def edad_txt(a):
        return "Bebe 0-11 meses" if a == 0 else f"{a} " + ("ano" if a == 1 else "anos")

    def cel(v):
        return usd(v) if v else "-"

    con_op = [b for b in bloques if b["opciones"]]
    combinar = len(con_op) > 1   # multidestino: tabla de hoteles combinada al final
    for b in bloques:
        _banda_destino(pdf, b["subtitulo"])
        for titulo_s, filas, sub in b["base_secciones"]:
            _seccion_tabla(pdf, titulo_s, filas, sub)
        ops = b["opciones"]
        # Opciones de hotel POR DESTINO (solo si NO es multidestino)
        if ops and not combinar:
            pdf.ln(2)
            if pdf.get_y() > 240:
                pdf.add_page()
            pdf.set_font("Helvetica", "B", 10); pdf.set_text_color(*PDF_BLUE)
            titulo_op = "OPCIONES DE HOTEL - precio por persona (el cliente elige)" \
                if len(ops) > 1 else "ALOJAMIENTO - precio por persona"
            pdf.cell(0, 6, T(titulo_op), ln=1)
            w_h, w_c, w_p = 74, 30, 25
            pdf.set_font("Helvetica", "B", 9)
            pdf.set_fill_color(*PDF_PRIM); pdf.set_text_color(255, 255, 255)
            pdf.cell(w_h, 7, T("  Hotel"), fill=True)
            pdf.cell(w_c, 7, T("Categoria"), fill=True, align="C")
            pdf.cell(w_p, 7, T("Sencilla"), fill=True, align="C")
            pdf.cell(w_p, 7, T("Doble"), fill=True, align="C")
            pdf.cell(w_p, 7, T("Triple"), fill=True, align="C", ln=1)
            pdf.set_text_color(*PDF_TXT)
            for j, op in enumerate(ops):
                relleno = (j % 2 == 1)
                if relleno:
                    pdf.set_fill_color(*PDF_CLARO)
                pdf.set_font("Helvetica", "B", 9)
                pdf.cell(w_h, 8, T("  " + op["nombre"][:38]), fill=relleno)
                pdf.set_font("Helvetica", "", 9)
                pdf.cell(w_c, 8, T(op["categoria"] or "-"), align="C", fill=relleno)
                pdf.set_font("Helvetica", "B", 9)
                pdf.cell(w_p, 8, T(cel(op["sencilla"])), align="R", fill=relleno)
                pdf.cell(w_p, 8, T(cel(op["doble"])), align="R", fill=relleno)
                pdf.cell(w_p, 8, T(cel(op["triple"])), align="R", fill=relleno, ln=1)
            pdf.set_font("Helvetica", "I", 8); pdf.set_text_color(110, 120, 135)
            pdf.cell(0, 5, T("  Valores POR PERSONA (adulto) segun acomodacion, "
                             "por todo el viaje. Incluye traslados y actividades."), ln=1)
            pdf.set_text_color(*PDF_TXT)
        # Precio por nino (fijo, no depende del hotel ni de la acomodacion)
        if b["ninos"]:
            pdf.set_font("Helvetica", "", 8); pdf.set_text_color(*PDF_TXT)
            partes = [f"{edad_txt(a)} (x{c}): {usd(pr)}" for a, c, pr in b["ninos"]]
            pdf.multi_cell(0, 4.5, T("  Precio por nino: " + "   |   ".join(partes)))

    # ---- Multidestino: tabla de hoteles combinada (suma de destinos) ----
    if combinar:
        _tabla_hoteles_combinada(pdf, con_op)

    # ---- Costo total de la reserva: 1a opcion de hotel + habitaciones indicadas ----
    hab = datos.get("habitaciones") or {}
    OCCP = {"sencilla": 1, "doble": 2, "triple": 3}
    ocup = sum(hab.get(k, 0) * OCCP[k] for k in OCCP)
    if con_op and ocup > 0:
        n_ni = sum(c for b in con_op for a, c, pr in b["ninos"])
        total_res = 0.0; detalle_hab = []
        for acc in ("sencilla", "doble", "triple"):
            n = hab.get(acc, 0)
            if not n:
                continue
            detalle_hab.append(f"{n} {acc}")
            for b in con_op:
                pp = b["opciones"][0].get(acc)
                if pp:
                    total_res += n * OCCP[acc] * pp
        total_res += sum(c * pr for b in con_op for a, c, pr in b["ninos"])
        hoteles_op1 = " + ".join(b["opciones"][0]["nombre"] for b in con_op)
        pdf.ln(4)
        if pdf.get_y() > 245:
            pdf.add_page()
        pax_txt = (f"{con_op[0]['n_adultos']} adulto(s)"
                   + (f" + {n_ni} nino(s)" if n_ni else ""))
        pdf.set_font("Helvetica", "B", 10); pdf.set_fill_color(*PDF_PRIM)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(0, 8, T(f"  COSTO TOTAL DE LA RESERVA - 1a opcion ({pax_txt})"),
                 ln=1, fill=True)
        pdf.set_text_color(*PDF_TXT); pdf.set_font("Helvetica", "", 9)
        pdf.set_fill_color(*PDF_CLARO)
        pdf.cell(0, 7, T("  Habitaciones solicitadas: " + ", ".join(detalle_hab)),
                 fill=True, ln=1)
        pdf.set_font("Helvetica", "B", 12); pdf.set_fill_color(*PDF_PRIM)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(120, 9, T("  TOTAL DE LA RESERVA (USD)"), fill=True)
        pdf.cell(0, 9, T(usd(total_res)), align="R", fill=True, ln=1)
        pdf.set_text_color(110, 120, 135); pdf.set_font("Helvetica", "I", 8)
        pdf.cell(0, 5, T(f"  Calculado con la 1a opcion: {hoteles_op1[:110]}"), ln=1)
        pdf.set_text_color(*PDF_TXT)

    # ---- Itinerario dia por dia (opcional) ----
    itin = (datos.get("itinerario") or "").strip()
    if itin:
        pdf.add_page()
        pdf.set_fill_color(*PDF_PRIM); pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 13)
        pdf.cell(0, 9, T("  ITINERARIO DIA POR DIA"), ln=1, fill=True); pdf.ln(2)
        pdf.set_text_color(*PDF_TXT)
        for par in itin.split("\n"):
            par = par.strip()
            if not par:
                pdf.ln(1); continue
            if pdf.get_y() > 270:
                pdf.add_page()
            if re.match(r"(?i)^d[ií]a\s*\d", par):
                pdf.ln(1); pdf.set_fill_color(*PDF_BLUE); pdf.set_text_color(255, 255, 255)
                pdf.set_font("Helvetica", "B", 10)
                pdf.cell(0, 7, T("  " + par), ln=1, fill=True)
                pdf.set_text_color(*PDF_TXT)
            else:
                pdf.set_font("Helvetica", "", 9)
                pdf.multi_cell(0, 4.8, T(par))

    pdf.ln(6); pdf.set_text_color(*PDF_PRIM); pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 6, T("Notas y condiciones"), ln=1)
    pdf.set_text_color(*PDF_TXT); pdf.set_font("Helvetica", "", 9)
    notas = datos.get("notas", "") or cfg.get("notas", "")
    if notas:
        pdf.multi_cell(0, 5, T(notas))

    firma_nom = (datos.get("firma_nombre") or cfg.get("firma_nombre", "")).strip()
    firma_cargo = (datos.get("firma_cargo") or cfg.get("firma_cargo", "")).strip()
    if firma_nom:
        if pdf.get_y() > 240:
            pdf.add_page()
        pdf.ln(16); pdf.set_text_color(*PDF_TXT); pdf.set_font("Helvetica", "", 9)
        pdf.cell(0, 5, T("Cordialmente,"), ln=1)
        pdf.ln(10)
        pdf.set_draw_color(*PDF_PRIM); pdf.set_line_width(0.4)
        y = pdf.get_y(); pdf.line(15, y, 85, y); pdf.ln(1)
        pdf.set_font("Helvetica", "B", 10); pdf.set_text_color(*PDF_PRIM)
        pdf.cell(0, 5, T(firma_nom), ln=1)
        if firma_cargo:
            pdf.set_font("Helvetica", "", 9); pdf.set_text_color(*PDF_TXT)
            pdf.cell(0, 5, T(firma_cargo), ln=1)

    pdf.output(ruta_salida)


# ============================================================================
# Widgets auxiliares
# ============================================================================
class Stepper(ctk.CTkFrame):
    def __init__(self, master, value=0, minimo=0, maximo=99, width=104,
                 command=None, **kw):
        super().__init__(master, fg_color=CARD2, corner_radius=8, **kw)
        self.minimo, self.maximo = minimo, maximo
        self.command = command
        self.var = tk.StringVar(value=str(value))
        ctk.CTkButton(self, text="-", width=28, height=28, corner_radius=8,
                      fg_color=NAVY, hover_color=BLUE, font=("Segoe UI", 15, "bold"),
                      command=self._menos).pack(side="left", padx=3, pady=3)
        ctk.CTkEntry(self, textvariable=self.var, width=width - 74, height=28,
                     justify="center", border_width=0, fg_color=CARD2, text_color=TEXT,
                     font=("Segoe UI", 13, "bold")).pack(side="left")
        ctk.CTkButton(self, text="+", width=28, height=28, corner_radius=8,
                      fg_color=NAVY, hover_color=BLUE, font=("Segoe UI", 15, "bold"),
                      command=self._mas).pack(side="left", padx=3, pady=3)
        self.var.trace_add("write", lambda *a: self.command() if self.command else None)

    def get(self):
        try:
            return int(float(self.var.get()))
        except Exception:
            return self.minimo

    def set(self, v):
        self.var.set(str(v))

    def _menos(self):
        self.set(max(self.minimo, self.get() - 1))

    def _mas(self):
        self.set(min(self.maximo, self.get() + 1))


# ============================================================================
# Ventana de busqueda rapida de cliente (lupa)
# ============================================================================
class BuscadorClientes(ctk.CTkToplevel):
    def __init__(self, master, clientes, on_pick, on_editar=None, on_eliminar=None):
        super().__init__(master)
        self.clientes = clientes; self.on_pick = on_pick
        self.on_editar = on_editar; self.on_eliminar = on_eliminar
        self.title("Buscar cliente")
        self.geometry("480x540"); self.configure(fg_color=BG)
        self.transient(master); self.grab_set()
        self.grid_columnconfigure(0, weight=1); self.grid_rowconfigure(2, weight=1)
        ctk.CTkLabel(self, text="Buscar cliente", text_color=NAVY,
                     font=("Segoe UI", 16, "bold")).grid(row=0, column=0, sticky="w",
                                                         padx=16, pady=(14, 2))
        self.var_q = tk.StringVar()
        e = ctk.CTkEntry(self, textvariable=self.var_q, height=40, corner_radius=10,
                         border_color=BLUE, border_width=2, fg_color=CARD,
                         font=("Segoe UI", 13),
                         placeholder_text="Escribe la empresa o el asesor...")
        e.grid(row=1, column=0, sticky="ew", padx=16, pady=(2, 8))
        self.var_q.trace_add("write", lambda *a: self._pintar())
        self.lista = ctk.CTkScrollableFrame(self, fg_color=CARD, corner_radius=12)
        self.lista.grid(row=2, column=0, sticky="nsew", padx=16, pady=(0, 14))
        self._pintar()
        self.after(120, e.focus_set)
        self.bind("<Escape>", lambda ev: self.destroy())

    def _pintar(self):
        for w in self.lista.winfo_children():
            w.destroy()
        q = _nz(self.var_q.get())
        n = 0
        for c in self.clientes:
            emp = c.get("empresa", "")
            vends = c.get("vendedores", []) or []
            if q and not (q in _nz(emp) or any(q in _nz(v.get("nombre", "")) for v in vends)):
                continue
            n += 1
            if n > 250:
                break
            sub = ", ".join(v.get("nombre", "") for v in vends[:3])
            txt = emp + (("\n   " + sub) if sub else "")
            fila = ctk.CTkFrame(self.lista, fg_color=CARD2, corner_radius=8)
            fila.pack(fill="x", padx=4, pady=2)
            fila.grid_columnconfigure(0, weight=1)
            ctk.CTkButton(fila, text=txt, anchor="w", height=42, corner_radius=8,
                          fg_color=CARD2, text_color=NAVY, hover_color=LINE,
                          font=("Segoe UI", 12, "bold"),
                          command=lambda x=emp: self._pick(x)).grid(row=0, column=0, sticky="ew")
            if self.on_editar:
                ctk.CTkButton(fila, text="✏", width=34, height=42, corner_radius=8,
                              fg_color=CARD2, text_color=BLUE, hover_color=LINE,
                              font=("Segoe UI", 14),
                              command=lambda x=emp: self._editar(x)).grid(row=0, column=1)
            if self.on_eliminar:
                ctk.CTkButton(fila, text="🗑", width=34, height=42, corner_radius=8,
                              fg_color=CARD2, text_color=RED, hover_color=LINE,
                              font=("Segoe UI", 14),
                              command=lambda x=emp: self._eliminar(x)).grid(row=0, column=2)
        if not n:
            ctk.CTkLabel(self.lista, text="Sin resultados.", text_color=MUTED).pack(pady=24)

    def _pick(self, emp):
        self.on_pick(emp); self.destroy()

    def _editar(self, emp):
        self.destroy()
        if self.on_editar:
            self.on_editar(emp)

    def _eliminar(self, emp):
        if messagebox.askyesno("Eliminar cliente",
                               f"¿Eliminar definitivamente a:\n\n{emp}\n\nEsta accion no se puede deshacer.",
                               parent=self):
            if self.on_eliminar:
                self.on_eliminar(emp)
            self._pintar()


# ============================================================================
# Ventana de Historial de Cotizaciones (consecutivo + busqueda + seguimiento)
# ============================================================================
ESTADOS_COT = ["Pendiente", "Enviada", "En seguimiento", "Ganada", "Perdida"]
ESTADO_COLOR = {"Pendiente": MUTED, "Enviada": BLUE, "En seguimiento": "#D9A400",
                "Ganada": GREEN, "Perdida": RED}
# color de fondo de la fila segun estado (Ganada verde, Perdida rojo, En segui. AMARILLO)
ESTADO_FILA = {"Pendiente": "#F1F5FB", "Enviada": "#EAF2FD", "En seguimiento": "#FFF3C4",
               "Ganada": "#E3F5EA", "Perdida": "#FBE6E6"}


class VentanaCotizacionDetalle(ctk.CTkToplevel):
    """Editar la cotizacion y gestionar tareas de seguimiento."""
    def __init__(self, master, item, on_save):
        super().__init__(master)
        self.item = item; self.on_save = on_save
        self.title("Cotizacion " + item.get("numero", ""))
        self.geometry("660x700"); self.configure(fg_color=BG)
        self.transient(master); self.grab_set()
        cont = ctk.CTkScrollableFrame(self, fg_color=BG)
        cont.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(cont, text=f"{item.get('numero','')}   ·   {item.get('fecha','')}",
                     text_color=NAVY, font=("Segoe UI", 16, "bold")).pack(anchor="w")
        dest = ", ".join(item.get("destinos", []))
        ctk.CTkLabel(cont, text=f"{dest}   ·   {usd(item.get('total', 0))}",
                     text_color=MUTED, font=("Segoe UI", 11)).pack(anchor="w", pady=(0, 8))

        def campo(lbl, val):
            ctk.CTkLabel(cont, text=lbl, text_color=MUTED,
                         font=("Segoe UI", 11)).pack(anchor="w", padx=2)
            v = tk.StringVar(value=val)
            ctk.CTkEntry(cont, textvariable=v, height=32, corner_radius=8,
                         border_color=LINE).pack(fill="x", pady=(0, 8))
            return v
        self.v_cli = campo("Agencia / cliente", item.get("cliente", ""))
        self.v_ase = campo("Asesor", item.get("asesor", ""))
        ctk.CTkLabel(cont, text="Estado del seguimiento", text_color=MUTED,
                     font=("Segoe UI", 11)).pack(anchor="w", padx=2)
        self.v_estado = tk.StringVar(value=item.get("estado", "Pendiente"))
        ctk.CTkOptionMenu(cont, variable=self.v_estado, values=ESTADOS_COT, width=220,
                          height=32, corner_radius=8, fg_color=NAVY, button_color=NAVY2,
                          button_hover_color=BLUE, dropdown_fg_color=CARD,
                          dropdown_text_color=TEXT).pack(anchor="w", pady=(0, 8))
        self.v_fecha_seg = campo("Fecha de seguimiento / recordatorio (dd/mm/aaaa)",
                                 item.get("fecha_seg", ""))
        ctk.CTkLabel(cont, text="El sistema te avisara al abrir cuando llegue esa fecha.",
                     text_color=MUTED, font=("Segoe UI", 9)).pack(anchor="w", padx=2, pady=(0, 6))
        ctk.CTkLabel(cont, text="Notas", text_color=MUTED,
                     font=("Segoe UI", 11)).pack(anchor="w", padx=2)
        self.txt_notas = ctk.CTkTextbox(cont, height=70, corner_radius=8, border_width=1,
                                        border_color=LINE, fg_color=CARD)
        self.txt_notas.insert("1.0", item.get("notas", "")); self.txt_notas.pack(fill="x", pady=(0, 10))
        # --- tareas de seguimiento ---
        hdr = ctk.CTkFrame(cont, fg_color="transparent"); hdr.pack(fill="x", pady=(4, 2))
        ctk.CTkLabel(hdr, text="Tareas de seguimiento", text_color=NAVY,
                     font=("Segoe UI", 14, "bold")).pack(side="left")
        ctk.CTkButton(hdr, text="+ Agregar tarea", width=140, height=28, corner_radius=8,
                      fg_color=CARD2, text_color=NAVY, hover_color=LINE, border_width=1,
                      border_color=LINE, font=("Segoe UI", 11, "bold"),
                      command=lambda: self._add_tarea()).pack(side="right")
        self.tfr = ctk.CTkFrame(cont, fg_color="transparent"); self.tfr.pack(fill="x", pady=4)
        self.tareas_rows = []
        for t in item.get("tareas", []):
            self._add_tarea(t)
        if not item.get("tareas"):
            self._add_tarea()
        bts = ctk.CTkFrame(cont, fg_color="transparent"); bts.pack(fill="x", pady=12)
        ctk.CTkButton(bts, text="Guardar", fg_color=GREEN, hover_color=GREEN_H,
                      font=("Segoe UI", 13, "bold"), command=self._guardar).pack(side="right")
        ctk.CTkButton(bts, text="Cancelar", fg_color=CARD2, text_color=NAVY, hover_color=LINE,
                      command=self.destroy).pack(side="right", padx=8)

    def _add_tarea(self, t=None):
        t = t or {}
        row = ctk.CTkFrame(self.tfr, fg_color=CARD2, corner_radius=8)
        row.pack(fill="x", pady=3)
        hecha = tk.BooleanVar(value=bool(t.get("hecha")))
        ctk.CTkCheckBox(row, text="", variable=hecha, width=24, checkbox_width=20,
                        checkbox_height=20, corner_radius=5, fg_color=GREEN,
                        hover_color=GREEN_H).pack(side="left", padx=(8, 2), pady=6)
        v_txt = tk.StringVar(value=t.get("texto", ""))
        ctk.CTkEntry(row, textvariable=v_txt, height=30, corner_radius=6, border_color=LINE,
                     placeholder_text="Que hacer (ej. llamar al cliente)").pack(
            side="left", fill="x", expand=True, padx=3, pady=6)
        v_fec = tk.StringVar(value=t.get("fecha", ""))
        ctk.CTkEntry(row, textvariable=v_fec, width=110, height=30, corner_radius=6,
                     border_color=LINE, placeholder_text="dd/mm/aaaa").pack(side="left", padx=3)
        entry = {"frame": row, "hecha": hecha, "texto": v_txt, "fecha": v_fec}
        ctk.CTkButton(row, text="✕", width=28, height=28, corner_radius=6, fg_color="transparent",
                      text_color=RED, hover_color=LINE,
                      command=lambda: (self.tareas_rows.remove(entry), row.destroy())).pack(
            side="left", padx=(2, 6))
        self.tareas_rows.append(entry)

    def _guardar(self):
        self.item["cliente"] = self.v_cli.get().strip()
        self.item["asesor"] = self.v_ase.get().strip()
        self.item["estado"] = self.v_estado.get()
        self.item["fecha_seg"] = self.v_fecha_seg.get().strip()
        self.item["notas"] = self.txt_notas.get("1.0", "end").strip()
        self.item["tareas"] = [
            {"texto": r["texto"].get().strip(), "fecha": r["fecha"].get().strip(),
             "hecha": r["hecha"].get()}
            for r in self.tareas_rows if r["texto"].get().strip()]
        self.on_save()
        # ofrecer enviar recordatorio de calendario (.ics) para la fecha de seguimiento
        fseg = parse_fecha(self.item.get("fecha_seg", ""))
        if fseg and self.item.get("estado") not in ("Ganada", "Perdida"):
            app = getattr(self.master, "master", None)
            cfg = getattr(app, "cfg", None) if app else None
            if cfg and cfg.get("correo_remitente") and cfg.get("smtp_password"):
                if messagebox.askyesno(
                        "Recordatorio de seguimiento",
                        f"¿Enviar un recordatorio de calendario para el "
                        f"{self.item.get('fecha_seg')} a tu correo?", parent=self):
                    try:
                        dest = [cfg.get("correo_remitente")]
                        enviar_recordatorio_ics(cfg, dest, self.item, fseg)
                        messagebox.showinfo("Recordatorio",
                                            "Recordatorio de calendario enviado.", parent=self)
                    except Exception as e:
                        messagebox.showwarning("Recordatorio",
                                               "No se pudo enviar el recordatorio:\n" + str(e),
                                               parent=self)
        self.destroy()


class VentanaCotizaciones(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Historial de cotizaciones")
        self.geometry("920x620"); self.configure(fg_color=BG)
        self.transient(master); self.grab_set()
        # traer cotizaciones nuevas hechas por clientes en el HTML
        try:
            importar_cotizaciones_html()
        except Exception:
            pass
        self.data = cargar_cotizaciones()
        self.grid_columnconfigure(0, weight=1); self.grid_rowconfigure(2, weight=1)
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 2))
        top.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(top, text="Historial de cotizaciones", text_color=NAVY,
                     font=("Segoe UI", 16, "bold")).grid(row=0, column=0, sticky="w")
        if WEBHOOK_URL:
            ctk.CTkButton(top, text="↻ Importar del HTML", width=160, height=32, corner_radius=8,
                          fg_color=NAVY, hover_color=BLUE, font=("Segoe UI", 11, "bold"),
                          command=self._importar_html).grid(row=0, column=1, sticky="e")
        self.var_q = tk.StringVar()
        e = ctk.CTkEntry(self, textvariable=self.var_q, height=36, corner_radius=10,
                         border_color=BLUE, border_width=2, fg_color=CARD, font=("Segoe UI", 12),
                         placeholder_text="Buscar por consecutivo, agencia o asesor...")
        e.grid(row=1, column=0, sticky="ew", padx=16, pady=(2, 8))
        self.var_q.trace_add("write", lambda *a: self._pintar())
        self.lista = ctk.CTkScrollableFrame(self, fg_color=CARD, corner_radius=12)
        self.lista.grid(row=2, column=0, sticky="nsew", padx=16, pady=(0, 14))
        self.after(120, e.focus_set)
        self.bind("<Escape>", lambda ev: self.destroy())
        self._pintar()

    def _pintar(self):
        for w in self.lista.winfo_children():
            w.destroy()
        q = _nz(self.var_q.get())
        items = list(reversed(self.data.get("items", [])))   # mas recientes primero
        n = 0
        for it in items:
            campos = " ".join([it.get("numero", ""), it.get("cliente", ""),
                               it.get("asesor", "")])
            if q and q not in _nz(campos):
                continue
            n += 1
            estado = it.get("estado", "Pendiente")
            fila = ctk.CTkFrame(self.lista, fg_color=ESTADO_FILA.get(estado, CARD2),
                                corner_radius=8, border_width=1, border_color=LINE)
            fila.pack(fill="x", padx=4, pady=2)
            fila.grid_columnconfigure(0, weight=1)
            dest = ", ".join(it.get("destinos", []))
            tareas = it.get("tareas", []) or []
            pend = sum(1 for t in tareas if not t.get("hecha"))
            fseg = parse_fecha(it.get("fecha_seg", ""))
            venc = fseg and fseg <= datetime.date.today() and estado not in ("Ganada", "Perdida")
            cel = ctk.CTkFrame(fila, fg_color="transparent")
            cel.grid(row=0, column=0, sticky="w", padx=10, pady=6)
            top = ctk.CTkFrame(cel, fg_color="transparent"); top.pack(anchor="w")
            ctk.CTkLabel(top, text=f"{it.get('numero','')}    {it.get('cliente') or '(sin agencia)'}",
                         text_color=NAVY, anchor="w",
                         font=("Segoe UI", 12, "bold")).pack(side="left")
            ctk.CTkLabel(top, text=" " + estado + " ", text_color="#FFFFFF",
                         fg_color=ESTADO_COLOR.get(estado, MUTED), corner_radius=6,
                         font=("Segoe UI", 9, "bold")).pack(side="left", padx=8)
            if it.get("fecha_seg"):
                ctk.CTkLabel(top, text=("🔔 seguimiento " + it.get("fecha_seg", "")),
                             text_color=(RED if venc else MUTED),
                             font=("Segoe UI", 9, "bold")).pack(side="left", padx=6)
            if pend:
                ctk.CTkLabel(top, text=f"{pend} tarea(s) pend.", text_color=RED,
                             font=("Segoe UI", 9, "bold")).pack(side="left", padx=4)
            linea2 = (f"Asesor: {it.get('asesor') or '-'}     Fecha: {it.get('fecha','')}"
                      f"     {dest}     {usd(it.get('total', 0))}")
            ctk.CTkLabel(cel, text=linea2, text_color=MUTED, anchor="w",
                         font=("Segoe UI", 10)).pack(anchor="w")
            ctk.CTkButton(fila, text="Editar / Tareas", width=110, height=30, corner_radius=8,
                          fg_color=CYAN, hover_color=BLUE, font=("Segoe UI", 11, "bold"),
                          command=lambda x=it: self._detalle(x)).grid(row=0, column=1, padx=4)
            if it.get("pdf"):
                ctk.CTkButton(fila, text="Abrir PDF", width=90, height=30, corner_radius=8,
                              fg_color=NAVY, hover_color=BLUE, font=("Segoe UI", 11, "bold"),
                              command=lambda p=it["pdf"]: self._abrir(p)).grid(row=0, column=2, padx=4)
            ctk.CTkButton(fila, text="🗑", width=34, height=30, corner_radius=8,
                          fg_color=CARD2, text_color=RED, hover_color=LINE,
                          font=("Segoe UI", 13),
                          command=lambda x=it: self._eliminar(x)).grid(row=0, column=3, padx=(0, 6))
        if not n:
            msg = ("Aun no hay cotizaciones guardadas.\nGenera un PDF y aparecera aqui."
                   if not self.data.get("items") else "Sin resultados.")
            ctk.CTkLabel(self.lista, text=msg, text_color=MUTED).pack(pady=24)

    def _importar_html(self):
        try:
            n = importar_cotizaciones_html()
        except Exception as e:
            messagebox.showerror("Importar", str(e), parent=self); return
        self.data = cargar_cotizaciones(); self._pintar()
        messagebox.showinfo("Importar del HTML",
                            (f"Se importaron {n} cotizacion(es) nueva(s) del HTML."
                             if n else "No hay cotizaciones nuevas del HTML."), parent=self)

    def _detalle(self, it):
        VentanaCotizacionDetalle(self, it, self._on_guardado)

    def _on_guardado(self):
        guardar_cotizaciones(self.data)
        self._pintar()

    def _abrir(self, p):
        if p and os.path.exists(p):
            try:
                os.startfile(p)
            except Exception as e:
                messagebox.showerror("Error", str(e))
        else:
            messagebox.showinfo("PDF no encontrado",
                                "El archivo PDF ya no esta en:\n" + str(p))

    def _eliminar(self, it):
        if messagebox.askyesno("Quitar del historial",
                               f"¿Quitar {it.get('numero','')} del historial?\n"
                               "(No borra el archivo PDF)", parent=self):
            self.data["items"] = [x for x in self.data["items"] if x is not it]
            guardar_cotizaciones(self.data)
            self._pintar()


# ============================================================================
# Ventana de Clientes / Empresas (con vendedores)
# ============================================================================
class VentanaClientes(ctk.CTkToplevel):
    def __init__(self, master, on_cambio=None, preseleccion=None):
        super().__init__(master)
        self.on_cambio = on_cambio
        self.title("Clientes / Empresas")
        self.geometry("980x640"); self.configure(fg_color=BG)
        self.transient(master); self.grab_set()
        self.clientes = cargar_clientes()
        self.sel = None            # indice seleccionado (None = nuevo)
        self.vend_rows = []        # filas de vendedores (dicts de StringVars)

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)
        # ---- barra superior ----
        top = ctk.CTkFrame(self, fg_color=CARD, corner_radius=0, height=52)
        top.grid(row=0, column=0, columnspan=2, sticky="ew")
        ctk.CTkLabel(top, text="Clientes / Empresas", text_color=NAVY,
                     font=("Segoe UI", 16, "bold")).pack(side="left", padx=16, pady=10)
        ctk.CTkButton(top, text="Importar Excel/CSV", width=150, height=32, corner_radius=8,
                      fg_color=NAVY, hover_color=BLUE, font=("Segoe UI", 12, "bold"),
                      command=self._importar).pack(side="right", padx=(4, 16))
        ctk.CTkButton(top, text="Exportar Excel", width=130, height=32, corner_radius=8,
                      fg_color=GREEN, hover_color=GREEN_H, font=("Segoe UI", 12, "bold"),
                      command=self._exportar).pack(side="right", padx=4)
        ctk.CTkButton(top, text="+ Nueva empresa", width=140, height=32, corner_radius=8,
                      fg_color=CARD2, text_color=NAVY, hover_color=LINE, border_width=1,
                      border_color=LINE, font=("Segoe UI", 12, "bold"),
                      command=self._nuevo).pack(side="right", padx=4)
        # ---- lista izquierda ----
        izq = ctk.CTkFrame(self, fg_color=CARD, corner_radius=12); izq.grid(
            row=1, column=0, sticky="nsew", padx=(12, 6), pady=12)
        izq.grid_rowconfigure(1, weight=1); izq.grid_columnconfigure(0, weight=1)
        self.var_busca = tk.StringVar()
        eb = ctk.CTkEntry(izq, textvariable=self.var_busca, height=32, corner_radius=8,
                          border_color=LINE, placeholder_text="Buscar empresa...", width=240)
        eb.grid(row=0, column=0, sticky="ew", padx=8, pady=8)
        self.var_busca.trace_add("write", lambda *a: self._rebuild_list())
        self.lista = ctk.CTkScrollableFrame(izq, fg_color=CARD2, corner_radius=10, width=240)
        self.lista.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
        # ---- form derecha ----
        self.form = ctk.CTkScrollableFrame(self, fg_color=CARD, corner_radius=12)
        self.form.grid(row=1, column=1, sticky="nsew", padx=(6, 12), pady=12)
        self.form.grid_columnconfigure(0, weight=1)
        self._build_form()
        self._rebuild_list()
        self._nuevo()
        if preseleccion:
            idx = next((i for i, c in enumerate(self.clientes)
                        if c.get("empresa", "") == preseleccion), None)
            if idx is not None:
                self.var_busca.set(preseleccion)
                self._cargar(idx)

    def _build_form(self):
        self.vars = {}
        def campo(clave, etq):
            ctk.CTkLabel(self.form, text=etq, text_color=MUTED,
                         font=("Segoe UI", 11)).pack(anchor="w", padx=12, pady=(6, 0))
            v = tk.StringVar(); self.vars[clave] = v
            ctk.CTkEntry(self.form, textvariable=v, height=32, corner_radius=8,
                         border_color=LINE).pack(fill="x", padx=12)
        ctk.CTkLabel(self.form, text="Datos de la empresa", text_color=NAVY,
                     font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=12, pady=(10, 2))
        campo("empresa", "Nombre de la empresa *")
        campo("nit", "NIT / Documento fiscal")
        campo("telefono", "Telefono")
        campo("email", "Email")
        campo("web", "Sitio web")
        campo("pais", "Pais")
        # vendedores
        hdr = ctk.CTkFrame(self.form, fg_color="transparent"); hdr.pack(fill="x", padx=12, pady=(12, 0))
        ctk.CTkLabel(hdr, text="Vendedores / contactos", text_color=NAVY,
                     font=("Segoe UI", 14, "bold")).pack(side="left")
        ctk.CTkButton(hdr, text="+ Agregar vendedor", width=150, height=28, corner_radius=8,
                      fg_color=CARD2, text_color=NAVY, hover_color=LINE, border_width=1,
                      border_color=LINE, font=("Segoe UI", 11, "bold"),
                      command=lambda: self._add_vend()).pack(side="right")
        self.vends_frame = ctk.CTkFrame(self.form, fg_color="transparent")
        self.vends_frame.pack(fill="x", padx=6, pady=4)
        # botones
        bts = ctk.CTkFrame(self.form, fg_color="transparent"); bts.pack(fill="x", padx=12, pady=14)
        ctk.CTkButton(bts, text="Guardar", fg_color=GREEN, hover_color=GREEN_H,
                      font=("Segoe UI", 13, "bold"), command=self._guardar).pack(side="left")
        ctk.CTkButton(bts, text="Eliminar", fg_color=CARD2, text_color=RED, hover_color=LINE,
                      command=self._eliminar).pack(side="left", padx=8)
        self.lbl_estado = ctk.CTkLabel(bts, text="", text_color=MUTED); self.lbl_estado.pack(side="left", padx=10)

    def _add_vend(self, v=None):
        row = ctk.CTkFrame(self.vends_frame, fg_color=CARD2, corner_radius=8)
        row.pack(fill="x", padx=6, pady=3)
        vv = {k: tk.StringVar(value=(v or {}).get(k, "")) for k in ("nombre", "telefono", "email", "cargo")}
        for clave, ph, w in (("nombre", "Nombre", 150), ("telefono", "Telefono", 110),
                             ("email", "Email", 170), ("cargo", "Cargo", 110)):
            ctk.CTkEntry(row, textvariable=vv[clave], height=28, width=w, corner_radius=6,
                         border_color=LINE, placeholder_text=ph).pack(side="left", padx=3, pady=4)
        ctk.CTkButton(row, text="✕", width=26, height=26, corner_radius=6, fg_color="transparent",
                      text_color=RED, hover_color=LINE,
                      command=lambda: (self.vend_rows.remove(entry), row.destroy())).pack(side="left", padx=2)
        entry = {"frame": row, "vars": vv}
        self.vend_rows.append(entry)

    def _clear_vends(self):
        for e in self.vend_rows:
            e["frame"].destroy()
        self.vend_rows = []

    def _rebuild_list(self):
        for w in self.lista.winfo_children():
            w.destroy()
        q = self.var_busca.get().lower()
        for i, c in enumerate(self.clientes):
            if q and q not in c.get("empresa", "").lower():
                continue
            act = (i == self.sel)
            b = ctk.CTkButton(self.lista, text=c.get("empresa", "(sin nombre)"),
                              anchor="w", height=32, corner_radius=8,
                              fg_color=NAVY if act else CARD, text_color="#FFFFFF" if act else NAVY,
                              hover_color=BLUE, font=("Segoe UI", 12, "bold" if act else "normal"),
                              command=lambda x=i: self._cargar(x))
            b.pack(fill="x", padx=4, pady=2)
        if not self.clientes:
            ctk.CTkLabel(self.lista, text="Sin clientes.\nCrea uno o importa un Excel.",
                         text_color=MUTED).pack(pady=20)

    def _cargar(self, idx):
        self.sel = idx
        c = self.clientes[idx]
        for k, v in self.vars.items():
            v.set(c.get(k, ""))
        self._clear_vends()
        for vend in c.get("vendedores", []):
            self._add_vend(vend)
        self.lbl_estado.configure(text="")
        self._rebuild_list()

    def _nuevo(self):
        self.sel = None
        for v in self.vars.values():
            v.set("")
        self._clear_vends(); self._add_vend()
        self.lbl_estado.configure(text="Nueva empresa")
        self._rebuild_list()

    def _recoger(self):
        c = {k: v.get().strip() for k, v in self.vars.items()}
        c["vendedores"] = []
        for e in self.vend_rows:
            d = {k: e["vars"][k].get().strip() for k in ("nombre", "telefono", "email", "cargo")}
            if d["nombre"]:
                c["vendedores"].append(d)
        return c

    def _guardar(self):
        c = self._recoger()
        if not c["empresa"]:
            messagebox.showwarning("Falta el nombre", "Escribe el nombre de la empresa."); return
        if self.sel is None:
            self.clientes.append(c); self.sel = len(self.clientes) - 1
        else:
            self.clientes[self.sel] = c
        guardar_clientes(self.clientes)
        self.lbl_estado.configure(text="Guardado ✓")
        self._rebuild_list()
        if self.on_cambio:
            self.on_cambio()

    def _eliminar(self):
        if self.sel is None:
            return
        if messagebox.askyesno("Eliminar", "¿Eliminar esta empresa?"):
            del self.clientes[self.sel]
            guardar_clientes(self.clientes)
            self._nuevo()
            if self.on_cambio:
                self.on_cambio()

    def _importar(self):
        ruta = filedialog.askopenfilename(title="Importar clientes de Excel/CSV",
                                          filetypes=[("Excel/CSV", "*.xlsx *.xlsm *.csv *.txt"),
                                                     ("Excel", "*.xlsx *.xlsm"),
                                                     ("CSV", "*.csv *.txt")])
        if not ruta:
            return
        try:
            nuevos = importar_clientes_excel(ruta)
        except Exception as e:
            messagebox.showerror("Error al importar", str(e)); return
        if not nuevos:
            messagebox.showinfo("Importar", "No se encontraron empresas en el archivo."); return
        # combinar por nombre (actualiza / agrega)
        idx = {c["empresa"].lower(): i for i, c in enumerate(self.clientes)}
        add = upd = 0
        for c in nuevos:
            k = c["empresa"].lower()
            if k in idx:
                self.clientes[idx[k]] = c; upd += 1
            else:
                self.clientes.append(c); idx[k] = len(self.clientes) - 1; add += 1
        guardar_clientes(self.clientes)
        self._nuevo()
        if self.on_cambio:
            self.on_cambio()
        messagebox.showinfo("Importado", f"Importados: {add} nuevos, {upd} actualizados.")

    def _exportar(self):
        ruta = filedialog.asksaveasfilename(title="Exportar clientes a Excel",
                                            defaultextension=".xlsx", initialfile="Clientes_Innoba.xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if not ruta:
            return
        try:
            exportar_clientes_excel(self.clientes, ruta)
        except Exception as e:
            messagebox.showerror("Error al exportar", str(e)); return
        if messagebox.askyesno("Exportado", "Excel guardado:\n" + ruta + "\n\n¿Abrirlo?"):
            try: os.startfile(ruta)
            except Exception: pass


# ============================================================================
# Calendario (selector de fecha)
# ============================================================================
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
         "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
DIAS_SEM = ["Lu", "Ma", "Mi", "Ju", "Vi", "Sa", "Do"]


class CalendarioPopup(ctk.CTkToplevel):
    def __init__(self, master, fecha_ini, minimo, on_pick):
        super().__init__(master)
        self.on_pick = on_pick
        self.minimo = minimo
        self.title("Elegir fecha")
        self.configure(fg_color=CARD)
        self.resizable(False, False)
        self.transient(master)
        base = fecha_ini or minimo or datetime.date.today()
        self.y, self.m = base.year, base.month
        self.cont = ctk.CTkFrame(self, fg_color=CARD)
        self.cont.pack(padx=10, pady=10)
        self._build()
        self.after(60, self._centrar)
        self.grab_set()

    def _centrar(self):
        try:
            self.update_idletasks()
            w = self.winfo_width(); h = self.winfo_height()
            sw = self.winfo_screenwidth(); sh = self.winfo_screenheight()
            mx = self.master.winfo_pointerx(); my = self.master.winfo_pointery()
            x = min(max(mx - 120, 10), sw - w - 10)
            y = min(max(my + 12, 10), sh - h - 10)
            self.geometry(f"+{int(x)}+{int(y)}")
        except Exception:
            pass

    def _build(self):
        for w in self.cont.winfo_children():
            w.destroy()
        top = ctk.CTkFrame(self.cont, fg_color=NAVY, corner_radius=8)
        top.pack(fill="x", pady=(0, 8))
        ctk.CTkButton(top, text="◀", width=34, height=30, fg_color=NAVY2,
                      hover_color=BLUE, command=self._prev).pack(side="left", padx=4, pady=4)
        ctk.CTkLabel(top, text=f"{MESES[self.m-1]} {self.y}", text_color="#FFFFFF",
                     font=("Segoe UI", 13, "bold")).pack(side="left", expand=True)
        ctk.CTkButton(top, text="▶", width=34, height=30, fg_color=NAVY2,
                      hover_color=BLUE, command=self._next).pack(side="right", padx=4, pady=4)
        grid = ctk.CTkFrame(self.cont, fg_color=CARD)
        grid.pack()
        for i, d in enumerate(DIAS_SEM):
            ctk.CTkLabel(grid, text=d, text_color=MUTED, width=36,
                         font=("Segoe UI", 10, "bold")).grid(row=0, column=i, padx=1, pady=2)
        cal = calendar.Calendar(firstweekday=0)   # lunes
        for r, semana in enumerate(cal.monthdatescalendar(self.y, self.m), start=1):
            for c, dia in enumerate(semana):
                del_mes = (dia.month == self.m)
                deshab = (self.minimo and dia < self.minimo)
                if del_mes and not deshab:
                    fg, tc = CARD2, NAVY
                else:
                    fg, tc = CARD, MUTED
                b = ctk.CTkButton(grid, text=str(dia.day), width=36, height=30,
                                  corner_radius=6, fg_color=fg, text_color=tc,
                                  hover_color=BLUE, font=("Segoe UI", 11),
                                  command=lambda dd=dia: self._pick(dd))
                if deshab:
                    b.configure(state="disabled")
                b.grid(row=r, column=c, padx=1, pady=1)

    def _prev(self):
        self.m -= 1
        if self.m < 1:
            self.m = 12; self.y -= 1
        self._build()

    def _next(self):
        self.m += 1
        if self.m > 12:
            self.m = 1; self.y += 1
        self._build()

    def _pick(self, dia):
        self.on_pick(dia)
        self.destroy()


class SelectorFecha(ctk.CTkFrame):
    def __init__(self, master, command=None, minimo=None, **kw):
        super().__init__(master, fg_color="transparent", **kw)
        self.command = command
        self.minimo = minimo
        self._fecha = None
        self.btn = ctk.CTkButton(self, text="📅  Elegir...", height=34, corner_radius=8,
                                 fg_color=CARD2, text_color=NAVY, hover_color=LINE,
                                 border_width=1, border_color=LINE, font=("Segoe UI", 12),
                                 anchor="w", command=self._abrir)
        self.btn.pack(fill="x")

    def _abrir(self):
        CalendarioPopup(self.winfo_toplevel(), self._fecha, self.minimo, self._set)

    def _set(self, dia):
        self._fecha = dia
        self.btn.configure(text="📅  " + dia.strftime("%d/%m/%Y"),
                           fg_color="#E7F0FB", text_color=NAVY)
        if self.command:
            self.command()

    def get(self):
        return self._fecha

    def get_str(self):
        return self._fecha.strftime("%d/%m/%Y") if self._fecha else ""

    def clear(self):
        self._fecha = None
        self.btn.configure(text="📅  Elegir...", fg_color=CARD2, text_color=NAVY)


# ============================================================================
# Aplicacion
# ============================================================================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light")
        # Escala mas compacta: letra y controles ~15% mas pequenos -> se ven mas
        # opciones de hotel/terrestres y todo cabe mejor en pantalla.
        try:
            ctk.set_widget_scaling(0.85)
        except Exception:
            pass
        self.cfg = cargar_config()
        respaldar_datos(self.cfg)   # respaldo automatico de cotizaciones/clientes
        if not self.cfg.get("logo") or not os.path.exists(self.cfg["logo"]):
            lg = recurso("logo_innoba.png")
            if os.path.exists(lg):
                self.cfg["logo"] = lg
        try:
            self.precios = cargar_precios()
        except Exception as e:
            messagebox.showerror("Error", str(e)); self.destroy(); return
        self.descripciones = cargar_descripciones()
        self.clientes = cargar_clientes()

        self.title(f"Cotizador INNOBA Colombia DMC   v{VERSION}")
        self.geometry("1180x820"); self.minsize(1040, 620)
        self.configure(fg_color=BG)
        try:
            self.iconbitmap(recurso("app.ico"))
        except Exception:
            pass

        self.tramos = []          # itinerario: lista de destinos
        self.activo = None        # indice del destino activo
        self._cargando = False    # evita feedback al cargar widgets
        self.tab = "hotel"        # pestana activa
        self.q = {"hotel": "", "trans": "", "act": ""}   # textos de busqueda
        self._itinerario = ""     # itinerario dia por dia (editable)

        # TRM CRUDA (sin descuento). El descuento se aplica segun la fecha de viaje.
        self._trm = _trm_valida(self.cfg.get("ultima_trm", ""))
        self.var_trm_status = tk.StringVar(value="Consultando tasa...")

        self._construir()
        self._nueva()
        # abrir maximizada para que TODO (incluido el pie con Generar PDF) sea visible
        try:
            self.after(60, lambda: self.state("zoomed"))
        except Exception:
            pass
        self.after(250, lambda: self._actualizar_trm(silencioso=True))
        self.after(1800, self._chequear_actualizacion)   # busca nueva version en el repo
        self.after(1200, self._chequear_seguimientos)    # alerta de cotizaciones a seguir

    # ------------------------------------------------------------------ UI
    def _construir(self):
        self.grid_columnconfigure(0, weight=1)
        # el panel de seleccion tiene una altura minima decente (muestra varias
        # opciones); como la app abre maximizada, el pie tambien cabe.
        self.grid_rowconfigure(5, weight=1, minsize=240)

        # Encabezado
        head = ctk.CTkFrame(self, fg_color=CARD, corner_radius=0, height=60)
        head.grid(row=0, column=0, sticky="ew"); head.grid_propagate(False)
        head.grid_columnconfigure(1, weight=1)
        try:
            img = Image.open(recurso("logo_innoba.png")); w, h = img.size; hh = 40
            self.logo_img = ctk.CTkImage(light_image=img, size=(int(w * hh / h), hh))
            ctk.CTkLabel(head, image=self.logo_img, text="").grid(
                row=0, column=0, padx=(20, 14), pady=8)
        except Exception:
            ctk.CTkLabel(head, text="INNOBA", font=("Segoe UI", 22, "bold"),
                         text_color=NAVY).grid(row=0, column=0, padx=20)
        tit = ctk.CTkFrame(head, fg_color="transparent"); tit.grid(row=0, column=1, sticky="w")
        ctk.CTkLabel(tit, text="Cotizador de Paquetes", text_color=NAVY,
                     font=("Segoe UI", 17, "bold"), height=20).pack(anchor="w")
        ctk.CTkLabel(tit, text=f"INNOBA Colombia DMC  ·  v{VERSION}  ·  Itinerario hasta 5 destinos",
                     text_color=MUTED, font=("Segoe UI", 11), height=15).pack(anchor="w")
        hbtns = ctk.CTkFrame(head, fg_color="transparent"); hbtns.grid(row=0, column=2, padx=20)
        ctk.CTkButton(hbtns, text="Cotizaciones", width=120, height=36, corner_radius=10,
                      fg_color=GREEN, hover_color=GREEN_H, font=("Segoe UI", 12, "bold"),
                      command=self._abrir_cotizaciones).pack(side="left", padx=(0, 8))
        ctk.CTkButton(hbtns, text="Itinerario", width=110, height=36, corner_radius=10,
                      fg_color=CYAN, hover_color=BLUE, font=("Segoe UI", 12, "bold"),
                      command=self._abrir_itinerario).pack(side="left", padx=(0, 8))
        ctk.CTkButton(hbtns, text="Clientes", width=110, height=36, corner_radius=10,
                      fg_color=NAVY, hover_color=BLUE, font=("Segoe UI", 12, "bold"),
                      command=self._abrir_clientes).pack(side="left", padx=(0, 8))
        ctk.CTkButton(hbtns, text="Datos de mi empresa", width=170, height=36,
                      corner_radius=10, fg_color=CARD2, text_color=NAVY, hover_color=LINE,
                      border_width=1, border_color=LINE, font=("Segoe UI", 12, "bold"),
                      command=self._abrir_empresa).pack(side="left")

        # Barra de estado de la tasa (SIN mostrar el valor)
        trm = ctk.CTkFrame(self, fg_color=NAVY, corner_radius=0, height=34)
        trm.grid(row=1, column=0, sticky="ew"); trm.grid_propagate(False)
        ins = ctk.CTkFrame(trm, fg_color="transparent"); ins.pack(fill="both", expand=True,
                                                                  padx=20, pady=4)
        ctk.CTkLabel(ins, text="Tasa del dia (USD):", text_color="#BFD4F0",
                     font=("Segoe UI", 12)).pack(side="left")
        ctk.CTkLabel(ins, textvariable=self.var_trm_status, text_color="#FFFFFF",
                     font=("Segoe UI", 12, "bold")).pack(side="left", padx=(6, 0))
        ctk.CTkButton(ins, text="↻ Actualizar", width=110, height=28, corner_radius=8,
                      fg_color=BLUE, hover_color=CYAN, font=("Segoe UI", 11, "bold"),
                      command=lambda: self._actualizar_trm(False)).pack(side="right")

        # Datos globales del viaje
        g = ctk.CTkFrame(self, fg_color=CARD, corner_radius=14)
        g.grid(row=2, column=0, sticky="ew", padx=16, pady=(12, 6))
        for c in range(6):
            g.grid_columnconfigure(c, weight=1)
        ctk.CTkLabel(g, text="Datos del viaje (cliente)", text_color=NAVY,
                     font=("Segoe UI", 14, "bold"), height=18).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=16, pady=(8, 0))
        # Quien realiza la cotizacion
        cotz = ctk.CTkFrame(g, fg_color="transparent")
        cotz.grid(row=0, column=3, columnspan=3, sticky="e", padx=16, pady=(6, 0))
        ctk.CTkLabel(cotz, text="Cotizado por:", text_color=MUTED,
                     font=("Segoe UI", 11)).pack(side="left", padx=(0, 6))
        self._cotz_map = {f"{n}  -  {c}": (n, c) for n, c in COTIZADORES}
        ops = list(self._cotz_map.keys())
        self.var_cotizador = tk.StringVar(value=ops[0])
        ctk.CTkOptionMenu(cotz, variable=self.var_cotizador, values=ops, width=320, height=30,
                          corner_radius=8, fg_color=NAVY, button_color=NAVY2,
                          button_hover_color=BLUE, dropdown_fg_color=CARD,
                          dropdown_text_color=TEXT, font=("Segoe UI", 11)).pack(side="left")
        def lab(t, r, c, span=1):
            ctk.CTkLabel(g, text=t, text_color=MUTED, font=("Segoe UI", 10), height=13).grid(
                row=r, column=c, columnspan=span, sticky="w", padx=16, pady=0)
        lab("EMAIL CLIENTE", 1, 2)
        lab("FECHAS DEL VIAJE (ida - regreso)", 1, 4, 2)
        cli_lab = ctk.CTkFrame(g, fg_color="transparent")
        cli_lab.grid(row=1, column=0, columnspan=2, sticky="ew", padx=16, pady=0)
        ctk.CTkLabel(cli_lab, text="CLIENTE", text_color=MUTED,
                     font=("Segoe UI", 10), height=13).pack(side="left")
        ctk.CTkButton(cli_lab, text="✏ Editar", width=66, height=22, corner_radius=6,
                      fg_color=CARD2, text_color=NAVY, hover_color=LINE, border_width=1,
                      border_color=LINE, font=("Segoe UI", 10, "bold"),
                      command=self._editar_cliente_actual).pack(side="right", padx=(4, 0))
        ctk.CTkButton(cli_lab, text="🔍 Buscar", width=84, height=22, corner_radius=6,
                      fg_color=NAVY, hover_color=BLUE, font=("Segoe UI", 10, "bold"),
                      command=self._buscar_cliente).pack(side="right")
        self.var_cli = tk.StringVar(); self.var_email = tk.StringVar()
        ctk.CTkEntry(g, textvariable=self.var_cli, height=30, corner_radius=8,
                     border_color=LINE, fg_color=CARD2, placeholder_text="Nombre del cliente"
                     ).grid(row=2, column=0, columnspan=2, sticky="ew", padx=16, pady=(0, 6))
        ctk.CTkEntry(g, textvariable=self.var_email, height=30, corner_radius=8,
                     border_color=LINE, fg_color=CARD2,
                     placeholder_text="correo@cliente.com  (se enviara aqui)"
                     ).grid(row=2, column=2, columnspan=2, sticky="ew", padx=16, pady=(0, 6))
        fechas_fr = ctk.CTkFrame(g, fg_color="transparent")
        fechas_fr.grid(row=2, column=4, columnspan=2, sticky="ew", padx=16, pady=(0, 6))
        hoy = datetime.date.today()
        self.sel_desde = SelectorFecha(fechas_fr, command=self._on_fecha_desde, minimo=hoy)
        self.sel_desde.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(fechas_fr, text="al", text_color=MUTED).pack(side="left", padx=6)
        self.sel_hasta = SelectorFecha(fechas_fr, command=self._recalcular, minimo=hoy)
        self.sel_hasta.pack(side="left", fill="x", expand=True)
        # Asesor / contacto de la empresa
        aso_lab = ctk.CTkFrame(g, fg_color="transparent")
        aso_lab.grid(row=3, column=0, columnspan=2, sticky="ew", padx=16, pady=(2, 0))
        ctk.CTkLabel(aso_lab, text="ASESOR (contacto)", text_color=MUTED,
                     font=("Segoe UI", 10), height=13).pack(side="left")
        self.opt_asesor = ctk.CTkOptionMenu(
            aso_lab, values=["(vendedor)"], width=150, height=22, corner_radius=6,
            fg_color=NAVY, button_color=NAVY2, button_hover_color=BLUE, dropdown_fg_color=CARD,
            dropdown_text_color=TEXT, font=("Segoe UI", 10), command=self._usar_asesor)
        self.opt_asesor.pack(side="right")
        lab("TELEFONO ASESOR", 3, 2, 2)
        self.var_asesor = tk.StringVar(); self.var_aso_tel = tk.StringVar()
        ctk.CTkEntry(g, textvariable=self.var_asesor, height=30, corner_radius=8,
                     border_color=LINE, fg_color=CARD2, placeholder_text="Nombre del asesor"
                     ).grid(row=4, column=0, columnspan=2, sticky="ew", padx=16, pady=(0, 6))
        ctk.CTkEntry(g, textvariable=self.var_aso_tel, height=30, corner_radius=8,
                     border_color=LINE, fg_color=CARD2, placeholder_text="Telefono del asesor"
                     ).grid(row=4, column=2, columnspan=2, sticky="ew", padx=16, pady=(0, 6))
        lab("ADULTOS", 5, 0); lab("NINOS", 5, 1)
        lab("HABITACIONES (Sen/Dob/Tri)", 5, 2, 2)
        lab("EDAD DE CADA NINO", 5, 4, 2)
        self.st_ad = Stepper(g, value=2, minimo=1, maximo=60, command=self._on_pax)
        self.st_ad.grid(row=6, column=0, sticky="w", padx=16, pady=(0, 2))
        self.st_ninos = Stepper(g, value=0, minimo=0, maximo=10, command=self._on_ninos_count)
        self.st_ninos.grid(row=6, column=1, sticky="w", padx=16, pady=(0, 2))
        # Habitaciones al lado de ninos: 3 contadores compactos + Sugerir
        habf = ctk.CTkFrame(g, fg_color="transparent")
        habf.grid(row=6, column=2, columnspan=2, sticky="w", padx=16, pady=(0, 2))
        for attr, val in (("st_hab_s", 0), ("st_hab_d", 1), ("st_hab_t", 0)):
            st = Stepper(habf, value=val, minimo=0, maximo=40, width=84,
                         command=self._on_hab_change)
            st.pack(side="left", padx=(0, 4)); setattr(self, attr, st)
        ctk.CTkButton(habf, text="Sugerir", width=72, height=28, corner_radius=8,
                      fg_color=CARD2, text_color=NAVY, hover_color=LINE, border_width=1,
                      border_color=LINE, font=("Segoe UI", 10, "bold"),
                      command=self._sugerir_hab).pack(side="left", padx=(6, 0))
        self.frame_edades = ctk.CTkFrame(g, fg_color=CARD2, corner_radius=8, height=34)
        self.frame_edades.grid(row=6, column=4, columnspan=2, sticky="ew", padx=16, pady=(0, 2))
        self.frame_edades.grid_propagate(False)
        self.edad_vars = []
        self.lbl_hab = ctk.CTkLabel(g, text="", text_color=MUTED, font=("Segoe UI", 10), height=12)
        self.lbl_hab.grid(row=7, column=0, columnspan=4, sticky="w", padx=16, pady=(0, 4))

        # Barra de destinos (itinerario)
        dst = ctk.CTkFrame(self, fg_color="#E4EDFA", corner_radius=14)
        dst.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 6))
        dst.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(dst, text="Destinos del itinerario (max 5):", text_color=NAVY,
                     font=("Segoe UI", 13, "bold")).grid(row=0, column=0, padx=14, pady=10,
                                                         sticky="w")
        self.chips = ctk.CTkFrame(dst, fg_color="transparent")
        self.chips.grid(row=0, column=1, sticky="w", padx=4, pady=6)
        self.var_add = tk.StringVar(value="+ Agregar destino")
        self.opt_add = ctk.CTkOptionMenu(
            dst, variable=self.var_add, values=list(self.precios.keys()),
            width=190, height=34, corner_radius=8, fg_color=GREEN, button_color=GREEN_H,
            button_hover_color=CYAN, dropdown_fg_color=CARD, dropdown_text_color=TEXT,
            font=("Segoe UI", 12, "bold"), command=self._add_destino)
        self.opt_add.grid(row=0, column=2, padx=14, pady=8)

        # Configuracion del destino activo
        cfgd = ctk.CTkFrame(self, fg_color="transparent")
        cfgd.grid(row=4, column=0, sticky="ew", padx=16)
        self.lbl_activo = ctk.CTkLabel(cfgd, text="", text_color=NAVY,
                                       font=("Segoe UI", 13, "bold"))
        self.lbl_activo.pack(side="left", padx=(4, 16))
        ctk.CTkLabel(cfgd, text="Temporada:", text_color=MUTED).pack(side="left")
        self.var_temp = tk.StringVar()
        self.opt_temp = ctk.CTkOptionMenu(cfgd, variable=self.var_temp, values=["Baja"],
                                          width=140, height=32, corner_radius=8, fg_color=NAVY,
                                          button_color=NAVY2, button_hover_color=BLUE,
                                          dropdown_fg_color=CARD, dropdown_text_color=TEXT,
                                          font=("Segoe UI", 12, "bold"),
                                          command=self._on_temp)
        self.opt_temp.pack(side="left", padx=8)
        ctk.CTkLabel(cfgd, text="Noches:", text_color=MUTED).pack(side="left", padx=(12, 2))
        self.st_noches = Stepper(cfgd, value=3, minimo=1, maximo=60, command=self._on_noches)
        self.st_noches.pack(side="left")

        # Pestanas propias (robustas) + panel del destino activo
        mid = ctk.CTkFrame(self, fg_color="transparent")
        mid.grid(row=5, column=0, sticky="nsew", padx=16, pady=(2, 4))
        mid.grid_columnconfigure(0, weight=1)
        mid.grid_rowconfigure(1, weight=1)
        tabbar = ctk.CTkFrame(mid, fg_color="transparent")
        tabbar.grid(row=0, column=0, sticky="w")
        self._tab_btns = {}
        for key, txt in (("hotel", "  Hotel  "), ("trans", "  Transportes  "),
                         ("act", "  Actividades  ")):
            b = ctk.CTkButton(tabbar, text=txt, height=36, corner_radius=10,
                              font=("Segoe UI", 12, "bold"),
                              command=lambda k=key: self._set_tab(k))
            b.pack(side="left", padx=(0, 4))
            self._tab_btns[key] = b
        self.panel = ctk.CTkFrame(mid, fg_color=CARD, corner_radius=12)
        self.panel.grid(row=1, column=0, sticky="nsew", pady=(4, 0))
        self.panel.grid_columnconfigure(0, weight=1)
        self.panel.grid_rowconfigure(1, weight=1)

        # Barra total
        foot = ctk.CTkFrame(self, fg_color=NAVY, corner_radius=0, height=64)
        foot.grid(row=6, column=0, sticky="ew"); foot.grid_propagate(False)
        fin = ctk.CTkFrame(foot, fg_color="transparent"); fin.pack(fill="both", expand=True,
                                                                   padx=20, pady=6)
        ctk.CTkButton(fin, text="Nueva cotizacion", width=140, height=40, corner_radius=10,
                      fg_color=NAVY2, hover_color=BLUE, font=("Segoe UI", 12, "bold"),
                      command=self._nueva).pack(side="left")
        ctk.CTkButton(fin, text="Generar PDF y enviar  ✉", width=220, height=40,
                      corner_radius=10, fg_color=GREEN, hover_color=GREEN_H,
                      font=("Segoe UI", 14, "bold"),
                      command=self._generar).pack(side="right")
        ct = ctk.CTkFrame(fin, fg_color="transparent"); ct.pack(side="right", padx=24)
        self.lbl_total = ctk.CTkLabel(ct, text="USD 0.00", text_color="#FFFFFF",
                                      font=("Segoe UI", 22, "bold"), height=26); self.lbl_total.pack(anchor="e")
        self.lbl_desglose = ctk.CTkLabel(ct, text="Total del itinerario", text_color="#BFD4F0",
                                         font=("Segoe UI", 10), height=13); self.lbl_desglose.pack(anchor="e")

        self._rebuild_edades()      # muestra "Sin ninos"
        self._refrescar_clientes_picker()
        self._set_tab("hotel")      # estiliza botones y renderiza panel

    OCC = {"sencilla": 1, "doble": 2, "triple": 3}

    def _set_tab(self, key):
        self.tab = key
        for k, b in self._tab_btns.items():
            if k == key:
                b.configure(fg_color=NAVY, text_color="#FFFFFF", hover_color=BLUE)
            else:
                b.configure(fg_color=CARD2, text_color=NAVY, hover_color=LINE)
        self._render_panel()

    def _render_panel(self):
        for w in self.panel.winfo_children():
            w.destroy()
        if self.activo is None:
            ctk.CTkLabel(self.panel, text="Agrega un destino para comenzar.",
                         text_color=MUTED, font=("Segoe UI", 13)).grid(row=0, column=0, pady=30)
            return
        if self.tab == "hotel":
            self._render_hotel()
        else:
            self._render_lista(self.tab)

    # ---- panel HOTEL: lista con precio POR PERSONA ----
    def _render_hotel(self):
        tr = self.tramos[self.activo]
        bar = ctk.CTkFrame(self.panel, fg_color=CARD2, corner_radius=8)
        bar.grid(row=0, column=0, sticky="ew", padx=8, pady=(6, 4))
        ctk.CTkLabel(bar, text="Buscar hotel:", text_color=MUTED,
                     font=("Segoe UI", 11)).pack(side="left", padx=(10, 6), pady=5)
        eb = ctk.CTkEntry(bar, width=220, height=26, corner_radius=8, border_color=LINE,
                          placeholder_text="nombre del hotel...")
        eb.pack(side="left"); eb.insert(0, self.q["hotel"])
        eb.bind("<KeyRelease>", lambda e: (self.q.__setitem__("hotel", eb.get()),
                                           self._fill_hoteles(tr)))
        ctk.CTkLabel(bar, text="Precio POR PERSONA en Sencilla / Doble / Triple",
                     text_color=MUTED, font=("Segoe UI", 9)).pack(side="right", padx=12)
        self._hcont = ctk.CTkScrollableFrame(self.panel, fg_color=CARD2, corner_radius=10)
        self._hcont.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
        self._hcont.grid_columnconfigure(0, weight=1)
        self._fill_hoteles(tr)

    def _margenes(self, dd):
        """Margenes efectivos (hotel, terrestre) segun el periodo del viaje."""
        _, mh_over, mt_over = self._periodo()
        mh = mh_over if mh_over else (dd["hoteles"]["margen"] if dd.get("hoteles") else 0.88)
        mt = mt_over if mt_over else (dd["terrestres"]["margen"] if dd.get("terrestres") else 0.75)
        return mh, mt

    def _fill_hoteles(self, tr):
        for w in self._hcont.winfo_children():
            w.destroy()
        tasa = self._tasa()
        dd = self.precios.get(tr["destino"], {})
        mh = self._margenes(dd)[0]
        noches = max(tr["noches"], 1)
        hoteles, nombres = self._hoteles_de(tr["destino"], tr["temporada"])
        # limpiar seleccionados que ya no existen
        tr["hoteles"] = [n for n in tr["hoteles"] if n in nombres]
        info = ctk.CTkLabel(self._hcont,
                            text=f"Marca hasta 5 hoteles como opciones  "
                                 f"({len(tr['hoteles'])}/5 elegidos)",
                            text_color=NAVY, font=("Segoe UI", 10, "bold"))
        info.pack(anchor="w", padx=8, pady=(2, 4))
        q = (self.q["hotel"] or "").lower()
        n_mostrados = 0
        for h, nom in zip(hoteles, nombres):
            if q and q not in nom.lower():
                continue
            sel = (nom in tr["hoteles"])
            cat = h.get("categoria", "")
            row = ctk.CTkFrame(self._hcont, fg_color="#E7F0FB" if sel else CARD,
                               corner_radius=6, border_width=2 if sel else 0, border_color=NAVY)
            row.pack(fill="x", padx=6, pady=1)
            inner = ctk.CTkFrame(row, fg_color="transparent")
            inner.pack(fill="x", padx=8, pady=2)
            ctk.CTkLabel(inner, text=("☑ " if sel else "☐ ") + nom, text_color=NAVY,
                         font=("Segoe UI", 10, "bold"), anchor="w").pack(side="left")
            if cat:
                ctk.CTkLabel(inner, text=cat, text_color="#FFFFFF", fg_color=CYAN, corner_radius=6,
                             font=("Segoe UI", 8, "bold"), width=54, height=15).pack(side="left", padx=6)

            def pp(k, hh=h):
                v = hh.get(k)
                if not v or not tasa:
                    return "N/D"
                return usd(precio_hotel_usd_noche(v, tasa, mh) * noches / self.OCC[k])
            ctk.CTkLabel(inner, text=f"Sen {pp('sencilla')}    Dob {pp('doble')}    Tri {pp('triple')}",
                         text_color=TEXT, font=("Segoe UI", 10), anchor="e").pack(side="right")
            for wdg in (row, inner) + tuple(inner.winfo_children()):
                wdg.bind("<Button-1>", lambda e, n=nom: self._toggle_hotel(n))
            n_mostrados += 1
        if not n_mostrados:
            ctk.CTkLabel(self._hcont, text="No hay hoteles para este filtro.",
                         text_color=MUTED).pack(pady=20)

    def _toggle_hotel(self, nom):
        tr = self.tramos[self.activo]
        if nom in tr["hoteles"]:
            tr["hoteles"].remove(nom)
        else:
            if len(tr["hoteles"]) >= 5:
                messagebox.showinfo("Limite", "Puedes elegir hasta 5 hoteles como opciones.")
                return
            tr["hoteles"].append(nom)
        self._fill_hoteles(tr)
        self._recalcular()

    def _on_hab(self, k, val):
        if self.activo is None:
            return
        self.tramos[self.activo]["hab"][k] = val
        if self.tab == "hotel" and hasattr(self, "_hcont") and self._hcont.winfo_exists():
            self._fill_hoteles(self.tramos[self.activo])
        self._recalcular()

    # ---- panel TRANSPORTES / ACTIVIDADES: lista con precio POR PERSONA ----
    def _render_lista(self, tipo):
        tr = self.tramos[self.activo]
        bar = ctk.CTkFrame(self.panel, fg_color=CARD2, corner_radius=8)
        bar.grid(row=0, column=0, sticky="ew", padx=8, pady=(6, 4))
        ctk.CTkLabel(bar, text="Buscar:", text_color=MUTED,
                     font=("Segoe UI", 11)).pack(side="left", padx=(10, 6), pady=5)
        eb = ctk.CTkEntry(bar, width=240, height=26, corner_radius=8, border_color=LINE,
                          placeholder_text="nombre del servicio...")
        eb.pack(side="left"); eb.insert(0, self.q[tipo])
        eb.bind("<KeyRelease>", lambda e: (self.q.__setitem__(tipo, eb.get()),
                                           self._fill_lista(tr, tipo)))
        ctk.CTkLabel(bar, text="(valor por persona a la derecha)", text_color=MUTED,
                     font=("Segoe UI", 9)).pack(side="right", padx=12)
        self._lcont = ctk.CTkScrollableFrame(self.panel, fg_color=CARD2, corner_radius=10)
        self._lcont.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
        self._lcont.grid_columnconfigure(0, weight=1)
        self._fill_lista(tr, tipo)

    def _fill_lista(self, tr, tipo):
        for w in self._lcont.winfo_children():
            w.destroy()
        sel = tr[tipo]
        tasa = self._tasa()
        dd = self.precios.get(tr["destino"], {})
        margen = self._margenes(dd)[1]
        ages = self._ninos_ages()
        N = max(self.st_ad.get() + sum(1 for a in ages if a >= 1), 1)
        q = (self.q[tipo] or "").lower()
        n_mostrados = 0
        for serv in self._servicios_por_tipo(tr["destino"], tipo):
            nombre = serv["nombre"]
            if q and q not in nombre.lower():
                continue
            fila = ctk.CTkFrame(self._lcont, fg_color=CARD, corner_radius=6)
            fila.pack(fill="x", padx=6, pady=1)
            fila.grid_columnconfigure(0, weight=1)
            chk = tk.BooleanVar(value=(nombre in sel))
            txt = nombre + ("   (privado)" if es_privado(nombre) else "")

            def on_toggle(n=nombre, cv=chk, s=sel):
                if cv.get():
                    s.add(n)
                else:
                    s.discard(n)
                self._recalcular()

            ctk.CTkCheckBox(fila, text=txt, variable=chk, onvalue=True, offvalue=False,
                            corner_radius=5, fg_color=NAVY, hover_color=BLUE, text_color=TEXT,
                            font=("Segoe UI", 10), command=on_toggle, height=18,
                            checkbox_width=18, checkbox_height=18).grid(
                row=0, column=0, sticky="w", padx=10, pady=3)
            pp = (precio_terrestre_usd(serv, N, tasa, margen) / N) if tasa else None
            ctk.CTkLabel(fila, text=(usd(pp) + " p/p") if pp else "", text_color=NAVY,
                         font=("Segoe UI", 10, "bold")).grid(row=0, column=1, padx=10)
            n_mostrados += 1
        if not n_mostrados:
            ctk.CTkLabel(self._lcont, text="No hay servicios para este filtro.",
                         text_color=MUTED).pack(pady=20)

    # ------------------------------------------------------------- TRM / periodo
    def _periodo(self):
        """(descuento, margen_hotel, margen_terrestre) segun la fecha de ida."""
        return periodo_por_fecha(self.sel_desde.get() if hasattr(self, "sel_desde") else None)

    def _tasa(self):
        """Dolar aplicado = TRM del dia - descuento del periodo (segun fecha de viaje)."""
        if not self._trm or self._trm < 1000:
            return None
        desc = self._periodo()[0]
        tasa = self._trm - desc
        return tasa if tasa > 0 else None   # nunca tasa negativa -> nunca precios negativos

    def _actualizar_trm(self, silencioso=True):
        self.var_trm_status.set("actualizando...")
        def worker():
            trm = obtener_trm()
            self.after(0, lambda: self._aplicar_trm(trm, silencioso))
        threading.Thread(target=worker, daemon=True).start()

    def _aplicar_trm(self, trm, silencioso):
        hoy = datetime.date.today().strftime("%d/%m/%Y")
        if trm:
            self._trm = trm
            self.var_trm_status.set("actualizada " + hoy + "  ✓")
            self.cfg["ultima_trm"] = str(trm); self.cfg["ultima_trm_fecha"] = hoy
            try: guardar_config(self.cfg)
            except Exception: pass
            if self.activo is not None:
                self._render_panel()
            self._recalcular()
        else:
            if self._trm:
                self.var_trm_status.set("sin conexion (usando ultima del "
                                        + self.cfg.get("ultima_trm_fecha", "?") + ")")
                if self.activo is not None:
                    self._render_panel()
                self._recalcular()
            else:
                self.var_trm_status.set("no disponible - conectate a internet")
            if not silencioso:
                messagebox.showwarning("Tasa no disponible",
                                       "No se pudo consultar la tasa del dia. "
                                       "Verifica tu conexion y pulsa 'Actualizar'.")

    # ------------------------------------------------------- actualizaciones
    def _chequear_actualizacion(self):
        """En segundo plano, compara la version instalada con la del repositorio."""
        def worker():
            info = hay_actualizacion()
            if info:
                self.after(0, lambda: self._ofrecer_actualizacion(info))
        threading.Thread(target=worker, daemon=True).start()

    def _ofrecer_actualizacion(self, info):
        ver = info.get("version", "?")
        notas = (info.get("notas", "") or "").strip()
        url = info.get("installer", "")
        msg = (f"Hay una nueva version disponible: v{ver}\n"
               f"(tienes instalada la v{VERSION})\n")
        if notas:
            msg += f"\nNovedades:\n{notas}\n"
        msg += "\n¿Descargar e instalar ahora?"
        if messagebox.askyesno("Actualizacion disponible", msg):
            self._descargar_e_instalar(url, ver)

    def _descargar_e_instalar(self, url, ver):
        if not url:
            messagebox.showinfo("Actualizacion",
                                "No se encontro el instalador en el repositorio.\n"
                                "Descarga la ultima version desde GitHub.")
            return
        def worker():
            try:
                destino = os.path.join(tempfile.gettempdir(),
                                       f"CotizadorInnoba-Setup-{ver}.exe")
                ctx = ssl.create_default_context()
                req = urllib.request.Request(url,
                                             headers={"User-Agent": "CotizadorInnoba"})
                with urllib.request.urlopen(req, context=ctx, timeout=180) as r, \
                        open(destino, "wb") as f:
                    shutil.copyfileobj(r, f)
                self.after(0, lambda: self._lanzar_instalador(destino))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror(
                    "Error al descargar", f"No se pudo descargar la actualizacion:\n{e}"))
        self.var_trm_status.set("Descargando actualizacion...")
        threading.Thread(target=worker, daemon=True).start()

    def _lanzar_instalador(self, ruta):
        # avisamos ANTES de abrir el instalador
        messagebox.showinfo("Actualizacion",
                            "La aplicacion se cerrara y se abrira el instalador "
                            "para completar la actualizacion.")
        try:
            os.startfile(ruta)   # abre el instalador
        except Exception as e:
            messagebox.showerror("Error al abrir el instalador", str(e))
            return
        # cierre INMEDIATO del proceso para liberar el .exe (evita el error de
        # "acceso denegado" al reemplazarlo). El instalador ademas cierra la app.
        os._exit(0)

    # ------------------------------------------------------------- destinos
    def _temporadas_de(self, destino):
        dd = self.precios.get(destino, {})
        temps = []
        if dd.get("hoteles"):
            for h in dd["hoteles"]["hoteles"]:
                t = (h.get("temporada", "") or "Baja").strip()
                if t not in temps:
                    temps.append(t)
        return temps or ["Baja"]

    def _hoteles_de(self, destino, temporada):
        dd = self.precios.get(destino, {}); hoteles = []; nombres = []
        if dd.get("hoteles"):
            for h in dd["hoteles"]["hoteles"]:
                ht = (h.get("temporada", "") or "Baja").strip()
                if ht == temporada:
                    hoteles.append(h)
                    zona = (" [" + h["zona"] + "]") if h.get("zona") else ""
                    nombres.append(h["nombre"] + zona)
        return hoteles, nombres

    def _add_destino(self, nombre):
        self.var_add.set("+ Agregar destino")
        if nombre not in self.precios:
            return
        for i, tr in enumerate(self.tramos):
            if tr["destino"] == nombre:
                self._set_activo(i); return
        if len(self.tramos) >= 5:
            messagebox.showinfo("Limite", "Puedes cotizar hasta 5 destinos.")
            return
        temp = self._temporadas_de(nombre)[0]
        self.tramos.append({"destino": nombre, "temporada": temp, "noches": 3,
                            "hoteles": [], "hab": {"s": 0, "d": 1, "t": 0},
                            "trans": set(), "act": set()})
        self._rebuild_chips()
        self._set_activo(len(self.tramos) - 1)

    def _remove_destino(self, idx):
        if 0 <= idx < len(self.tramos):
            del self.tramos[idx]
            if not self.tramos:
                self.activo = None
            elif self.activo >= len(self.tramos):
                self.activo = len(self.tramos) - 1
            self._rebuild_chips()
            if self.activo is not None:
                self._cargar_activo()
            else:
                self._limpiar_activo()
            self._recalcular()

    def _rebuild_chips(self):
        for w in self.chips.winfo_children():
            w.destroy()
        for i, tr in enumerate(self.tramos):
            activo = (i == self.activo)
            chip = ctk.CTkFrame(self.chips, fg_color=NAVY if activo else CARD,
                                corner_radius=16)
            chip.pack(side="left", padx=4)
            ctk.CTkButton(chip, text=f"{i+1}. {tr['destino']}",
                          fg_color="transparent", hover_color=BLUE if activo else CARD2,
                          text_color="#FFFFFF" if activo else NAVY, height=30, width=10,
                          font=("Segoe UI", 12, "bold"),
                          command=lambda x=i: self._set_activo(x)).pack(side="left", padx=(8, 0))
            ctk.CTkButton(chip, text="✕", width=24, height=26, corner_radius=13,
                          fg_color="transparent", hover_color=RED,
                          text_color="#FFFFFF" if activo else MUTED,
                          font=("Segoe UI", 12, "bold"),
                          command=lambda x=i: self._remove_destino(x)).pack(side="left", padx=2)
        # limite de agregar
        disponibles = [d for d in self.precios if d not in [t["destino"] for t in self.tramos]]
        self.opt_add.configure(values=disponibles or ["(todos agregados)"])

    def _set_activo(self, idx):
        if idx < 0 or idx >= len(self.tramos):
            return
        self.activo = idx
        self._rebuild_chips()
        self._cargar_activo()
        self._recalcular()

    def _limpiar_activo(self):
        self.lbl_activo.configure(text="Agrega un destino para comenzar")
        self._render_panel()

    def _cargar_activo(self):
        if self.activo is None:
            self._limpiar_activo(); return
        tr = self.tramos[self.activo]
        self._cargando = True
        self.lbl_activo.configure(text="Destino: " + tr["destino"])
        temps = self._temporadas_de(tr["destino"])
        self.opt_temp.configure(values=temps)
        if tr["temporada"] not in temps:
            tr["temporada"] = temps[0]
        self.var_temp.set(tr["temporada"])
        self.st_noches.set(tr["noches"])
        _, nombres = self._hoteles_de(tr["destino"], tr["temporada"])
        tr["hoteles"] = [n for n in tr["hoteles"] if n in nombres]
        self._cargando = False
        self._render_panel()

    # handlers de widgets del destino activo
    def _on_temp(self, *_):
        if self._cargando or self.activo is None:
            return
        tr = self.tramos[self.activo]; tr["temporada"] = self.var_temp.get()
        tr["hoteles"] = []   # los hoteles cambian por temporada
        self._render_panel(); self._recalcular()

    def _on_noches(self, *_):
        if self._cargando or self.activo is None:
            return
        self.tramos[self.activo]["noches"] = self.st_noches.get()
        if self.tab == "hotel":
            self._fill_hoteles(self.tramos[self.activo])   # precios p/p dependen de noches
        self._recalcular()

    def _servicios_por_tipo(self, destino, tipo):
        dd = self.precios.get(destino, {})
        if not dd.get("terrestres"):
            return []
        servs = dd["terrestres"]["servicios"]
        if tipo == "trans":
            return [s for s in servs if es_transporte(s["nombre"])]
        return [s for s in servs if not es_transporte(s["nombre"])]

    # ------------------------------------------------------------- ninos / pax
    def _int(self, s, d=0):
        try:
            return int(float(str(s).replace(",", ".")))
        except Exception:
            return d

    def _on_pax(self):
        """Al cambiar pasajeros: refresca precios por persona de la lista visible."""
        if self.activo is not None and self.tab in ("trans", "act") \
                and hasattr(self, "_lcont") and self._lcont.winfo_exists():
            self._fill_lista(self.tramos[self.activo], self.tab)
        if self.activo is not None and self.tab == "hotel" \
                and hasattr(self, "_hcont") and self._hcont.winfo_exists():
            self._fill_hoteles(self.tramos[self.activo])
        if hasattr(self, "lbl_hab"):
            self._val_hab()
        self._recalcular()

    def _on_ninos_count(self):
        self._rebuild_edades()

    def _rebuild_edades(self):
        prev = [v.get() for v in self.edad_vars]
        for w in self.frame_edades.winfo_children():
            w.destroy()
        self.edad_vars = []
        n = self.st_ninos.get()
        if n == 0:
            ctk.CTkLabel(self.frame_edades, text="  Sin ninos", text_color=MUTED,
                         font=("Segoe UI", 11)).pack(side="left", padx=8, pady=10)
        for i in range(n):
            v = tk.StringVar(value=prev[i] if i < len(prev) else "3 anos")
            self.edad_vars.append(v)
            cont = ctk.CTkFrame(self.frame_edades, fg_color="transparent")
            cont.pack(side="left", padx=4, pady=2)
            ctk.CTkLabel(cont, text=f"Nino {i+1}", text_color=MUTED,
                         font=("Segoe UI", 9)).pack(anchor="w")
            ctk.CTkOptionMenu(cont, variable=v, values=EDAD_OPCIONES, width=100, height=26,
                              corner_radius=8, fg_color=NAVY, button_color=NAVY2,
                              button_hover_color=BLUE, dropdown_fg_color=CARD,
                              dropdown_text_color=TEXT, font=("Segoe UI", 11),
                              command=lambda e: self._on_pax()).pack()
        self._on_pax()

    def _ninos_ages(self):
        ages = []
        for v in self.edad_vars:
            try:
                ages.append(EDAD_OPCIONES.index(v.get()))
            except Exception:
                ages.append(3)
        return ages

    def _pax_total(self):
        return self.st_ad.get() + self.st_ninos.get()

    def _hoteles_sel(self, tr):
        """Lista [(nombre_display, hotel_dict)] de los hoteles elegidos como opciones."""
        hoteles, nombres = self._hoteles_de(tr["destino"], tr["temporada"])
        d = {n: h for h, n in zip(hoteles, nombres)}
        return [(n, d[n]) for n in tr["hoteles"] if n in d]

    def _calcular_tramo(self, tr, tasa):
        """Devuelve un bloque: servicios base + opciones de hotel (cada una con TOTAL)."""
        dd = self.precios.get(tr["destino"], {})
        adultos = max(self.st_ad.get(), 1)
        ages = self._ninos_ages()
        det_pax = f"{adultos} ad" + (f" + {len(ages)} nino(s)" if ages else "")
        mh, mt = self._margenes(dd)
        noches = max(int(tr.get("noches", 1)), 1)
        # --- base: transporte + actividades (no depende del hotel) ---
        base_sec = []; base = 0.0; adult_pp = 0.0
        ninos_serv = [0.0] * len(ages)
        if dd.get("terrestres"):
            servs = {s["nombre"]: s for s in dd["terrestres"]["servicios"]}
            N = max(adultos + sum(1 for a in ages if a >= 1), 1)
            for tipo, titulo in (("trans", "TRANSPORTE TERRESTRE"), ("act", "ACTIVIDADES")):
                filas = []; ss = 0.0
                for nombre in sorted(tr[tipo]):
                    serv = servs.get(nombre)
                    if not serv:
                        continue
                    priv = es_privado(nombre)
                    total_serv = precio_servicio_grupo(serv, adultos, ages, tasa, mt, priv)
                    ss += total_serv
                    pp = precio_terrestre_usd(serv, N, tasa, mt) / N
                    adult_pp += pp
                    for i, a in enumerate(ages):
                        if a == 0:
                            continue
                        elif a <= 2:
                            ninos_serv[i] += (CHILD_PRIVADO_USD if priv else pp)
                        else:
                            ninos_serv[i] += 0.5 * pp
                    desc = ""
                    if tipo == "act":
                        reg = buscar_descripcion(nombre, tr["destino"], self.descripciones)
                        if reg:
                            desc = texto_descripcion(reg)
                    # fila: (concepto, detalle_pax, total, por_pasajero_adulto, descripcion)
                    filas.append((nombre, det_pax, total_serv, pp, desc))
                if filas:
                    base_sec.append((titulo, filas, ss)); base += ss
        # --- ninos: precio fijo (servicios + tarifa hotel 3-9); no varia por hotel ---
        surch = precio_hotel_nino_noche(tasa, mh) * noches
        ninos_tot = [ninos_serv[i] + (surch if a >= 3 else 0.0) for i, a in enumerate(ages)]
        ninos_total = sum(ninos_tot)
        grupos = {}
        for a, pr in zip(ages, ninos_tot):
            g = grupos.setdefault(a, [0, 0.0]); g[0] += 1; g[1] = pr
        ninos_fijo = [(a, c, pr) for a, (c, pr) in sorted(grupos.items())]
        # --- opciones de hotel: precio POR PERSONA en sencilla / doble / triple ---
        opciones = []
        for nom, h in self._hoteles_sel(tr):
            fila = {"nombre": h["nombre"], "categoria": h.get("categoria", "")}
            for k in ("sencilla", "doble", "triple"):
                v = h.get(k)
                fila[k] = (adult_pp + precio_hotel_usd_noche(v, tasa, mh) * noches / self.OCC[k]) \
                    if v else None
            opciones.append(fila)
        return {"destino": tr["destino"],
                "subtitulo": f"{tr['destino']}   ·   Temporada {tr['temporada']}"
                             f"   ·   {noches} noches",
                "base_secciones": base_sec, "base": base,
                "opciones": opciones, "ninos": ninos_fijo, "n_adultos": adultos,
                "solo_servicios": adult_pp}

    def _calcular_todo(self):
        tasa = self._tasa()
        if tasa is None:
            return [], 0.0, False
        bloques = []; ref = 0.0; opc_mode = False
        for tr in self.tramos:
            b = self._calcular_tramo(tr, tasa)
            if len(b["opciones"]) > 1:
                opc_mode = True
            if b["opciones"]:
                o = b["opciones"][0]
                ref += o.get("doble") or o.get("sencilla") or o.get("triple") or 0.0
            else:
                ref += b["base"]
            bloques.append(b)
        return bloques, ref, opc_mode

    # ------------------------------------------------------------- habitaciones
    def _hab(self):
        return {"sencilla": self.st_hab_s.get(),
                "doble": self.st_hab_d.get(),
                "triple": self.st_hab_t.get()}

    def _hab_ocupacion(self, hab=None):
        h = hab or self._hab()
        return h["sencilla"] * 1 + h["doble"] * 2 + h["triple"] * 3

    def _sugerir_hab(self):
        a = max(self.st_ad.get(), 0)
        self.st_hab_t.set(0)
        self.st_hab_d.set(a // 2)
        self.st_hab_s.set(a % 2)
        self._on_hab_change()

    def _val_hab(self):
        try:
            ocup = self._hab_ocupacion(); ad = self.st_ad.get()
            if ocup == ad:
                self.lbl_hab.configure(text=f"OK: {ocup} plazas = {ad} adultos",
                                       text_color=GREEN)
            else:
                self.lbl_hab.configure(
                    text=f"Ojo: {ocup} plazas vs {ad} adultos", text_color=RED)
        except Exception:
            pass

    def _on_hab_change(self):
        self._val_hab()
        self._recalcular()

    def _total_reserva(self, bloques):
        """Precio TOTAL de la reserva usando la 1a opcion de hotel y las
           habitaciones indicadas (sencilla/doble/triple)."""
        hab = self._hab()
        con_op = [b for b in bloques if b["opciones"]]
        if not con_op or self._hab_ocupacion(hab) == 0:
            return None
        total = 0.0
        for b in con_op:
            o = b["opciones"][0]
            for acc, occ in (("sencilla", 1), ("doble", 2), ("triple", 3)):
                pp = o.get(acc)
                if pp and hab[acc]:
                    total += hab[acc] * occ * pp
            total += sum(c * pr for a, c, pr in b["ninos"])
        return total

    # ------------------------------------------------------------- itinerario
    def _itinerario_auto(self):
        """Arma un borrador de itinerario dia por dia a partir de los destinos y
           actividades elegidas (Dia 1 llegada, actividades, ultimo dia salida)."""
        lineas = []
        dia = 1
        if self.tramos:
            d0 = self.tramos[0]["destino"]
            lineas.append(f"DIA {dia:02d}: LLEGADA A {d0.upper()}")
            lineas.append("Recepcion en el aeropuerto por nuestro equipo y traslado al "
                          "hotel seleccionado. Registro y alojamiento.")
            dia += 1
        for tr in self.tramos:
            for act in sorted(tr["act"]):
                reg = buscar_descripcion(act, tr["destino"], self.descripciones)
                d = texto_descripcion(reg) if reg else ""
                lineas.append(f"DIA {dia:02d}: {tr['destino'].upper()} - {act.upper()}")
                lineas.append(d or "Actividad programada. Alojamiento.")
                dia += 1
        lineas.append(f"DIA {dia:02d}: TRASLADO AL AEROPUERTO")
        lineas.append("Desayuno. A la hora indicada realizamos el traslado al aeropuerto. "
                      "Fin de nuestros servicios.")
        return "\n".join(lineas)

    def _abrir_itinerario(self):
        if not self.tramos:
            messagebox.showinfo("Itinerario", "Agrega al menos un destino primero.")
            return
        win = ctk.CTkToplevel(self)
        win.title("Itinerario dia por dia")
        win.geometry("760x620"); win.configure(fg_color=BG)
        win.transient(self); win.grab_set()
        ctk.CTkLabel(win, text="Itinerario dia por dia (editable)", text_color=NAVY,
                     font=("Segoe UI", 16, "bold")).pack(anchor="w", padx=16, pady=(14, 0))
        ctk.CTkLabel(win, text="Se incluira en el PDF. Las lineas que empiezan por 'DIA' "
                     "salen resaltadas.", text_color=MUTED,
                     font=("Segoe UI", 11)).pack(anchor="w", padx=16, pady=(0, 6))
        cont = ctk.CTkFrame(win, fg_color=CARD, corner_radius=10)
        cont.pack(fill="both", expand=True, padx=16, pady=6)
        txt = tk.Text(cont, wrap="word", font=("Segoe UI", 11), bd=0, relief="flat",
                      padx=12, pady=12, background="#FFFFFF", foreground=TEXT)
        txt.pack(fill="both", expand=True, padx=4, pady=4)
        txt.insert("1.0", getattr(self, "_itinerario", "") or self._itinerario_auto())
        bar = ctk.CTkFrame(win, fg_color="transparent"); bar.pack(fill="x", padx=16, pady=(0, 14))
        def regen():
            txt.delete("1.0", "end"); txt.insert("1.0", self._itinerario_auto())
        def guardar():
            self._itinerario = txt.get("1.0", "end").strip(); win.destroy()
        ctk.CTkButton(bar, text="Regenerar automatico", width=180, fg_color=CARD2,
                      text_color=NAVY, hover_color=LINE, border_width=1, border_color=LINE,
                      font=("Segoe UI", 12, "bold"), command=regen).pack(side="left")
        ctk.CTkButton(bar, text="Guardar", width=140, fg_color=GREEN, hover_color=GREEN_H,
                      font=("Segoe UI", 13, "bold"), command=guardar).pack(side="right")

    def _recalcular(self, *a):
        try:
            bloques, total, opc_mode = self._calcular_todo()
        except Exception:
            return
        if self._tasa() is None:
            self.lbl_total.configure(text="USD --")
            self.lbl_desglose.configure(text="Esperando tasa del dia...")
            return
        reserva = self._total_reserva(bloques)
        if reserva is not None:
            self.lbl_total.configure(text=usd(reserva))
            h = self._hab()
            det = ", ".join(f"{h[k]} {k}" for k in ("sencilla", "doble", "triple") if h[k])
            self.lbl_desglose.configure(text=f"Total reserva ({det}) - 1a opcion")
        else:
            self.lbl_total.configure(text=usd(total))
            self.lbl_desglose.configure(text="Indica las habitaciones para el total")

    # ------------------------------------------------------------- acciones
    def _on_fecha_desde(self):
        d = self.sel_desde.get()
        if d:
            self.sel_hasta.minimo = d
            if self.sel_hasta.get() and self.sel_hasta.get() < d:
                self.sel_hasta.clear()
        # el descuento y los margenes dependen de la fecha de viaje -> refrescar
        if self.activo is not None:
            self._render_panel()
        self._recalcular()

    def _abrir_empresa(self):
        VentanaEmpresa(self, self.cfg, self._on_cfg)

    def _on_cfg(self, cfg):
        self.cfg = cfg

    def _abrir_cotizaciones(self):
        VentanaCotizaciones(self)

    def _chequear_seguimientos(self):
        """Al abrir: importa las del HTML (clientes) y alerta de seguimientos vencidos."""
        def worker():
            try:
                importar_cotizaciones_html()
            except Exception:
                pass
            try:
                pend = seguimientos_pendientes()
            except Exception:
                pend = []
            if pend:
                self.after(0, lambda: self._alerta_seguimientos(pend))
        threading.Thread(target=worker, daemon=True).start()

    def _alerta_seguimientos(self, pend):
        lineas = [f"• {it.get('numero','')}  {it.get('cliente','') or '(sin agencia)'}"
                  f"  ->  {motivo}" for it, motivo in pend[:15]]
        extra = f"\n\n(y {len(pend) - 15} mas)" if len(pend) > 15 else ""
        if messagebox.askyesno(
                "Seguimiento de cotizaciones",
                f"Tienes {len(pend)} cotizacion(es) para dar seguimiento hoy:\n\n"
                + "\n".join(lineas) + extra + "\n\n¿Abrir el historial de cotizaciones?"):
            self._abrir_cotizaciones()

    def _abrir_clientes(self):
        VentanaClientes(self, on_cambio=self._on_clientes_cambio)

    def _on_clientes_cambio(self):
        self.clientes = cargar_clientes()

    def _refrescar_clientes_picker(self):
        pass  # el buscador (lupa) lee self.clientes directamente

    def _buscar_cliente(self):
        BuscadorClientes(self, self.clientes, self._usar_cliente,
                         on_editar=self._editar_cliente_por_nombre,
                         on_eliminar=self._eliminar_cliente_por_nombre)

    def _eliminar_cliente_por_nombre(self, nombre):
        # filtra en el mismo objeto lista para que el buscador vea el cambio
        self.clientes[:] = [c for c in self.clientes if c.get("empresa", "") != nombre]
        guardar_clientes(self.clientes)
        if self.var_cli.get().strip() == nombre:
            self.var_cli.set(""); self.var_email.set("")
            self.var_asesor.set(""); self.var_aso_tel.set("")
            self._vendedores = []
            self.opt_asesor.configure(values=["(vendedor)"]); self.opt_asesor.set("(vendedor)")

    def _editar_cliente_actual(self):
        self._editar_cliente_por_nombre(self.var_cli.get().strip())

    def _editar_cliente_por_nombre(self, nombre):
        VentanaClientes(self, on_cambio=self._on_clientes_cambio,
                        preseleccion=nombre or None)

    def _usar_cliente(self, nombre):
        c = next((x for x in self.clientes if x.get("empresa") == nombre), None)
        if not c:
            return
        self.var_cli.set(c.get("empresa", ""))
        email = c.get("email", "")
        if not email and c.get("vendedores"):
            email = c["vendedores"][0].get("email", "")
        if email:
            self.var_email.set(email)
        # llenar el selector de asesores con los vendedores de la empresa
        self._vendedores = c.get("vendedores", []) or []
        vals = [v.get("nombre", "") for v in self._vendedores if v.get("nombre")]
        self.opt_asesor.configure(values=vals or ["(vendedor)"])
        if vals:
            self.opt_asesor.set(vals[0])
            self._usar_asesor(vals[0])
        else:
            self.opt_asesor.set("(vendedor)")
            self.var_asesor.set(""); self.var_aso_tel.set("")

    def _usar_asesor(self, nombre):
        v = next((x for x in getattr(self, "_vendedores", []) if x.get("nombre") == nombre), None)
        self.var_asesor.set(nombre if nombre != "(vendedor)" else "")
        if v:
            self.var_aso_tel.set(v.get("telefono", ""))
            if v.get("email"):
                self.var_email.set(v.get("email", ""))

    def _nueva(self):
        hoy = datetime.date.today()
        self.var_cli.set(""); self.var_email.set("")
        self.var_asesor.set(""); self.var_aso_tel.set("")
        self._vendedores = []
        self.opt_asesor.configure(values=["(vendedor)"]); self.opt_asesor.set("(vendedor)")
        self.sel_desde.clear(); self.sel_hasta.clear()
        self.sel_hasta.minimo = hoy
        self._numero = "COT-" + hoy.strftime("%Y%m%d") + "-001"
        self._fecha = hoy.strftime("%d/%m/%Y")
        self._valida = add_months(hoy, 1).strftime("%d/%m/%Y")
        self.tramos = []; self.activo = None
        self._itinerario = ""
        self._rebuild_chips(); self._limpiar_activo()
        # agregar un primer destino por comodidad
        self._add_destino(list(self.precios.keys())[0])
        if hasattr(self, "st_hab_s"):
            self._sugerir_hab()
        self._recalcular()

    def _generar(self):
        if self._tasa() is None:
            messagebox.showwarning("Falta la tasa",
                                   "Aun no hay tasa del dia. Pulsa 'Actualizar'.")
            return
        if not self.tramos:
            messagebox.showwarning("Sin destinos", "Agrega al menos un destino."); return
        bloques, total, opc_mode = self._calcular_todo()
        if total <= 0 or not any(b["base_secciones"] or b["opciones"] for b in bloques):
            messagebox.showwarning("Cotizacion vacia",
                                   "Selecciona al menos un hotel, transporte o actividad."); return
        if not self.sel_desde.get() or not self.sel_hasta.get():
            messagebox.showwarning("Falta la fecha del viaje",
                                   "Elige la FECHA DEL VIAJE (ida y regreso) en el calendario."); return
        if self.sel_hasta.get() < self.sel_desde.get():
            messagebox.showwarning("Fechas invalidas",
                                   "La fecha de regreso no puede ser anterior a la de ida."); return
        # las noches del itinerario deben cuadrar con las fechas del viaje
        total_noches = sum(max(int(t.get("noches", 1)), 1) for t in self.tramos)
        dias = (self.sel_hasta.get() - self.sel_desde.get()).days
        if dias != total_noches:
            if not messagebox.askyesno(
                    "Noches vs fechas del viaje",
                    f"Las noches del itinerario suman {total_noches}, pero entre las fechas "
                    f"elegidas hay {dias} noche(s).\n\nNo coinciden. "
                    "¿Deseas continuar de todos modos?"):
                return
        if not self.var_email.get().strip():
            messagebox.showwarning("Falta el email del cliente",
                                   "Debes ingresar el EMAIL del cliente antes de generar."); return
        ad = self.st_ad.get(); ages = self._ninos_ages()
        pax_txt = f"{ad} adultos"
        if ages:
            det = ", ".join(("bebe" if a == 0 else f"{a} anos") for a in ages)
            pax_txt += f", {len(ages)} ninos ({det})"
        cliente = self.var_cli.get().strip(); email = self.var_email.get().strip()
        firma_nom, firma_cargo = self._cotz_map.get(self.var_cotizador.get(),
                                                    (COTIZADORES[0][0], COTIZADORES[0][1]))
        # consecutivo de la cotizacion (se guarda en el historial al final)
        self._numero = peek_numero_cotizacion()
        reserva = self._total_reserva(bloques)
        total_mostrar = reserva if reserva is not None else total
        datos = {"numero": self._numero, "fecha": self._fecha, "valida_hasta": self._valida,
                 "cliente": cliente, "cli_email": email,
                 "asesor": self.var_asesor.get().strip(),
                 "asesor_tel": self.var_aso_tel.get().strip(),
                 "fechas_viaje": f"{self.sel_desde.get_str()} al {self.sel_hasta.get_str()}",
                 "pax_txt": pax_txt, "opc_mode": opc_mode,
                 "habitaciones": self._hab(),
                 "itinerario": (self._itinerario or self._itinerario_auto()),
                 "firma_nombre": firma_nom, "firma_cargo": firma_cargo,
                 "notas": (f"Vigencia: esta cotizacion tiene una validez de un (1) mes, "
                           f"hasta el {self._valida}. " + self.cfg.get("notas", ""))}
        slug = "".join(c for c in cliente if c.isalnum() or c in " _-").strip().replace(" ", "_")
        destinos_txt = "-".join(t["destino"][:3] for t in self.tramos)
        ruta = filedialog.asksaveasfilename(
            title="Guardar cotizacion PDF", defaultextension=".pdf",
            initialfile=f"Cotizacion_{destinos_txt}_{slug or 'cliente'}.pdf",
            filetypes=[("PDF", "*.pdf")])
        if not ruta:
            return
        try:
            generar_pdf(self.cfg, datos, bloques, total, ruta)
        except Exception as e:
            messagebox.showerror("Error al generar PDF", str(e)); return
        # guardar en el historial de cotizaciones (con consecutivo)
        try:
            numero = registrar_cotizacion({
                "cliente": cliente, "asesor": self.var_asesor.get().strip(),
                "asesor_tel": self.var_aso_tel.get().strip(),
                "fecha": self._fecha, "fechas_viaje": datos["fechas_viaje"],
                "cotizado_por": firma_nom, "email": email,
                "destinos": [t["destino"] for t in self.tramos],
                "total": total_mostrar, "pdf": ruta})
            self.lbl_desglose.configure(text=f"Guardada {numero}")
        except Exception:
            pass
        # enviar por correo
        self._enviar_pdf(ruta, email, cliente, datos.get("asesor", ""))

    def _enviar_pdf(self, ruta, email, cliente, asesor=""):
        smtp_ok = self.cfg.get("correo_remitente") and self.cfg.get("smtp_password")
        if not email:
            if messagebox.askyesno("Sin email",
                                   "El cliente no tiene email, no se puede enviar.\n\n"
                                   "PDF guardado. ¿Abrirlo ahora?"):
                self._abrir(ruta)
            return
        if not smtp_ok:
            messagebox.showinfo("Correo no configurado",
                                "PDF guardado.\n\nPara enviarlo automaticamente, configura el "
                                "'Correo remitente' y su contrasena en 'Datos de mi empresa'.")
            if messagebox.askyesno("PDF", "¿Abrir el PDF ahora?"):
                self._abrir(ruta)
            return
        f_nom, f_cargo = self._cotz_map.get(self.var_cotizador.get(),
                                            (COTIZADORES[0][0], COTIZADORES[0][1]))
        asunto = "Cotizacion de viaje - INNOBA Colombia DMC"
        if cliente:
            asunto += f" - {cliente}"
        saludo = asesor.strip() or cliente or "cliente"
        ref = f"para {cliente}" if (cliente and asesor.strip()) else "para su viaje"
        cuerpo = (f"Estimado(a) {saludo},\n\n"
                  f"Adjunto encontrara la cotizacion solicitada {ref}.\n"
                  "Quedamos atentos a cualquier inquietud.\n\n"
                  "Cordialmente,\n"
                  f"{f_nom}\n{f_cargo}")
        dlg = messagebox.showinfo("Enviando", "Enviando la cotizacion a " + email + "...")
        def worker():
            try:
                enviar_correo(self.cfg, email, asunto, cuerpo, ruta)
                self.after(0, lambda: self._envio_ok(ruta, email))
            except Exception as e:
                self.after(0, lambda: self._envio_error(ruta, str(e)))
        threading.Thread(target=worker, daemon=True).start()

    def _envio_ok(self, ruta, email):
        if messagebox.askyesno("Enviado",
                               "Cotizacion enviada correctamente a:\n" + email +
                               "\n\n¿Abrir el PDF?"):
            self._abrir(ruta)

    def _envio_error(self, ruta, err):
        messagebox.showerror("Error al enviar",
                             "El PDF se guardo, pero no se pudo enviar el correo:\n\n" + err +
                             "\n\nRevisa el correo remitente y la contrasena en "
                             "'Datos de mi empresa'.")
        if messagebox.askyesno("PDF", "¿Abrir el PDF ahora?"):
            self._abrir(ruta)

    def _abrir(self, ruta):
        try: os.startfile(ruta)
        except Exception: pass


class VentanaEmpresa(ctk.CTkToplevel):
    def __init__(self, master, cfg, callback):
        super().__init__(master)
        self.cfg = dict(cfg); self.callback = callback
        self.title("Datos de mi empresa"); self.geometry("580x720"); self.configure(fg_color=BG)
        self.transient(master); self.grab_set()
        cont = ctk.CTkScrollableFrame(self, fg_color=BG); cont.pack(fill="both", expand=True,
                                                                    padx=16, pady=16)
        ctk.CTkLabel(cont, text="Datos de mi empresa", text_color=NAVY,
                     font=("Segoe UI", 18, "bold")).pack(anchor="w", pady=(0, 10))
        self.entradas = {}
        def campo(clave, etq, secreto=False):
            ctk.CTkLabel(cont, text=etq, text_color=MUTED, font=("Segoe UI", 11)).pack(
                anchor="w", padx=2)
            v = tk.StringVar(value=str(self.cfg.get(clave, "")))
            self.entradas[clave] = v
            e = ctk.CTkEntry(cont, textvariable=v, height=34, corner_radius=8, border_color=LINE,
                             show="*" if secreto else "")
            e.pack(fill="x", pady=(0, 8))
        for clave, etq in [("empresa", "Nombre de la empresa"), ("nit", "NIT / RUC"),
                           ("direccion", "Direccion"), ("telefono", "Telefono"),
                           ("email", "Email"), ("web", "Sitio web"),
                           ("firma_nombre", "Firma - Nombre"), ("firma_cargo", "Firma - Cargo")]:
            campo(clave, etq)
        ctk.CTkLabel(cont, text="—  Envio de correo (Office 365)  —", text_color=NAVY,
                     font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(6, 4))
        campo("correo_remitente", "Correo remitente (desde donde se envia)")
        campo("smtp_password", "Contrasena del correo (o contrasena de aplicacion)", secreto=True)
        campo("smtp_servidor", "Servidor SMTP")
        campo("smtp_puerto", "Puerto SMTP")
        ctk.CTkLabel(cont, text="Logo (PNG/JPG)", text_color=MUTED,
                     font=("Segoe UI", 11)).pack(anchor="w", padx=2)
        lf = ctk.CTkFrame(cont, fg_color="transparent"); lf.pack(fill="x", pady=(0, 8))
        self.logo_var = tk.StringVar(value=self.cfg.get("logo", ""))
        ctk.CTkLabel(lf, textvariable=self.logo_var, text_color=TEXT, font=("Segoe UI", 10),
                     anchor="w").pack(side="left", fill="x", expand=True)
        ctk.CTkButton(lf, text="Elegir", width=70, fg_color=NAVY, hover_color=BLUE,
                      command=self._logo).pack(side="left", padx=4)
        ctk.CTkButton(lf, text="Quitar", width=70, fg_color=CARD2, text_color=NAVY,
                      hover_color=LINE, command=lambda: self.logo_var.set("")).pack(side="left")
        ctk.CTkLabel(cont, text="Notas por defecto", text_color=MUTED,
                     font=("Segoe UI", 11)).pack(anchor="w", padx=2)
        self.txt = ctk.CTkTextbox(cont, height=80, corner_radius=8, border_width=1,
                                  border_color=LINE, fg_color=CARD)
        self.txt.insert("1.0", self.cfg.get("notas", "")); self.txt.pack(fill="x", pady=(0, 12))
        bts = ctk.CTkFrame(cont, fg_color="transparent"); bts.pack(fill="x")
        ctk.CTkButton(bts, text="Guardar", fg_color=GREEN, hover_color=GREEN_H,
                      font=("Segoe UI", 13, "bold"), command=self._guardar).pack(side="right")
        ctk.CTkButton(bts, text="Cancelar", fg_color=CARD2, text_color=NAVY, hover_color=LINE,
                      command=self.destroy).pack(side="right", padx=8)

    def _logo(self):
        r = filedialog.askopenfilename(title="Elegir logo",
                                       filetypes=[("Imagenes", "*.png *.jpg *.jpeg *.gif *.bmp")])
        if r: self.logo_var.set(r)

    def _guardar(self):
        for clave, v in self.entradas.items():
            self.cfg[clave] = v.get().strip()
        self.cfg["logo"] = self.logo_var.get().strip()
        self.cfg["notas"] = self.txt.get("1.0", "end").strip()
        try:
            guardar_config(self.cfg)
        except Exception as e:
            messagebox.showerror("Error", str(e)); return
        self.callback(self.cfg)
        messagebox.showinfo("Guardado", "Datos guardados.")
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
