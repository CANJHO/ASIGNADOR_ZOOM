# app.py
import io
from datetime import datetime, timedelta, time as dtime
import pandas as pd
import streamlit as st
import re

# ==============================
# MAPEOS Y CONSTANTES
# ==============================

# Día
DIA_NUM_TO_NAME_ACC = {
    1: "LUNES", 2: "MARTES", 3: "MIÉRCOLES", 4: "JUEVES",
    5: "VIERNES", 6: "SÁBADO", 7: "DOMINGO"
}
DIA_NUM_TO_NAME_PLAIN = {
    1: "LUNES", 2: "MARTES", 3: "MIERCOLES", 4: "JUEVES",
    5: "VIERNES", 6: "SABADO", 7: "DOMINGO"
}
DIA_NAME_PLAIN_TO_NUM = {
    "LUNES":1,"MARTES":2,"MIERCOLES":3,"JUEVES":4,"VIERNES":5,"SABADO":6,"DOMINGO":7
}
DIA_NUM_TO_ABBR = {1:"LU",2:"MA",3:"MI",4:"JU",5:"VI",6:"SA",7:"DO"}

# Facultad
FACULTAD_MAP = {
    "I": "INGENIERÍA, CIENCIAS Y ADMINISTRACIÓN",
    "S": "CIENCIAS DE LA SALUD",
}

# Escuela (código <-> nombre completo)
ESCUELA_CODE_TO_FULL = {
    "AF":"ADMINISTRACIÓN Y FINANZAS",
    "IS":"INGENIERÍA DE SISTEMAS",
    "II":"INGENIERÍA EN INDUSTRIAS ALIMENTARIAS",
    "IN":"INGENIERÍA INDUSTRIAL",
    "IC":"INGENIERÍA CIVIL",
    "DE":"DERECHO",
    "CA":"CONTABILIDAD",
    "EN":"ENFERMERÍA",
    "PS":"PSICOLOGÍA",
    "MH":"MEDICINA HUMANA",
    "OB":"OBSTETRICIA",
    "AR":"ARQUITECTURA",
    "T1":"TECNOLOGÍA MÉDICA - ESPECIALIDAD EN LABORATORIO CLINICO Y ANATOMÍA PATOLÓGICA",
    "T3":"TECNOLOGÍA MÉDICA - ESPECIALIDAD EN TERAPIA FÍSICA Y REHABILITACIÓN",
    "T2":"TECNOLOGÍA MÉDICA - ESPECIALIDAD EN TERAPIA DE LENGUAJE",
    "T4":"TECNOLOGÍA MÉDICA - OPTOMETRÍA",
}
ESCUELA_FULL_TO_CODE = {v.upper():k for k,v in ESCUELA_CODE_TO_FULL.items()}
ESCUELA_FULL_SET = set(ESCUELA_FULL_TO_CODE.keys())

# Modalidad -> V/P
MODALIDAD_MAP = {"V":"V","VIRTUAL":"V","P":"P","PRESENCIAL":"P"}

# Local base -> código y código -> texto estándar (para tema)
LOCAL_TO_CODE_BASE = {
    "FILIAL":"IC","ICA":"IC","IC":"IC",
    "PRINCIPAL":"CH","CHINCHA":"CH","CH":"CH","CHP":"CH",
    "SUNAMPE":"SU","SUMANPE":"SU","SU":"SU",
    "HUARUA":"HU","HU":"HU"
}
LOCAL_CODE_TO_TEXT_STD = {"IC":"FILIAL","CH":"PRINCIPAL","SU":"SUNAMPE","HU":"HUARUA"}

REQUERIDAS = ["DOCENTE","DIA","HORA INICIO","HORA FIN"]

# ==============================
# UTILIDADES
# ==============================

def norm_txt(x):
    if pd.isna(x): return ""
    return str(x).strip()

def norm_upper(x): return norm_txt(x).upper()

def convertir_hora(hora_raw):
    if pd.isna(hora_raw):
        raise ValueError("Hora vacía")
    if isinstance(hora_raw,(datetime,pd.Timestamp)):
        return datetime.combine(datetime.today().date(), hora_raw.time())
    if isinstance(hora_raw,dtime):
        return datetime.combine(datetime.today().date(), hora_raw)
    s = str(hora_raw).strip().upper().replace(".",":").replace("H",":")
    if s.isdigit():
        if len(s) in (1,2):  return datetime.strptime(f"{int(s):02d}:00","%H:%M")
        if len(s)==3:        s="0"+s
        if len(s)==4:        return datetime.strptime(f"{s[:2]}:{s[2:]}","%H:%M")
    for fmt in ("%H:%M","%H:%M:%S","%I:%M %p","%I:%M:%S %p"):
        try: return datetime.strptime(s,fmt)
        except ValueError: pass
    try:
        f=float(s)
        if 0<=f<1:
            total=int(round(f*24*60)); h=total//60; m=total%60
            return datetime.strptime(f"{h:02d}:{m:02d}","%H:%M")
    except: pass
    raise ValueError(f"Formato de hora no reconocido: {hora_raw}")

def hora_hhmm(v): return convertir_hora(v).strftime("%H:%M")

def parse_dia_to_num(d):
    if pd.isna(d): return None
    s=norm_upper(d)
    try:
        n=int(float(s))
        if 1<=n<=7: return n
    except: pass
    abbr_to_num={v:k for k,v in DIA_NUM_TO_ABBR.items()}
    if s in abbr_to_num: return abbr_to_num[s]
    s2 = s.replace("Á","A").replace("É","E").replace("Í","I").replace("Ó","O").replace("Ú","U")
    return DIA_NAME_PLAIN_TO_NUM.get(s2)

def dia_num_to_abbr(n):       return DIA_NUM_TO_ABBR.get(n,"")
def dia_num_to_full_acc(n):   return DIA_NUM_TO_NAME_ACC.get(n,"")
def dia_num_to_full_plain(n): return DIA_NUM_TO_NAME_PLAIN.get(n,"")

def normalizar_facultad(v):
    s=norm_upper(v)
    return FACULTAD_MAP.get(s, norm_txt(v))

def escuela_to_full(v):
    s=norm_upper(v)
    if s in ESCUELA_CODE_TO_FULL: return ESCUELA_CODE_TO_FULL[s]
    if s in ESCUELA_FULL_SET:     return norm_txt(v)
    return norm_txt(v)

def escuela_to_code(v):
    s=norm_upper(v)
    if s in ESCUELA_CODE_TO_FULL: return s
    if s in ESCUELA_FULL_SET:     return ESCUELA_FULL_TO_CODE[s]
    m=re.match(r"[A-Z]{2,3}", s)
    return m.group(0) if m else s

def normalizar_modalidad(v):
    s=norm_upper(v)
    return MODALIDAD_MAP.get(s, s if s in ("V","P") else s)

def parse_custom_local_map(text):
    out={}
    for line in text.splitlines():
        if "=" in line:
            k,v=line.split("=",1)
            out[k.strip().upper()]=v.strip().upper()
    return out

def local_to_code(v, custom_map):
    s=norm_upper(v)
    if s in custom_map: return custom_map[s]
    if s in LOCAL_TO_CODE_BASE: return LOCAL_TO_CODE_BASE[s]
    if 2<=len(s)<=3 and s.isalpha(): return s
    parts=[p for p in s.split() if p]
    if parts:
        cand=("".join(p[0] for p in parts)[:3]).upper()
        return cand if cand else s
    return s

def local_code_to_text(v):
    s=norm_upper(v)
    return LOCAL_CODE_TO_TEXT_STD.get(s, s)

def seccion_letra(seccion):
    s=norm_upper(seccion)
    if "-" in s:
        try:
            suf=s.split("-",1)[1]
            n=int("".join(ch for ch in suf if ch.isdigit()))
            if n>=1:
                idx=min(n,26)
                return chr(64+idx)  # 1->A
        except: pass
    for ch in s:
        if ch.isalpha(): return ch
    return ""

def construir_grupo(seccion, modalidad, local_code):
    letra = seccion_letra(seccion)
    mod = normalizar_modalidad(modalidad)
    mod = "V" if mod in ("V","VIRTUAL") else ("P" if mod in ("P","PRESENCIAL") else mod)
    return f"{letra}{mod} - {local_code}".strip()

def format_dni(v):
    s = norm_txt(v)
    if not s: return ""
    digits = "".join(ch for ch in s if ch.isdigit())
    if not digits: return ""
    return digits.zfill(8)

def construir_tema_zoom(row, custom_local_map):
    """
    {PLAN}-{COD_CURSO}-{CURSO}-{SECCION}-{ESCUELA_CODE}-{LOCAL_TXT}|{DNI}|{DIA_ACC} {HH:MM}-{HH:MM}-{DURACION}
    (DNI va SIEMPRE en el medio, entre dos pipes)
    """
    plan = norm_txt(row.get("PLAN","")) or norm_txt(row.get("COD_PLAN",""))
    cod_curso = norm_txt(row.get("COD_CURSO",""))
    curso = norm_txt(row.get("CURSO",""))
    seccion = norm_txt(row.get("SECCION",""))

    escuela_code = escuela_to_code(row.get("ESCUELA",""))

    local_raw = row.get("LOCAL","")
    local_code = local_to_code(local_raw, custom_local_map)
    local_txt = norm_upper(local_raw)
    if local_txt in LOCAL_TO_CODE_BASE or len(local_txt) <= 3:
        local_txt = local_code_to_text(local_code)  # IC -> FILIAL, CH -> PRINCIPAL, etc.

    dni = format_dni(row.get("DNI_DOC",""))

    n = parse_dia_to_num(row.get("DIA",""))
    dia_acc = dia_num_to_full_acc(n) if n else ""   # con tilde para público

    hi = hora_hhmm(row.get("HORA INICIO",""))
    hf = hora_hhmm(row.get("HORA FIN",""))
    dur = norm_txt(row.get("DURACION",""))

    left = "-".join([p for p in [plan, cod_curso, curso, seccion, escuela_code, local_txt] if p!=""])
    right = f"{dia_acc} {hi}-{hf}-{dur}".strip()

    return f"{left}|{dni}|{right}"

# ==============================
# ASIGNACIÓN DE ZOOM (motor)
# ==============================

def contar_conflictos(nueva_inicio, nueva_fin, asignaciones, margen_minutos):
    conflictos=0
    for a in asignaciones:
        ei=a['inicio']; ef=a['fin']
        if not (nueva_fin + timedelta(minutes=margen_minutos) <= ei or
                nueva_inicio >= ef + timedelta(minutes=margen_minutos)):
            conflictos += 1
    return conflictos

def generar_correo_zoom(numero, prefijo):
    return f"{prefijo}{str(numero).zfill(4)}@autonomadeica.edu.pe"

def asignar_zoom(df, margen_minutos, max_simult, prefijo_correo):
    df=df.copy()
    df['Zoom asignado']=None
    zoom_usos={}; zoom_counter=1
    for idx,fila in df.iterrows():
        n=parse_dia_to_num(fila['DIA'])
        dia_key=dia_num_to_full_acc(n) if n else ""
        inicio=convertir_hora(fila['HORA INICIO'])
        fin=convertir_hora(fila['HORA FIN'])
        asignado=False
        for zoom_id,usos in zoom_usos.items():
            conflictos=contar_conflictos(inicio,fin,usos.get(dia_key,[]),margen_minutos)
            if conflictos < max_simult:
                usos.setdefault(dia_key,[]).append({'inicio':inicio,'fin':fin})
                df.at[idx,'Zoom asignado']=zoom_id
                asignado=True
                break
        if not asignado:
            nuevo=generar_correo_zoom(zoom_counter,prefijo_correo)
            zoom_counter+=1
            zoom_usos[nuevo]={dia_key:[{'inicio':inicio,'fin':fin}]}
            df.at[idx,'Zoom asignado']=nuevo
    return df

def generar_resumen(df_asignado):
    t=df_asignado.copy()
    t["DIA_NUM"]=t["DIA"].apply(parse_dia_to_num)
    t["DIA_FULL"]=t["DIA_NUM"].apply(dia_num_to_full_acc)
    res=(t.groupby("DIA_FULL")
           .agg({'Zoom asignado':pd.Series.nunique})
           .rename(columns={'Zoom asignado':'N° de Licencias Zoom'})
           .reset_index()
           .sort_values(by='DIA_FULL',key=lambda s:s.map({v:k for k,v in DIA_NUM_TO_NAME_ACC.items()})))
    return res

# ==============================
# TRANSFORMACIÓN DE EXPORTACIÓN
# ==============================

def convertir_a_excel_export(df, custom_local_map):
    """
    Devuelve un DF listo para exportar:
    - DIA -> abreviatura (LU..DO)
    - FACULTAD -> nombre completo
    - ESCUELA -> nombre completo
    - MODALIDAD -> V/P
    - LOCAL -> código (IC/CH/SU/HU/…)
    - + GRUPO, RANGO HORARIO, TEMA_ZOOM, DETALLE HORARIO
    """
    out = df.copy()

    # DIA_NUM y DIA abreviado
    out["DIA_NUM"] = out["DIA"].apply(parse_dia_to_num)
    out["DIA"] = out["DIA_NUM"].apply(dia_num_to_abbr)

    # FACULTAD
    if "FACULTAD" in out.columns:
        out["FACULTAD"] = out["FACULTAD"].apply(normalizar_facultad)

    # ESCUELA (nombre completo)
    if "ESCUELA" in out.columns:
        out["ESCUELA"] = out["ESCUELA"].apply(escuela_to_full)

    # MODALIDAD -> V/P
    if "MODALIDAD" in out.columns:
        out["MODALIDAD"] = out["MODALIDAD"].apply(normalizar_modalidad).map(
            lambda x: "V" if x in ("V","VIRTUAL") else ("P" if x in ("P","PRESENCIAL") else x)
        )

    # LOCAL -> código
    if "LOCAL" in out.columns:
        out["LOCAL"] = out["LOCAL"].apply(lambda v: local_to_code(v, custom_local_map))

    # GRUPO
    out["GRUPO"] = out.apply(
        lambda r: construir_grupo(r.get("SECCION",""), r.get("MODALIDAD",""), r.get("LOCAL","")),
        axis=1
    )

    # RANGO HORARIO
    out["RANGO HORARIO"] = out.apply(
        lambda r: f"{hora_hhmm(r.get('HORA INICIO',''))}-{hora_hhmm(r.get('HORA FIN',''))}",
        axis=1
    )

    # TEMA_ZOOM (DNI al medio)
    out["TEMA_ZOOM"] = out.apply(lambda r: construir_tema_zoom(r, custom_local_map), axis=1)

    # DETALLE HORARIO (con tildes)
    out["DETALLE HORARIO"] = out.apply(
        lambda r: f"{r.get('DIA_NUM','')} - {dia_num_to_full_acc(r.get('DIA_NUM'))} de - {r.get('RANGO HORARIO','')}",
        axis=1
    )

    return out

# ==============================
# UI STREAMLIT
# ==============================

st.set_page_config(page_title="Asignador UAI", layout="wide")
st.title("📅 Asignador de Licencias Zoom - UAI")

# Sidebar: enlace + locales personalizados
st.sidebar.header("VISITA TAMBIÉN")
url_moodle="https://moodle-admision-kkvkzem6ls2m4f458ln4ut.streamlit.app/#exportador-de-admision-moodle"
st.sidebar.markdown(
    f"""
    <a href="{url_moodle}" target="_blank" style="text-decoration:none;">
      <span style="display:inline-block;padding:0.6rem 1rem;background:#d32f2f;color:#fff;border-radius:8px;font-weight:600;">
        Admisión Moodle
      </span>
    </a>
    """,
    unsafe_allow_html=True
)

st.sidebar.subheader("Locales personalizados (opcional)")
custom_local_text = st.sidebar.text_area("NOMBRE=CODIGO (una por línea)", value="", height=120)
custom_local_map = parse_custom_local_map(custom_local_text)

# Plantillas
st.subheader("📥 Descargar plantillas")

# 1) Mínima (solo obligatorias)
plantilla_min = pd.DataFrame(columns=["DOCENTE","DIA","HORA INICIO","HORA FIN"])
buf_min = io.BytesIO()
with pd.ExcelWriter(buf_min, engine='openpyxl') as w:
    plantilla_min.to_excel(w, index=False, sheet_name="Plantilla")
st.download_button("📄 Descargar plantilla mínima", data=buf_min.getvalue(),
                   file_name="plantilla_minima.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# 2) Completa (20 encabezados como en la salida)
plantilla_full_cols = [
    "DOCENTE","DIA","HORA INICIO","HORA FIN",
    "FACULTAD","ESCUELA","COD_PLAN","COD_CURSO","SECCION","CURSO",
    "MODALIDAD","LOCAL","DNI_DOC","DURACION",
    "GRUPO","TEMA_ZOOM","Zoom asignado","DIA_NUM","RANGO HORARIO","DETALLE HORARIO"
]
fila_demo = {c:"" for c in plantilla_full_cols}
for c in ["GRUPO","TEMA_ZOOM","Zoom asignado","DIA_NUM","RANGO HORARIO","DETALLE HORARIO"]:
    fila_demo[c] = "AUTOGENERADO (no editar)"
plantilla_full = pd.DataFrame([fila_demo], columns=plantilla_full_cols)

buf_full = io.BytesIO()
with pd.ExcelWriter(buf_full, engine='openpyxl') as w:
    plantilla_full.to_excel(w, index=False, sheet_name="Plantilla")
st.download_button("📄 Descargar plantilla completa (20 columnas)", data=buf_full.getvalue(),
                   file_name="plantilla_completa_20.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Subida
archivo = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx"])

# Config
st.subheader("⚙️ Configuración")
margen_minutos = st.number_input("Margen (min) entre reuniones", 0, 60, 10)
max_reuniones = st.radio("Máximo de reuniones simultáneas por licencia", [1,2], index=1)
prefijo_correo = st.text_input("Prefijo de correo para generar licencias", "UAI")

if archivo is not None and st.button("🚀 Asignar Zoom y Descargar"):
    df = pd.read_excel(archivo)
    df.columns = [c.upper().strip() for c in df.columns]

    faltantes = [c for c in REQUERIDAS if c not in df.columns]
    if faltantes:
        st.error("Faltan columnas obligatorias: " + ", ".join(faltantes))
        st.stop()

    # DIA a número (para la lógica)
    df["DIA"] = df["DIA"].apply(parse_dia_to_num)
    if df["DIA"].isna().any():
        st.error("Hay filas con DIA inválido. Acepta 1-7, LU..DO o LUNES..DOMINGO.")
        st.stop()

    # Validar horas
    try:
        _ = df["HORA INICIO"].apply(convertir_hora)
        _ = df["HORA FIN"].apply(convertir_hora)
    except Exception as e:
        st.error(f"Formato de hora no reconocido: {e}")
        st.stop()

    # Asignación
    df_asig = asignar_zoom(df, margen_minutos, max_reuniones, prefijo_correo)

    # Export transform
    df_export = convertir_a_excel_export(df_asig, custom_local_map)

    # Resumen
    df_resumen = generar_resumen(df_asig)

    # Mostrar
    st.success("✅ Procesamiento completado")
    st.subheader("👩‍🏫 Horarios (con GRUPO y TEMA_ZOOM)")
    st.dataframe(df_export, use_container_width=True)

    st.subheader("📊 Resumen por día")
    st.dataframe(df_resumen, use_container_width=True)

    # Descargar
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        df_export.to_excel(w, sheet_name="Horarios", index=False)
        df_resumen.to_excel(w, sheet_name="Resumen", index=False)
    st.download_button("💾 Descargar archivo Excel",
                       data=out.getvalue(),
                       file_name="horario_con_zoom.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
