# app.py
import io
from datetime import datetime, timedelta, time as dtime
import pandas as pd
import streamlit as st

# ==============================
# MAPEOS Y CONSTANTES
# ==============================

# Día: número -> nombre (con tildes) y abreviatura
DIA_NUM_TO_NAME_ACCENT = {
    1: "LUNES", 2: "MARTES", 3: "MIÉRCOLES", 4: "JUEVES",
    5: "VIERNES", 6: "SÁBADO", 7: "DOMINGO"
}
DIA_NAME_NONTILDE_TO_NUM = {
    "LUNES":1, "MARTES":2, "MIERCOLES":3, "JUEVES":4, "VIERNES":5, "SABADO":6, "DOMINGO":7
}
DIA_NUM_TO_ABBR = {1: "LU", 2: "MA", 3: "MI", 4: "JU", 5: "VI", 6: "SA", 7: "DO"}

# Facultad: código -> nombre completo
FACULTAD_MAP = {
    "I": "INGENIERÍA, CIENCIAS Y ADMINISTRACIÓN",
    "S": "CIENCIAS DE LA SALUD",
}

# Escuela: código -> nombre completo (exportamos nombre completo)
ESCUELA_CODE_TO_FULL = {
    "AF": "ADMINISTRACIÓN Y FINANZAS",
    "IS": "INGENIERÍA DE SISTEMAS",
    "II": "INGENIERÍA EN INDUSTRIAS ALIMENTARIAS",
    "IN": "INGENIERÍA INDUSTRIAL",
    "IC": "INGENIERÍA CIVIL",
    "DE": "DERECHO",
    "CA": "CONTABILIDAD",
    "EN": "ENFERMERÍA",
    "PS": "PSICOLOGÍA",
    "MH": "MEDICINA HUMANA",
    "OB": "OBSTETRICIA",
    "AR": "ARQUITECTURA",
    "T1": "TECNOLOGÍA MÉDICA - ESPECIALIDAD EN LABORATORIO CLINICO Y ANATOMÍA PATOLÓGICA",
    "T3": "TECNOLOGÍA MÉDICA - ESPECIALIDAD EN TERAPIA FÍSICA Y REHABILITACIÓN",
    "T2": "TECNOLOGÍA MÉDICA - ESPECIALIDAD EN TERAPIA DE LENGUAJE",
    "T4": "TECNOLOGÍA MÉDICA - OPTOMETRÍA",
}
ESCUELA_FULL_SET = {v.upper() for v in ESCUELA_CODE_TO_FULL.values()}

# Modalidad: varias entradas -> V/P
MODALIDAD_MAP = {
    "V": "V", "VIRTUAL": "V",
    "P": "P", "PRESENCIAL": "P"
}

# Local base -> código (IC/CH/SU/HU, etc.)
LOCAL_TO_CODE_BASE = {
    "FILIAL": "IC", "ICA": "IC", "IC": "IC",
    "PRINCIPAL": "CH", "CHINCHA": "CH", "CH": "CH",
    "SUNAMPE": "SU", "SUMANPE": "SU", "SU": "SU",
    "HUARUA": "HU", "HU": "HU",
}

REQUERIDAS = ["DOCENTE", "DIA", "HORA INICIO", "HORA FIN"]

# ==============================
# UTILIDADES
# ==============================

def normalizar_texto(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def normalizar_mayus(x):
    return normalizar_texto(x).upper()

def convertir_hora(hora_raw):
    """
    Devuelve un datetime con fecha dummy (hoy) + hora.
    Acepta: 'HH:MM', 'HH:MM:SS', 'H:MM', '10:00 AM',
            números excel (0..1), enteros tipo 800/0830,
            y objetos datetime/time/pandas.Timestamp.
    """
    if pd.isna(hora_raw):
        raise ValueError("Hora vacía")

    if isinstance(hora_raw, (datetime, pd.Timestamp)):
        return datetime.combine(datetime.today().date(), hora_raw.time())
    if isinstance(hora_raw, dtime):
        return datetime.combine(datetime.today().date(), hora_raw)

    s = str(hora_raw).strip().upper()
    s = s.replace(".", ":").replace("H", ":")

    if s.isdigit():
        if len(s) in (1, 2):
            return datetime.strptime(f"{int(s):02d}:00", "%H:%M")
        if len(s) == 3:
            s = "0" + s
        if len(s) == 4:
            return datetime.strptime(f"{s[:2]}:{s[2:]}", "%H:%M")

    for fmt in ("%H:%M", "%H:%M:%S", "%I:%M %p", "%I:%M:%S %p"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass

    try:
        f = float(s)
        if 0 <= f < 1:
            total_min = int(round(f * 24 * 60))
            h = total_min // 60
            m = total_min % 60
            return datetime.strptime(f"{h:02d}:{m:02d}", "%H:%M")
    except ValueError:
        pass

    raise ValueError(f"Formato de hora no reconocido: {hora_raw}")

def hora_hhmm(v):
    return convertir_hora(v).strftime("%H:%M")

def parse_dia_to_num(d):
    """Acepta número (1-7), abreviaturas (LU..DO), nombres (con o sin tildes) -> 1..7"""
    if pd.isna(d):
        return None
    s = normalizar_mayus(d)
    # ¿número?
    try:
        n = int(float(s))
        if 1 <= n <= 7:
            return n
    except ValueError:
        pass
    # ¿abreviatura?
    abbr_to_num = {v: k for k, v in DIA_NUM_TO_ABBR.items()}
    if s in abbr_to_num:
        return abbr_to_num[s]
    # ¿nombre sin tildes?
    s2 = (s.replace("Á","A").replace("É","E").replace("Í","I").replace("Ó","O").replace("Ú","U"))
    if s2 in DIA_NAME_NONTILDE_TO_NUM:
        return DIA_NAME_NONTILDE_TO_NUM[s2]
    return None

def dia_num_to_abbr(n):   return DIA_NUM_TO_ABBR.get(n, "")
def dia_num_to_full_acc(n): return DIA_NUM_TO_NAME_ACCENT.get(n, "")

def normalizar_facultad(v):
    s = normalizar_mayus(v)
    if s in FACULTAD_MAP:
        return FACULTAD_MAP[s]
    return normalizar_texto(v)

def normalizar_escuela(v):
    s = normalizar_mayus(v)
    if s in ESCUELA_CODE_TO_FULL:
        return ESCUELA_CODE_TO_FULL[s]
    if s in ESCUELA_FULL_SET:
        return normalizar_texto(v)
    return normalizar_texto(v)

def normalizar_modalidad(v):
    s = normalizar_mayus(v)
    return MODALIDAD_MAP.get(s, s if s in ("V","P") else s)

def parse_custom_local_map(text):
    """
    Recibe líneas tipo 'NOMBRE=CODIGO'
    Retorna dict {UPPER(nombre): UPPER(codigo)}
    """
    out = {}
    for line in text.splitlines():
        line = line.strip()
        if not line or "=" not in line:
            continue
        k, v = line.split("=", 1)
        out[k.strip().upper()] = v.strip().upper()
    return out

def local_to_code(v, custom_map):
    s = normalizar_mayus(v)
    if s in custom_map:
        return custom_map[s]
    if s in LOCAL_TO_CODE_BASE:
        return LOCAL_TO_CODE_BASE[s]
    # Si ya viene como código de 2–3 letras, respetar
    if 2 <= len(s) <= 3 and s.isalpha():
        return s
    # Fallback: iniciales de palabras (ej. "SAN JUAN" -> "SJ")
    parts = [p for p in s.split() if p]
    if parts:
        cand = "".join(p[0] for p in parts)[:3].upper()
        return cand if cand else s
    return s

def seccion_letra(seccion):
    """
    Si SECCION tiene '-n', usa letra por el número (1->A, 2->B, ...).
    Si no, usa primera letra alfabética.
    """
    s = normalizar_mayus(seccion)
    if "-" in s:
        try:
            suf = s.split("-", 1)[1]
            n = int("".join(ch for ch in suf if ch.isdigit()))
            if n >= 1:
                idx = min(n, 26)
                return chr(64 + idx)  # 1->A, 2->B, ...
        except Exception:
            pass
    for ch in s:
        if ch.isalpha():
            return ch
    return ""

def construir_grupo(seccion, modalidad, local_code):
    letra = seccion_letra(seccion)
    mod = normalizar_modalidad(modalidad)
    mod = "V" if mod in ("V","VIRTUAL") else ("P" if mod in ("P","PRESENCIAL") else mod)
    return f"{letra}{mod} - {local_code}".strip()

def construir_tema_zoom(df_row, custom_local_map):
    """
    {PLAN}-{COD_CURSO}-{CURSO}-{SECCION}-{GRUPO_COMPACTO}-{DIA_ACC} {HH:MM}-{HH:MM}-{DURACION}|{DNI_DOC}
    * GRUPO_COMPACTO = (LETRA+V/P)-(LOCAL_CODE), sin espacios (ej: EV-SU)
    * Solo un '|', justo antes del DNI, y DNI al final.
    """
    plan = normalizar_texto(df_row.get("PLAN", "")) or normalizar_texto(df_row.get("COD_PLAN",""))
    cod_curso = normalizar_texto(df_row.get("COD_CURSO", ""))
    curso = normalizar_texto(df_row.get("CURSO", ""))
    seccion = normalizar_texto(df_row.get("SECCION", ""))

    # LOCAL code (según reglas y mapeos)
    local_code = local_to_code(df_row.get("LOCAL", ""), custom_local_map)

    # GRUPO y versión compacta (sin espacios): 'EV-SU'
    grupo = construir_grupo(seccion, df_row.get("MODALIDAD",""), local_code)
    grupo_compacto = grupo.replace(" ", "").replace("--", "-")  # 'EV-SU'

    # DIA con tildes
    n = parse_dia_to_num(df_row.get("DIA", ""))
    dia_acc = dia_num_to_full_acc(n) if n else ""

    hi = hora_hhmm(df_row.get("HORA INICIO", ""))
    hf = hora_hhmm(df_row.get("HORA FIN", ""))
    dur = normalizar_texto(df_row.get("DURACION", ""))

    left_parts = [plan, cod_curso, curso, seccion, grupo_compacto, f"{dia_acc} {hi}-{hf}-{dur}"]
    left = "-".join([p for p in left_parts if p != ""])
    dni_doc = normalizar_texto(df_row.get("DNI_DOC", ""))
    if dni_doc:
        return f"{left}|{dni_doc}"
    else:
        return left  # si no hay DNI, no ponemos '|'

def convertir_a_excel_export(df, custom_local_map):
    """Transformaciones de salida sobre copia del DF original (conserva orden de columnas donde aplica)."""
    out = df.copy()

    # DIA -> abreviatura LU..DO (para exportación de la columna DIA)
    out["DIA_NUM"] = out["DIA"].apply(parse_dia_to_num)
    out["DIA"] = out["DIA_NUM"].apply(dia_num_to_abbr)

    # FACULTAD nombre completo
    if "FACULTAD" in out.columns:
        out["FACULTAD"] = out["FACULTAD"].apply(normalizar_facultad)

    # ESCUELA nombre completo
    if "ESCUELA" in out.columns:
        out["ESCUELA"] = out["ESCUELA"].apply(normalizar_escuela)

    # MODALIDAD -> V/P
    if "MODALIDAD" in out.columns:
        out["MODALIDAD"] = out["MODALIDAD"].apply(normalizar_modalidad).map(
            lambda x: "V" if x in ("V","VIRTUAL") else ("P" if x in ("P","PRESENCIAL") else x)
        )

    # LOCAL -> código (IC/CH/SU/HU o mapeo personalizado)
    if "LOCAL" in out.columns:
        out["LOCAL"] = out["LOCAL"].apply(lambda v: local_to_code(v, custom_local_map))

    # GRUPO (con espacios)
    out["GRUPO"] = out.apply(
        lambda r: construir_grupo(r.get("SECCION",""), r.get("MODALIDAD",""), r.get("LOCAL","")),
        axis=1
    )

    # RANGO HORARIO (HH:MM-HH:MM)
    out["RANGO HORARIO"] = out.apply(
        lambda r: f"{hora_hhmm(r.get('HORA INICIO',''))}-{hora_hhmm(r.get('HORA FIN',''))}",
        axis=1
    )

    # TEMA_ZOOM (según tu formato final)
    out["TEMA_ZOOM"] = out.apply(lambda r: construir_tema_zoom(r, custom_local_map), axis=1)

    # DETALLE HORARIO usando nombre del día con tildes
    out["DETALLE HORARIO"] = out.apply(
        lambda r: f"{r.get('DIA_NUM','')} - {dia_num_to_full_acc(r.get('DIA_NUM'))} de - {r.get('RANGO HORARIO','')}",
        axis=1
    )

    return out

# ==============================
# ASIGNACIÓN DE ZOOM (motor)
# ==============================

def contar_conflictos(nueva_inicio, nueva_fin, asignaciones, margen_minutos):
    conflictos = 0
    for asignacion in asignaciones:
        existente_inicio = asignacion['inicio']
        existente_fin = asignacion['fin']
        if not (nueva_fin + timedelta(minutes=margen_minutos) <= existente_inicio or
                nueva_inicio >= existente_fin + timedelta(minutes=margen_minutos)):
            conflictos += 1
    return conflictos

def generar_correo_zoom(numero, prefijo):
    return f"{prefijo}{str(numero).zfill(4)}@autonomadeica.edu.pe"

def asignar_zoom(df, margen_minutos, max_simult, prefijo_correo):
    df = df.copy()
    df['Zoom asignado'] = None

    zoom_usos = {}  # {zoom_id: {DIA(full): [ {inicio, fin}, ... ]}}
    zoom_counter = 1

    for index, fila in df.iterrows():
        n = parse_dia_to_num(fila['DIA'])
        dia_key = dia_num_to_full_acc(n) if n else ""
        inicio = convertir_hora(fila['HORA INICIO'])
        fin = convertir_hora(fila['HORA FIN'])

        asignado = False
        for zoom_id, usos in zoom_usos.items():
            conflictos = contar_conflictos(inicio, fin, usos.get(dia_key, []), margen_minutos)
            if conflictos < max_simult:
                if dia_key not in usos:
                    usos[dia_key] = []
                usos[dia_key].append({'inicio': inicio, 'fin': fin})
                df.at[index, 'Zoom asignado'] = zoom_id
                asignado = True
                break

        if not asignado:
            nuevo_zoom = generar_correo_zoom(zoom_counter, prefijo_correo)
            zoom_counter += 1
            zoom_usos[nuevo_zoom] = {dia_key: [{'inicio': inicio, 'fin': fin}]}
            df.at[index, 'Zoom asignado'] = nuevo_zoom

    return df

def generar_resumen(df_asignado):
    tmp = df_asignado.copy()
    tmp["DIA_NUM"] = tmp["DIA"].apply(parse_dia_to_num)
    tmp["DIA_FULL"] = tmp["DIA_NUM"].apply(dia_num_to_full_acc)
    resumen = (
        tmp.groupby('DIA_FULL')
        .agg({'Zoom asignado': pd.Series.nunique})
        .rename(columns={'Zoom asignado': 'N° de Licencias Zoom'})
        .reset_index()
        .sort_values(by='DIA_FULL', key=lambda col: col.map({v:k for k,v in DIA_NUM_TO_NAME_ACCENT.items()}))
    )
    return resumen

# ==============================
# UI STREAMLIT
# ==============================

st.set_page_config(page_title="Asignador UAI", layout="wide")
st.title("📅 Asignador de Licencias Zoom - UAI")

# Sidebar: botón enlace
st.sidebar.header("VISITA TAMBIÉN")
url_moodle = "https://moodle-admision-kkvkzem6ls2m4f458ln4ut.streamlit.app/#exportador-de-admision-moodle"
st.sidebar.markdown(
    f"""
    <a href="{url_moodle}" target="_blank" style="text-decoration:none;">
      <span style="
        display:inline-block;
        padding:0.6rem 1rem;
        background:#d32f2f;
        color:#fff;
        border-radius:8px;
        font-weight:600;">
        Admisión Moodle
      </span>
    </a>
    """,
    unsafe_allow_html=True
)

# Sidebar: mapeos de locales personalizados
st.sidebar.subheader("Locales personalizados (opcional)")
custom_local_text = st.sidebar.text_area(
    "Formato: NOMBRE=CODIGO (una por línea). Ej: 'SEDE NUEVA=SN'",
    value="",
    height=120
)
custom_local_map = parse_custom_local_map(custom_local_text)

# Plantilla mínima (opcional)
st.subheader("📥 Descargar plantilla mínima")
plantilla_min = pd.DataFrame(columns=REQUERIDAS)
buf_plantilla = io.BytesIO()
with pd.ExcelWriter(buf_plantilla, engine='openpyxl') as writer:
    plantilla_min.to_excel(writer, index=False, sheet_name="Plantilla")
st.download_button(
    label="📄 Descargar plantilla Excel (mínima)",
    data=buf_plantilla.getvalue(),
    file_name="plantilla_horarios.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Subida de archivo
archivo = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx"])

# Controles
st.subheader("⚙️ Configuración")
margen_minutos = st.number_input("Margen (min) entre reuniones", min_value=0, max_value=60, value=10)
max_reuniones = st.radio("Máximo de reuniones simultáneas por licencia", options=[1, 2], index=1)
prefijo_correo = st.text_input("Prefijo de correo para generar licencias", value="UAI")

if archivo is not None and st.button("🚀 Asignar Zoom y Descargar"):
    # Leer Excel
    df = pd.read_excel(archivo)
    df.columns = [c.upper().strip() for c in df.columns]

    # Validación obligatorias
    faltantes = [c for c in REQUERIDAS if c not in df.columns]
    if faltantes:
        st.error("Faltan columnas obligatorias: " + ", ".join(faltantes))
        st.stop()

    # Normalizar DIA para lógica (a número)
    df["DIA"] = df["DIA"].apply(parse_dia_to_num)
    if df["DIA"].isna().any():
        bad = list((df[df["DIA"].isna()].index[:5] + 1).astype(int))
        st.error(f"DIA inválido en filas: {bad}. Acepta 1-7, LU..DO, LUNES..DOMINGO.")
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

    # Para exportar: convertir y calcular campos finales
    # OJO: df_asig["DIA"] aquí está en número; convertir_a_excel_export se encarga.
    df_export = convertir_a_excel_export(df_asig, custom_local_map)

    # Resumen
    df_resumen = generar_resumen(df_asig)

    # Mostrar tablas
    st.success("✅ Procesamiento completado")
    st.subheader("👩‍🏫 Horarios (con GRUPO y TEMA_ZOOM)")
    st.dataframe(df_export, use_container_width=True)

    st.subheader("📊 Resumen por día")
    st.dataframe(df_resumen, use_container_width=True)

    # Descargar Excel con 2 hojas
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        df_export.to_excel(writer, sheet_name='Horarios', index=False)
        df_resumen.to_excel(writer, sheet_name='Resumen', index=False)

    st.download_button(
        label="💾 Descargar archivo Excel",
        data=out.getvalue(),
        file_name="horario_con_zoom.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
