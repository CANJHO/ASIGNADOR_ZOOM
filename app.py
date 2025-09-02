st.sidebar.markdown(
    """
    <a href="https://moodle-admision-kkvkzem6ls2m4f458ln4ut.streamlit.app/#exportador-de-admision-moodle"
       target="_blank" style="text-decoration:none;">
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

# app.py
import io
from datetime import datetime, timedelta, time as dtime
import pandas as pd
import streamlit as st

# ==============================
# CONSTANTES Y MAPEOS
# ==============================
DIA_NUM = {
    'LUNES': 1, 'MARTES': 2,
    'MIERCOLES': 3, 'MIÉRCOLES': 3,
    'JUEVES': 4, 'VIERNES': 5,
    'SABADO': 6, 'SÁBADO': 6,
    'DOMINGO': 7
}
REQUERIDAS = ["DOCENTE", "DIA", "HORA INICIO", "HORA FIN"]

# ==============================
# UTILIDADES
# ==============================
def normalizar_dias(valor):
    if pd.isna(valor):
        return valor
    v = str(valor).strip().upper()
    # Unificar tildes
    v = (v.replace("Á", "A").replace("É", "E")
           .replace("Í", "I").replace("Ó", "O")
           .replace("Ú", "U"))
    return v

def convertir_hora(hora_raw):
    """
    Devuelve un datetime con fecha dummy (hoy) + hora.
    Acepta: 'HH:MM', 'HH:MM:SS', 'H:MM', '10:00 AM',
            números excel (0..1), enteros tipo 800/0830,
            y objetos datetime/time/pandas.Timestamp.
    """
    if pd.isna(hora_raw):
        raise ValueError("Hora vacía")

    # Ya datetime/time
    if isinstance(hora_raw, (datetime, pd.Timestamp)):
        return datetime.combine(datetime.today().date(), hora_raw.time())
    if isinstance(hora_raw, dtime):
        return datetime.combine(datetime.today().date(), hora_raw)

    s = str(hora_raw).strip().upper()
    s = s.replace(".", ":").replace("H", ":")  # '10.00' / '10h00' -> '10:00'

    # Enteros 9 / 09 / 800 / 0830
    if s.isdigit():
        if len(s) in (1, 2):      # '9' -> '09:00'
            return datetime.strptime(f"{int(s):02d}:00", "%H:%M")
        if len(s) == 3:           # '800' -> '08:00'
            s = "0" + s
        if len(s) == 4:           # '0830' -> '08:30'
            return datetime.strptime(f"{s[:2]}:{s[2:]}", "%H:%M")

    # Formatos texto comunes
    for fmt in ("%H:%M", "%H:%M:%S", "%I:%M %p", "%I:%M:%S %p"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass

    # Número Excel (fracción de día)
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
    df['RANGO HORARIO'] = None
    df['DETALLE HORARIO'] = None
    df['DIA_NUM'] = None

    zoom_usos = {}  # {zoom_id: {DIA: [ {inicio, fin}, ... ]}}
    zoom_counter = 1

    for index, fila in df.iterrows():
        dia = normalizar_dias(fila['DIA'])
        inicio_str = str(fila['HORA INICIO']).strip()
        fin_str   = str(fila['HORA FIN']).strip()
        inicio = convertir_hora(inicio_str)
        fin = convertir_hora(fin_str)

        asignado = False
        for zoom_id, usos in zoom_usos.items():
            conflictos = contar_conflictos(inicio, fin, usos.get(dia, []), margen_minutos)
            if conflictos < max_simult:
                if dia not in usos:
                    usos[dia] = []
                usos[dia].append({'inicio': inicio, 'fin': fin})
                df.at[index, 'Zoom asignado'] = zoom_id
                asignado = True
                break

        if not asignado:
            nuevo_zoom = generar_correo_zoom(zoom_counter, prefijo_correo)
            zoom_counter += 1
            zoom_usos[nuevo_zoom] = {dia: [{'inicio': inicio, 'fin': fin}]}
            df.at[index, 'Zoom asignado'] = nuevo_zoom

        rango_horario = f"{inicio_str}-{fin_str}"
        dia_num = DIA_NUM.get(dia, '?')
        dia_nombre = dia.capitalize()
        df.at[index, 'RANGO HORARIO'] = rango_horario
        df.at[index, 'DIA_NUM'] = dia_num
        df.at[index, 'DETALLE HORARIO'] = f"{dia_num} - {dia_nombre} de - {rango_horario}"

    return df

def generar_resumen(df):
    resumen = (
        df.groupby('DIA')
        .agg({'Zoom asignado': pd.Series.nunique})
        .rename(columns={'Zoom asignado': 'N° de Licencias Zoom'})
        .reset_index()
        .sort_values(by='DIA', key=lambda col: col.str.upper().map(DIA_NUM))
    )
    return resumen

def validar_horas(df):
    errores = []
    for col in ["HORA INICIO", "HORA FIN"]:
        for idx, val in df[col].items():
            try:
                convertir_hora(val)
            except Exception:
                # guardamos muestra (1-index en UI)
                errores.append((col, idx + 1, str(val)))
                if len(errores) >= 6:  # no saturar
                    return errores
    return errores

# ==============================
# UI STREAMLIT
# ==============================
st.title("📅 Asignador de Licencias Zoom - UAI")

# Plantilla descargable
st.subheader("📥 Descargar plantilla")
plantilla = pd.DataFrame(columns=REQUERIDAS)
plantilla_file = io.BytesIO()
with pd.ExcelWriter(plantilla_file, engine='openpyxl') as writer:
    plantilla.to_excel(writer, index=False, sheet_name="Plantilla")
st.download_button(
    label="📄 Descargar plantilla Excel (DOCENTE, DIA, HORA INICIO, HORA FIN)",
    data=plantilla_file.getvalue(),
    file_name="plantilla_horarios.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Subida de archivo
archivo = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx"])

# Controles
st.subheader("⚙️ Configuración")
margen_minutos = st.number_input(
    "Margen de minutos entre reuniones", min_value=0, max_value=60, value=10
)
max_reuniones = st.radio(
    "Máximo de reuniones simultáneas por licencia", options=[1, 2], index=1
)
prefijo_correo = st.text_input(
    "Prefijo de correo para generar licencias", value="UAI"
)

# Procesar
if archivo is not None and st.button("🚀 Asignar Zoom"):
    df = pd.read_excel(archivo)
    # Normalizar encabezados y valores de DIA
    df.columns = [col.upper().strip() for col in df.columns]
    if "DIA" in df.columns:
        df["DIA"] = df["DIA"].apply(normalizar_dias)

    # Validar columnas
    faltantes = [c for c in REQUERIDAS if c not in df.columns]
    if faltantes:
        st.error(
            "Faltan columnas obligatorias: "
            + ", ".join(faltantes)
            + ". Descarga la plantilla para asegurarte del formato."
        )
        st.stop()

    # Validar horas (muestra primeros problemas)
    problemas = validar_horas(df)
    if problemas:
        ejemplos = "; ".join([f"{c} fila {i}: '{v}'" for c, i, v in problemas[:5]])
        st.error(
            f"Se encontraron horas con formato no reconocido. Ejemplos -> {ejemplos}."
            " Aceptados: 'HH:MM', 'HH:MM:SS', '10:00 AM', 0830, fracción Excel."
        )
        st.stop()

    # Ejecutar lógica
    df_resultado = asignar_zoom(df, margen_minutos, max_reuniones, prefijo_correo)

    # Ordenar por día y hora (usamos columna temporal con hora parseada)
    df_resultado["_HORA_INICIO_DT"] = df_resultado["HORA INICIO"].apply(convertir_hora)
    df_resultado = df_resultado.sort_values(by=["DIA_NUM", "_HORA_INICIO_DT"])
    df_resultado = df_resultado.drop(columns=["_HORA_INICIO_DT"])

    df_resumen = generar_resumen(df_resultado)

    # Mostrar tablas
    st.success("✅ Procesamiento completado")
    st.subheader("👩‍🏫 Horarios (incluye DOCENTE)")
    st.dataframe(df_resultado, use_container_width=True)

    st.subheader("📊 Resumen por día")
    st.dataframe(df_resumen, use_container_width=True)

    # Archivo Excel para descargar
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_resultado.to_excel(writer, sheet_name='Horarios', index=False)
        df_resumen.to_excel(writer, sheet_name='Resumen', index=False)

    st.download_button(
        label="💾 Descargar archivo Excel",
        data=output.getvalue(),
        file_name="horario_con_zoom.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
