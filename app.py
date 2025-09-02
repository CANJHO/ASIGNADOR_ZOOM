import io
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st

# ==============================
# FUNCIONES AUXILIARES
# ==============================

DIA_NUM = {
    'LUNES': 1, 'MARTES': 2,
    'MIERCOLES': 3, 'MIÉRCOLES': 3,
    'JUEVES': 4, 'VIERNES': 5,
    'SABADO': 6, 'SÁBADO': 6,
    'DOMINGO': 7
}

REQUERIDAS = ["DOCENTE", "DIA", "HORA INICIO", "HORA FIN"]

def convertir_hora(hora_str: str) -> datetime:
    return datetime.strptime(str(hora_str).strip(), "%H:%M")

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
    # columnas calculadas
    df['Zoom asignado'] = None
    df['RANGO HORARIO'] = None
    df['DETALLE HORARIO'] = None
    df['DIA_NUM'] = None

    zoom_usos = {}  # {zoom_id: {DIA: [ {inicio, fin}, ... ]}}
    zoom_counter = 1

    for index, fila in df.iterrows():
        dia = str(fila['DIA']).strip().upper()
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

# ==============================
# STREAMLIT APP
# ==============================

st.title("📅 Asignador de Licencias Zoom - UAI")

# --- Plantilla descargable ---
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

# --- Subida de archivo ---
archivo = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx"])

st.subheader("⚙️ Configuración")

margen_minutos = st.number_input(
    "Margen de minutos entre reuniones",
    min_value=0, max_value=60, value=10
)

max_reuniones = st.radio(
    "Máximo de reuniones simultáneas por licencia",
    options=[1, 2],
    index=1
)

prefijo_correo = st.text_input(
    "Prefijo de correo para generar licencias",
    value="UAI"
)

if archivo is not None and st.button("🚀 Asignar Zoom"):
    # Leer Excel
    df = pd.read_excel(archivo)
    df.columns = [col.upper().strip() for col in df.columns]

    # Validar columnas requeridas (incluye DOCENTE)
    faltantes = [c for c in REQUERIDAS if c not in df.columns]
    if faltantes:
        st.error(f"Faltan columnas obligatorias: {', '.join(faltantes)}. "
                 "Descarga la plantilla para asegurarte del formato.")
        st.stop()

    # Ejecutar lógica
    df_resultado = asignar_zoom(df, margen_minutos, max_reuniones, prefijo_correo)
    df_resultado = df_resultado.sort_values(by=['DIA_NUM', 'HORA INICIO'])
    df_resumen = generar_resumen(df_resultado)

    # Mostrar tablas
    st.success("✅ Procesamiento completado")
    st.subheader("👩‍🏫 Horarios (incluye DOCENTE)")
    st.dataframe(df_resultado)

    st.subheader("📊 Resumen por día")
    st.dataframe(df_resumen)

    # Generar archivo Excel para descargar
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
