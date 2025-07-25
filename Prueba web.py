# -*- coding: utf-8 -*-
"""
Created on Mon Jul 21 09:20:21 2025

@author: NCGNpracpim
"""

import pandas as pd
from datetime import datetime, timedelta
import streamlit as st
import io

# --- 1. Definición de los Turnos ---
TURNOS = {
    "LV": {
        "Turno 1 LV": {"inicio": "05:40:00", "fin": "13:40:00", "duracion_hrs": 8},
        "Turno 2 LV": {"inicio": "13:40:00", "fin": "21:40:00", "duracion_hrs": 8},
        "Turno 3 LV": {"inicio": "21:40:00", "fin": "05:40:00", "duracion_hrs": 8},
    },
    "SAB": {
        "Turno 1 SAB": {"inicio": "05:40:00", "fin": "11:40:00", "duracion_hrs": 6},
        "Turno 2 SAB": {"inicio": "11:40:00", "fin": "17:40:00", "duracion_hrs": 6},
        "Turno 3 SAB": {"inicio": "21:40:00", "fin": "05:40:00", "duracion_hrs": 8},
    }
}

# --- 2. Configuración General ---
LUGARES_TRABAJO_PRINCIPAL = ["NOEL_MDE_OFIC_PRODUCCION_ENT",
    "NOEL_MDE_OFIC_PRODUCCION_SAL",
    "NOEL_MDE_MR_TUNEL_VIENTO_1_ENT",
    "NOEL_MDE_MR_MEZCLAS_ENT",
    "NOEL_MDE_MR_WAFER_RCH_CREMAS_ENT",
    "NOEL_MDE_MR_HORNO_11_ENT",
    "NOEL_MDE_MR_MEZCLAS_SAL",
    "NOEL_MDE_MR_HORNO_6-8-9_SAL",
    "NOEL_MDE_MR_SERVICIOS_2_SAL",
    "NOEL_MDE_MR_TUNEL_VIENTO_2_ENT",
    "NOEL_MDE_MR_SERVICIOS_2_ENT",
    "NOEL_MDE_MR_HORNO_1-3_ENT",
    "NOEL_MDE_MR_WAFER_RCH_CREMAS_SAL",
    "NOEL_MDE_MR_HORNO_6-8-9_ENT",
    "NOEL_MDE_MR_HORNO_11_SAL",
    "NOEL_MDE_MR_HORNOS_ENT",
    "NOEL_MDE_MR_HORNO_2-12_ENT",
    "NOEL_MDE_ING_MEN_CREMAS_ENT",
    "NOEL_MDE_ING_MEN_CREMAS_SAL",
    "NOEL_MDE_ING_MEN_ALERGENOS_ENT",
    "NOEL_MDE_MR_HORNO_4-5_ENT",
    "NOEL_MDE_ESENCIAS_2_SAL",
    "NOEL_MDE_ESENCIAS_1_ENT",
    "NOEL_MDE_ESENCIAS_1_SAL",
    "NOEL_MDE_MR_HORNO_6-8-9_SAL_2",
    "NOEL_MDE_MR_ASPIRACION_ENT",
    "NOEL_MDE_ING_MENORES_1_SAL",
    "NOEL_MDE_ING_MENORES_2_ENT",
    "NOEL_MDE_ING_MENORES_2_SAL",
    "NOEL_MDE_MR_HORNO_1-3_SAL",
    "NOEL_MDE_MR_HORNO_18_ENT",
    "NOEL_MDE_MR_HORNO_18_SAL",
    "NOEL_MDE_MR_HORNOS_SAL",
    "NOEL_MDE_ING_MENORES_1_ENT",
    "NOEL_MDE_MR_HORNO_7-10_SAL",
    "NOEL_MDE_MR_HORNO_7-10_ENT"
]
LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS = [lugar.strip().lower() for lugar in LUGARES_TRABAJO_PRINCIPAL]
TOLERANCIA_INFERENCIA_MINUTOS = 50
MAX_EXCESO_SALIDA_HRS = 3

# --- 3. Obtener turno basado en fecha y hora ---
def obtener_turno_para_registro(fecha_hora_evento: datetime, tolerancia_minutos: int):
    dia_semana = fecha_hora_evento.weekday()
    tipo_dia = "LV" if dia_semana < 5 else "SAB"

    mejor_turno = None
    menor_diferencia = timedelta(days=999)

    for nombre_turno, info_turno in TURNOS[tipo_dia].items():
        hora_inicio = datetime.strptime(info_turno["inicio"], "%H:%M:%S").time()
        hora_fin = datetime.strptime(info_turno["fin"], "%H:%M:%S").time()

        candidatos_inicio = [fecha_hora_evento.replace(hour=hora_inicio.hour, minute=hora_inicio.minute, second=hora_inicio.second)]

        if hora_inicio > hora_fin:
            candidatos_inicio.append((fecha_hora_evento - timedelta(days=1)).replace(hour=hora_inicio.hour, minute=hora_inicio.minute, second=hora_inicio.second))

        for inicio in candidatos_inicio:
            fin = inicio.replace(hour=hora_fin.hour, minute=hora_fin.minute, second=hora_fin.second)
            if hora_inicio > hora_fin:
                fin += timedelta(days=1)

            if not (inicio - timedelta(minutes=tolerancia_minutos) <= fecha_hora_evento <= fin + timedelta(minutes=tolerancia_minutos)):
                continue

            diferencia = abs(fecha_hora_evento - inicio)

            if mejor_turno is None or diferencia < menor_diferencia:
                mejor_turno = (nombre_turno, info_turno, inicio, fin)
                menor_diferencia = diferencia

    return mejor_turno if mejor_turno else (None, None, None, None)

# --- 4. Calculo de horas ---
def calcular_turnos(df: pd.DataFrame, lugares_normalizados: list, tolerancia_minutos: int):
    df = df[(df['PORTERIA_NORMALIZADA'].isin(lugares_normalizados)) & (df['TIPO_MARCACION'].isin(['ent', 'sal']))]
    df.sort_values(by=['ID_TRABAJADOR', 'FECHA_HORA'], inplace=True)

    if df.empty:
        return pd.DataFrame()

    resultados = []

    for (id_trabajador, fecha_base), grupo in df.groupby(['ID_TRABAJADOR', df['FECHA_HORA'].dt.date]):
        nombre = grupo['NOMBRE'].iloc[0]
        entradas = grupo[grupo['TIPO_MARCACION'] == 'ent']
        salidas = grupo[grupo['TIPO_MARCACION'] == 'sal']

        if entradas.empty or salidas.empty:
            continue

        entrada_real = entradas['FECHA_HORA'].min()
        salida_real = salidas['FECHA_HORA'].max()

        if salida_real <= entrada_real or (salida_real - entrada_real) < timedelta(hours=5):
            continue

        turno_nombre, info_turno, inicio_turno, fin_turno = obtener_turno_para_registro(entrada_real, tolerancia_minutos)
        if turno_nombre is None:
            continue

        if salida_real > fin_turno + timedelta(hours=MAX_EXCESO_SALIDA_HRS):
            continue

        duracion_real = salida_real - entrada_real
        horas_trabajadas = round(duracion_real.total_seconds() / 3600, 2)
        horas_turno = info_turno["duracion_hrs"]
        horas_extra = max(0, round(horas_trabajadas - horas_turno, 2))

        resultados.append({
            'NOMBRE': nombre,
            'ID_TRABAJADOR': id_trabajador,
            'FECHA': fecha_base,
            'Dia_Semana': fecha_base.strftime('%A'),
            'TURNO': turno_nombre,
            'Inicio_Turno_Programado': inicio_turno.strftime("%H:%M:%S"),
            'Fin_Turno_Programado': fin_turno.strftime("%H:%M:%S"),
            'Duracion_Turno_Programado_Hrs': horas_turno,
            'ENTRADA_REAL': entrada_real.strftime("%Y-%m-%d %H:%M:%S"),
            'SALIDA_REAL': salida_real.strftime("%Y-%m-%d %H:%M:%S"),
            'Horas_Trabajadas': horas_trabajadas,
            'Horas_Extra': horas_extra,
            'Horas_Extra_Enteras': int(horas_extra),
            'Minutos_Extra': round((horas_extra - int(horas_extra)) * 60)
        })

    return pd.DataFrame(resultados)

# --- Interfaz Streamlit ---
st.set_page_config(page_title="Calculadora de Horas Extra", layout="wide")
st.title("📊 Calculadora de Horas Extra")
st.write("Sube tu archivo de Excel para calcular las horas extra del personal.")

archivo_excel = st.file_uploader("Sube un archivo Excel (.xlsx)", type=["xlsx"])

if archivo_excel is not None:
    try:
        df_raw = pd.read_excel(archivo_excel, sheet_name='BaseDatos Modificada')

        columnas = ['COD_TRABAJADOR', 'APUNTADOR', 'DESC_APUNTADOR', 'NOMBRE', 'FECHA', 'HORA', 'PORTERIA', 'PuntoMarcacion']
        if not all(col in df_raw.columns for col in columnas):
            st.error(f"ERROR: Faltan columnas requeridas: {', '.join(columnas)}")
        else:
            df_raw['FECHA'] = pd.to_datetime(df_raw['FECHA'])
            df_raw['HORA'] = df_raw['HORA'].astype(str)
            df_raw['FECHA_HORA'] = pd.to_datetime(df_raw['FECHA'].dt.strftime('%Y-%m-%d') + ' ' + df_raw['HORA'])
            df_raw['PORTERIA_NORMALIZADA'] = df_raw['PORTERIA'].astype(str).str.strip().str.lower()
            df_raw['TIPO_MARCACION'] = df_raw['PuntoMarcacion'].astype(str).str.strip().str.lower().replace({'entrada': 'ent', 'salida': 'sal'})
            df_raw.rename(columns={'COD_TRABAJADOR': 'ID_TRABAJADOR'}, inplace=True)

            st.success("Archivo cargado y preprocesado con éxito.")
            df_resultado = calcular_turnos(df_raw.copy(), LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS, TOLERANCIA_INFERENCIA_MINUTOS)

            if not df_resultado.empty:
                st.subheader("Resultados de las horas extra")
                st.dataframe(df_resultado)

                buffer_excel = io.BytesIO()
                df_resultado.to_excel(buffer_excel, index=False, engine='openpyxl')
                buffer_excel.seek(0)

                st.download_button(
                    label="Descargar Reporte de Horas extra (Excel)",
                    data=buffer_excel,
                    file_name="reporte_horas_extra.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("No se pudieron asignar turnos o hubo inconsistencias en los registros.")

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")

st.markdown("---")
st.caption("Somos NOEL DE CORAZÓN ❤️")
