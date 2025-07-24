# -*- coding: utf-8 -*-
"""
Created on Mon Jul 21 09:20:21 2025

@author: NCGNpracpim
"""

import pandas as pd
from datetime import datetime, timedelta, time
import streamlit as st
import io # Importar io para manejar archivos en memoria

# --- 1. Definición de los Turnos ---
TURNOS = {
    "LV": { # Lunes a Viernes
        "Turno 1 LV": {"inicio": "05:40:00", "fin": "13:40:00", "duracion_hrs": 8},
        "Turno 2 LV": {"inicio": "13:40:00", "fin": "21:40:00", "duracion_hrs": 8},
        "Turno 3 LV": {"inicio": "21:40:00", "fin": "05:40:00", "duracion_hrs": 8},
    },
    "SAB": { # Sábados
        "Turno 1 SAB": {"inicio": "05:40:00", "fin": "11:40:00", "duracion_hrs": 6},
        "Turno 2 SAB": {"inicio": "11:40:00", "fin": "17:40:00", "duracion_hrs": 6},
        "Turno 3 SAB": {"inicio": "21:40:00", "fin": "05:40:00", "duracion_hrs": 8},
    }
}

# --- 2. Configuración General ---
LUGARES_TRABAJO_PRINCIPAL = [
    "NOEL_MDE_MR_WAFER_RCH_CREMAS_SAL", 
    "NOEL_MDE_OFIC_PRODUCCION_ENT", "NOEL_MDE_MR_TUNEL_VIENTO_2_ENT",
    "NOEL_MDE_OFIC_PRODUCCION_SAL", "NOEL_MDE_MR_MEZCLAS_ENT",
    "NOEL_MDE_MR_MEZCLAS_SAL", "NOEL_MDE_MR_HORNO_6-8-9_ENT",
    "NOEL_MDE_MR_HORNO_1-3_ENT", "NOEL_MDE_MR_WAFER_RCH_CREMAS_ENT",
    "NOEL_MDE_MR_HORNO_11_ENT","NOEL_MDE_MR_HORNO_6-8-9_SAL",
    "NOEL_MDE_MR_TUNEL_VIENTO_1_ENT","NOEL_MDE_MR_HORNO_18_SAL",
    "NOEL_MDE_MR_HORNO_1-3_SAL","NOEL_MDE_MR_HORNO_11_SAL",
    "NOEL_MDE_MR_ASPIRACION_ENT","NOEL_MDE_MR_HORNO_6-8-9_SAL_2", 
    "NOEL_MDE_ING_MEN_CREMAS_ENT", "NOEL_MDE_ING_MEN_CREMAS_SAL",
    "NOEL_MDE_ING_MENORES_2_ENT","NOEL_MDE_MR_HORNO_7-10_SAL",
    "NOEL_MDE_ING_MENORES_2_SAL", "NOEL_MDE_MR_HORNO_7-10_ENT", 
    "NOEL_MDE_ING_MENORES_1_ENT","NOEL_MDE_ESENCIAS_1_ENT",
    "NOEL_MDE_ING_MEN_ALERGENOS_ENT","NOEL_MDE_MR_HORNOS_SAL", 
    "NOEL_MDE_ESENCIAS_2_SAL","NOEL_MDE_ESENCIAS_1_SAL", 
    "NOEL_MDE_ING_MENORES_1_SAL","NOEL_MDE_MR_HORNO_4-5_ENT", "NOEL_MDE_MR_HORNO_2-12_ENT",
    "NOEL_MDE_MR_HORNOS_ENT"
]
    
LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS = [lugar.strip().lower() for lugar in LUGARES_TRABAJO_PRINCIPAL]
TOLERANCIA_INFERENCIA_MINUTOS = 30
JORNADA_SEMANAL_ESTANDAR = timedelta(hours=46) # Esta variable ya no se usará para cálculos de horas extra

# --- 3. Función para determinar el turno y sus horas de inicio/fin ajustadas ---
def obtener_turno_para_registro(fecha_hora_registro: datetime, tolerancia_minutos: int):
    # Determina si el día es entre semana (LV) o sábado (SAB)
    dia_de_semana = fecha_hora_registro.weekday()
    tipo_dia = "LV" if dia_de_semana < 5 else "SAB"

    mejor_turno_encontrado = None
    min_diferencia_tiempo = timedelta(days=999)

    # Itera sobre los turnos definidos para el tipo de día
    for nombre_turno, detalles_turno in TURNOS[tipo_dia].items():
        hora_inicio_turno_obj = datetime.strptime(detalles_turno["inicio"], "%H:%M:%S").time()
        hora_fin_turno_obj = datetime.strptime(detalles_turno["fin"], "%H:%M:%S").time()
        
        # Ajusta la fecha del inicio del turno al día del registro o al día anterior si es un turno nocturno
        candidatos_fecha_hora_inicio_turno = [
            fecha_hora_registro.replace(
                hour=hora_inicio_turno_obj.hour,
                minute=hora_inicio_turno_obj.minute,
                second=hora_inicio_turno_obj.second
            )
        ]
        # Si es un turno nocturno (inicio > fin), considera también el inicio en el día anterior
        if hora_inicio_turno_obj > hora_fin_turno_obj: # Turno nocturno
            candidatos_fecha_hora_inicio_turno.append(
                (fecha_hora_registro - timedelta(days=1)).replace(
                    hour=hora_inicio_turno_obj.hour,
                    minute=hora_inicio_turno_obj.minute,
                    second=hora_inicio_turno_obj.second
                )
            )

        # Itera sobre los candidatos de inicio de turno (uno para turnos diurnos, dos para nocturnos)
        for inicio_candidato in candidatos_fecha_hora_inicio_turno:
            # Calcula la hora de fin del turno candidato
            fin_candidato = inicio_candidato.replace(
                hour=hora_fin_turno_obj.hour,
                minute=hora_fin_turno_obj.minute,
                second=hora_fin_turno_obj.second
            )
            
            # Si el turno es nocturno, el fin del turno ocurre al día siguiente
            if hora_inicio_turno_obj > hora_fin_turno_obj:
                fin_candidato += timedelta(days=1)
                
            # Verifica si el registro cae dentro de la ventana del turno (inicio - tolerancia a fin + tolerancia)
            if not (inicio_candidato - timedelta(minutes=tolerancia_minutos) <=
                    fecha_hora_registro <=
                    fin_candidato + timedelta(minutes=tolerancia_minutos)):
                continue # Si no cae, pasa al siguiente turno candidato

            # Si el registro está dentro de la ventana, calcula la diferencia con el inicio del turno
            diferencia_tiempo = abs(fecha_hora_registro - inicio_candidato)

            # Si este turno es el mejor encontrado hasta ahora (el más cercano al inicio del turno)
            if mejor_turno_encontrado is None or diferencia_tiempo < min_diferencia_tiempo:
                mejor_turno_encontrado = (nombre_turno, detalles_turno, inicio_candidato, fin_candidato)
                min_diferencia_tiempo = diferencia_tiempo

    # Retorna el mejor turno encontrado con sus detalles y fechas/horas ajustadas, o None si no se encontró
    return mejor_turno_encontrado if mejor_turno_encontrado else (None, None, None, None)

# --- 4. Función Principal para Calcular Horas Extras (ahora solo asigna turnos) ---
def calcular_horas_extra(df_registros: pd.DataFrame, lugares_trabajo_normalizados: list, tolerancia_minutos: int):
    # Filtra los registros por lugares de trabajo principales y puntos de marcación 'ent' o 'sal'
    df_filtrado = df_registros[
        (df_registros['PORTERIA_NORMALIZED'].isin(lugares_trabajo_normalizados)) &
        (df_registros['PuntoMarcacion'].isin(['ent', 'sal']))
    ].sort_values(by=['COD_TRABAJADOR', 'FECHA_HORA_PROCESADA'])

    if df_filtrado.empty:
        return pd.DataFrame()

    resultados = []

    # Agrupa los registros por trabajador y por día para procesar cada jornada
    for (codigo_trabajador, fecha_dia_base), grupo in df_filtrado.groupby(['COD_TRABAJADOR', df_filtrado['FECHA_HORA_PROCESADA'].dt.date]):
        nombre_trabajador = grupo['NOMBRE'].iloc[0]
        entradas = grupo[grupo['PuntoMarcacion'] == 'ent']
        salidas = grupo[grupo['PuntoMarcacion'] == 'sal']

        # Si no hay entradas o salidas para la jornada, se salta
        if entradas.empty or salidas.empty:
            continue

        primera_entrada_hora_real = entradas['FECHA_HORA_PROCESADA'].min()
        ultima_salida_hora_real = salidas['FECHA_HORA_PROCESADA'].max()

        # Si la última salida es antes o igual a la primera entrada, se salta (datos inconsistentes)
        if ultima_salida_hora_real <= primera_entrada_hora_real:
            continue
        
        # NUEVA REGLA: Si la duración entre la primera entrada y la última salida es menor a 5 horas,
        # se considera una inconsistencia y se omite este registro.
        if (ultima_salida_hora_real - primera_entrada_hora_real) < timedelta(hours=5):
            # Opcional: Podrías añadir un log aquí para saber qué registros se están omitiendo
            # print(f"Registro omitido para {nombre_trabajador} en {fecha_dia_base} debido a duración menor a 5 horas.")
            continue


        # Llama a la función para determinar el turno basado en la primera entrada del día
        nombre_turno, detalles_turno, inicio_turno_calculado, fin_turno_calculado = \
            obtener_turno_para_registro(primera_entrada_hora_real, tolerancia_minutos)

        # Si no se pudo determinar un turno, se salta
        if nombre_turno is None:
            continue

        # Se mantienen los detalles del turno y las horas de entrada/salida reales y ajustadas

        resultados.append({
            'NOMBRE': nombre_trabajador,
            'COD_TRABAJADOR': codigo_trabajador,
            'FECHA': fecha_dia_base,
            'Dia_Semana': fecha_dia_base.strftime('%A'),
            'TURNO': nombre_turno,
            'Inicio_Turno_Programado': inicio_turno_calculado.strftime("%H:%M:%S"),
            'Fin_Turno_Programado': fin_turno_calculado.strftime("%H:%M:%S"),
            'Duracion_Turno_Programado_Hrs': detalles_turno["duracion_hrs"], # Duración estándar del turno
            'ENTRADA_AJUSTADA': inicio_turno_calculado.strftime("%Y-%m-%d %H:%M:%S"),
            'SALIDA_REAL': ultima_salida_hora_real.strftime("%Y-%m-%d %H:%M:%S")
        })
    return pd.DataFrame(resultados)

# --- Interfaz de usuario de Streamlit ---
st.set_page_config(page_title="Calculadora de Turnos", layout="wide") # Cambiado el título de la página
st.title("📊 Asignación de Turnos") # Cambiado el título de la aplicación
st.write("Sube tu archivo de Excel para ver la asignación de turnos de tus trabajadores.")

uploaded_file = st.file_uploader("Sube un archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Lee el archivo en memoria
        df_registros = pd.read_excel(uploaded_file, sheet_name='BaseDatos Modificada')

        columnas_requeridas = ['COD_TRABAJADOR', 'APUNTADOR', 'DESC_APUNTADOR', 'NOMBRE', 'FECHA', 'HORA', 'PORTERIA', 'PuntoMarcacion']
        if not all(col in df_registros.columns for col in columnas_requeridas):
            st.error(f"ERROR: Faltan columnas requeridas en la hoja 'BaseDatos Modificada'. Asegúrate de que existan: {', '.join(columnas_requeridas)}")
        else:
            # Preparación de datos consolidada
            df_registros['FECHA'] = pd.to_datetime(df_registros['FECHA'])
            df_registros['HORA'] = df_registros['HORA'].astype(str)
            df_registros['FECHA_HORA_PROCESADA'] = pd.to_datetime(df_registros['FECHA'].dt.strftime('%Y-%m-%d') + ' ' + df_registros['HORA'])
            df_registros['PORTERIA_NORMALIZED'] = df_registros['PORTERIA'].astype(str).str.strip().str.lower()
            df_registros['PuntoMarcacion'] = df_registros['PuntoMarcacion'].astype(str).str.strip().str.lower().replace({'entrada': 'ent', 'salida': 'sal'})
            df_registros.sort_values(by=['COD_TRABAJADOR', 'FECHA_HORA_PROCESADA'], inplace=True)
            df_registros.reset_index(drop=True, inplace=True)

            st.success("Archivo cargado y pre-procesado con éxito.")

            # Ejecutar el cálculo (ahora solo asigna turnos)
            st.subheader("Resultados de la Asignación de Turnos")
            df_resultados_diarios = calcular_horas_extra(df_registros.copy(), LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS, TOLERANCIA_INFERENCIA_MINUTOS)

            if not df_resultados_diarios.empty:
                st.write("### Reporte de Asignación de Turnos Diarios")
                st.dataframe(df_resultados_diarios)

                # Crear un buffer de Excel en memoria para el reporte diario
                excel_buffer_diario = io.BytesIO()
                df_resultados_diarios.to_excel(excel_buffer_diario, index=False, engine='openpyxl')
                excel_buffer_diario.seek(0) # Regresar al inicio del buffer

                st.download_button(
                    label="Descargar Reporte de Asignación de Turnos (Excel)",
                    data=excel_buffer_diario,
                    file_name="reporte_asignacion_turnos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("No se pudieron asignar turnos. Asegúrate de que el archivo Excel tenga los datos y formatos correctos y que las entradas/salidas no presenten inconsistencias de tiempo.")

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}. Asegúrate de que el archivo es un Excel válido y la hoja 'BaseDatos Modificada' existe.")

st.markdown("---")
st.caption("Somos NOEL DE CORAZÓN ❤️ ")


