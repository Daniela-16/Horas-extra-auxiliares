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
JORNADA_SEMANAL_ESTANDAR = timedelta(hours=46)

# --- 3. Función para determinar el turno y sus horas de inicio/fin ajustadas ---
def obtener_turno_para_registro(fecha_hora_registro: datetime, tolerancia_minutos: int):
    dia_de_semana = fecha_hora_registro.weekday()
    tipo_dia = "LV" if dia_de_semana < 5 else "SAB"

    mejor_turno_encontrado = None
    min_diferencia_tiempo = timedelta(days=999)

    for nombre_turno, detalles_turno in TURNOS[tipo_dia].items():
        hora_inicio_turno_obj = datetime.strptime(detalles_turno["inicio"], "%H:%M:%S").time()
        hora_fin_turno_obj = datetime.strptime(detalles_turno["fin"], "%H:%M:%S").time()
        
        #Ajustar las horas al estandar de inicio de los turnos. quedan con la misma fecha

        candidatos_fecha_hora_inicio_turno = [
            fecha_hora_registro.replace(
                hour=hora_inicio_turno_obj.hour,
                minute=hora_inicio_turno_obj.minute,
                second=hora_inicio_turno_obj.second
            )
        ]
        if hora_inicio_turno_obj > hora_fin_turno_obj: # Turno nocturno
            candidatos_fecha_hora_inicio_turno.append(
                (fecha_hora_registro - timedelta(days=1)).replace(
                    hour=hora_inicio_turno_obj.hour,
                    minute=hora_inicio_turno_obj.minute,
                    second=hora_inicio_turno_obj.second
                )
            )

#este for itera entre los candidatos, 1 o 2 si es nocturno

        for inicio_candidato in candidatos_fecha_hora_inicio_turno:
            fin_candidato = inicio_candidato.replace(
                hour=hora_fin_turno_obj.hour,
                minute=hora_fin_turno_obj.minute,
                second=hora_fin_turno_obj.second
            )
            
        #si el turno es nocturno se le añade un dia a la fecha de fin 
            if hora_inicio_turno_obj > hora_fin_turno_obj:
                fin_candidato += timedelta(days=1)
                
        #creacion de la ventana de tiempo con la tolerancia 

            if not (inicio_candidato - timedelta(minutes=tolerancia_minutos) <=
                    fecha_hora_registro <=
                    fin_candidato + timedelta(minutes=tolerancia_minutos)):
                continue
        #fecha_hora_registro pasó la verificación de tolerancia, se calcula la diferencia absoluta 
        #de tiempo entre el registro del empleado y el inicio_candidato del turno. 
        #Esto dice qué tan cerca está el registro del inicio del turno.

            diferencia_tiempo = abs(fecha_hora_registro - inicio_candidato)

            if mejor_turno_encontrado is None or diferencia_tiempo < min_diferencia_tiempo:
                mejor_turno_encontrado = (nombre_turno, detalles_turno, inicio_candidato, fin_candidato)
                min_diferencia_tiempo = diferencia_tiempo

    return mejor_turno_encontrado if mejor_turno_encontrado else (None, None, None, None)

# --- 4. Función Principal para Calcular Horas Extras ---
def calcular_horas_extra(df_registros: pd.DataFrame, lugares_trabajo_normalizados: list, tolerancia_minutos: int):
    df_filtrado = df_registros[
        (df_registros['PORTERIA_NORMALIZED'].isin(lugares_trabajo_normalizados)) &
        (df_registros['PuntoMarcacion'].isin(['ent', 'sal']))
    ].sort_values(by=['COD_TRABAJADOR', 'FECHA_HORA_PROCESADA'])

    if df_filtrado.empty:
        return pd.DataFrame()

    resultados = []

    for (codigo_trabajador, fecha_dia_base), grupo in df_filtrado.groupby(['COD_TRABAJADOR', df_filtrado['FECHA_HORA_PROCESADA'].dt.date]):
        nombre_trabajador = grupo['NOMBRE'].iloc[0]
        entradas = grupo[grupo['PuntoMarcacion'] == 'ent']
        salidas = grupo[grupo['PuntoMarcacion'] == 'sal']

        if entradas.empty or salidas.empty:
            continue

        primera_entrada_hora_real = entradas['FECHA_HORA_PROCESADA'].min()
        ultima_salida_hora_real = salidas['FECHA_HORA_PROCESADA'].max()

        if ultima_salida_hora_real <= primera_entrada_hora_real:
            continue

#llamada de la funcion que determina el turno
        nombre_turno, detalles_turno, inicio_turno_calculado, fin_turno_calculado = \
            obtener_turno_para_registro(primera_entrada_hora_real, tolerancia_minutos)

        if nombre_turno is None:
            continue

        horas_trabajadas_td = ultima_salida_hora_real - inicio_turno_calculado if ultima_salida_hora_real > inicio_turno_calculado else timedelta(0)
        horas_trabajadas_hrs = horas_trabajadas_td.total_seconds() / 3600

        duracion_estandar_hrs = detalles_turno["duracion_hrs"]
        horas_extra = max(0, horas_trabajadas_hrs - duracion_estandar_hrs)

        if horas_extra < 0.5: # Umbral de 30 minutos
            horas_extra = 0.0

        resultados.append({
            'NOMBRE': nombre_trabajador,
            'COD_TRABAJADOR': codigo_trabajador,
            'FECHA': fecha_dia_base,
            'Dia_Semana': fecha_dia_base.strftime('%A'),
            'TURNO': nombre_turno,
            'Inicio_Turno_Programado': inicio_turno_calculado.strftime("%H:%M:%S"),
            'Fin_Turno_Programado': fin_turno_calculado.strftime("%H:%M:%S"),
            'Duracion_Turno_Programado_Hrs': duracion_estandar_hrs,
            'ENTRADA_AJUSTADA': inicio_turno_calculado.strftime("%Y-%m-%d %H:%M:%S"),
            'SALIDA_REAL': ultima_salida_hora_real.strftime("%Y-%m-%d %H:%M:%S"),
            'HORAS_TRABAJADAS_CALCULADAS_HRS': round(horas_trabajadas_hrs, 2),
            'HORAS_EXTRA_HRS': round(horas_extra, 2),
            'HORAS_EXTRA_ENTERAS_HRS': int(horas_extra),
            'MINUTOS_EXTRA_CONVERTIDOS': round((horas_extra - int(horas_extra)) * 60, 2)
        })
    return pd.DataFrame(resultados)

# --- Interfaz de usuario de Streamlit ---
st.set_page_config(page_title="Calculadora de Horas Extra", layout="wide")
st.title("📊 Calculadora de Horas Extra")
st.write("Sube tu archivo de Excel para calcular las horas extra de tus trabajadores.")

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

            # Ejecutar el cálculo
            st.subheader("Resultados del Cálculo")
            df_resultados_diarios = calcular_horas_extra(df_registros.copy(), LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS, TOLERANCIA_INFERENCIA_MINUTOS)

            if not df_resultados_diarios.empty:
                df_resultados_diarios_filtrado_extras = df_resultados_diarios[df_resultados_diarios['HORAS_EXTRA_HRS'] > 0].copy()

                if not df_resultados_diarios_filtrado_extras.empty:
                    st.write("### Reporte Horas Extra Diarias")
                    st.dataframe(df_resultados_diarios_filtrado_extras)

                    # Crear un buffer de Excel en memoria para el reporte diario
                    excel_buffer_diario = io.BytesIO()
                    df_resultados_diarios_filtrado_extras.to_excel(excel_buffer_diario, index=False, engine='openpyxl')
                    excel_buffer_diario.seek(0) # Regresar al inicio del buffer

                    st.download_button(
                        label="Descargar Reporte Horas Extra Diarias (Excel)",
                        data=excel_buffer_diario,
                        file_name="reporte_horas_extra_diarias.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.info("No se encontraron horas extras diarias para reportar.")

                # Generar Resumen Semanal
                df_resumen_semanal = pd.DataFrame()
                df_resultados_diarios['Semana_Inicio'] = df_resultados_diarios['FECHA'].apply(lambda x: x - timedelta(days=x.weekday()))

                df_resumen_semanal = df_resultados_diarios.groupby(['COD_TRABAJADOR', 'NOMBRE', 'Semana_Inicio']).agg(
                    Horas_Trabajadas_Calculadas_Semana_Hrs=('HORAS_TRABAJADAS_CALCULADAS_HRS', 'sum')
                ).reset_index()

                df_resumen_semanal['Semana_Inicio'] = pd.to_datetime(df_resumen_semanal['Semana_Inicio'])

                df_resumen_semanal['Horas_Extra_Semanales_Hrs'] = (
                    df_resumen_semanal['Horas_Trabajadas_Calculadas_Semana_Hrs'] - (JORNADA_SEMANAL_ESTANDAR.total_seconds() / 3600)
                ).apply(lambda x: max(0, round(x, 2)))

                df_resumen_semanal = df_resumen_semanal[df_resumen_semanal['Horas_Extra_Semanales_Hrs'] > 0].copy()

                df_resumen_semanal['Semana_Inicio'] = df_resumen_semanal['Semana_Inicio'].dt.strftime('%Y-%m-%d')
                df_resumen_semanal['Horas_Trabajadas_Calculadas_Semana_Hrs'] = round(df_resumen_semanal['Horas_Trabajadas_Calculadas_Semana_Hrs'], 2)

                df_resumen_semanal = df_resumen_semanal[[
                    'COD_TRABAJADOR', 'NOMBRE', 'Semana_Inicio',
                    'Horas_Trabajadas_Calculadas_Semana_Hrs', 'Horas_Extra_Semanales_Hrs',
                ]]

                if not df_resumen_semanal.empty:
                    st.write("### Resumen Horas Extra Semanal")
                    st.dataframe(df_resumen_semanal)

                    # Crear un buffer de Excel en memoria para el resumen semanal
                    excel_buffer_semanal = io.BytesIO()
                    df_resumen_semanal.to_excel(excel_buffer_semanal, index=False, engine='openpyxl')
                    excel_buffer_semanal.seek(0) # Regresar al inicio del buffer

                    st.download_button(
                        label="Descargar Resumen Horas Extra Semanal (Excel)",
                        data=excel_buffer_semanal,
                        file_name="resumen_horas_extra_semanal.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.info("No se encontraron horas extras semanales para reportar.")

            else:
                st.warning("No se pudieron calcular horas extras. Asegúrate de que el archivo Excel tenga los datos y formatos correctos.")

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}. Asegúrate de que el archivo es un Excel válido y la hoja 'BaseDatos Modificada' existe.")

st.markdown("---")
st.caption("Somos NOEL DE CORAZÓN ❤️ ")
