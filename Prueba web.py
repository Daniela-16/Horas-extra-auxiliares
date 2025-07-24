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
JORNADA_SEMANAL_ESTANDAR = timedelta(hours=46) # Esta variable ya no se usa para el resumen semanal

# --- 3. Función para determinar el turno, sus horas ajustadas y la fecha lógica del turno ---
def obtener_turno_para_registro(fecha_hora_registro: datetime, tolerancia_minutos: int):
    dia_de_semana = fecha_hora_registro.weekday()
    tipo_dia = "LV" if dia_de_semana < 5 else "SAB"

    mejor_turno_encontrado = None
    min_diferencia_tiempo = timedelta(days=999)
    shift_logical_date = None # Fecha lógica de la jornada (fecha de inicio del turno)

    for nombre_turno, detalles_turno in TURNOS[tipo_dia].items():
        hora_inicio_turno_obj = datetime.strptime(detalles_turno["inicio"], "%H:%M:%S").time()
        hora_fin_turno_obj = datetime.strptime(detalles_turno["fin"], "%H:%M:%S").time()
        
        candidatos_fecha_hora_inicio_turno = []

        # Candidato 1: El turno comienza el mismo día del registro
        inicio_candidato_mismodia = fecha_hora_registro.replace(
            hour=hora_inicio_turno_obj.hour,
            minute=hora_inicio_turno_obj.minute,
            second=hora_inicio_turno_obj.second
        )
        candidatos_fecha_hora_inicio_turno.append(inicio_candidato_mismodia)
        
        # Candidato 2: Si es un turno nocturno, el registro actual (entrada o salida) 
        # podría pertenecer a un turno que empezó el día anterior.
        if hora_inicio_turno_obj > hora_fin_turno_obj: # Es un turno nocturno (ej. 21:40 a 05:40)
            inicio_candidato_diapre = (fecha_hora_registro - timedelta(days=1)).replace(
                hour=hora_inicio_turno_obj.hour,
                minute=hora_inicio_turno_obj.minute,
                second=hora_inicio_turno_obj.second
            )
            candidatos_fecha_hora_inicio_turno.append(inicio_candidato_diapre)
        
        # Iterar sobre los candidatos de fecha de inicio para encontrar el mejor ajuste
        for inicio_candidato in candidatos_fecha_hora_inicio_turno:
            fin_candidato = inicio_candidato.replace(
                hour=hora_fin_turno_obj.hour,
                minute=hora_fin_turno_obj.minute,
                second=hora_fin_turno_obj.second
            )
            
            # Ajustar la fecha de fin si es un turno nocturno
            if hora_inicio_turno_obj > hora_fin_turno_obj:
                fin_candidato += timedelta(days=1)
            
            # Verificar si el registro cae dentro de la ventana de tolerancia del turno candidato
            if (inicio_candidato - timedelta(minutes=tolerancia_minutos) <= 
                fecha_hora_registro <= 
                fin_candidato + timedelta(minutes=tolerancia_minutos)):
                
                # Calcular la diferencia de tiempo para encontrar el mejor turno (el más cercano al inicio)
                diferencia_tiempo = abs(fecha_hora_registro - inicio_candidato)

                if mejor_turno_encontrado is None or diferencia_tiempo < min_diferencia_tiempo:
                    mejor_turno_encontrado = (nombre_turno, detalles_turno, inicio_candidato, fin_candidato)
                    min_diferencia_tiempo = diferencia_tiempo
                    # La fecha lógica del turno es la fecha de inicio del turno calculado
                    shift_logical_date = inicio_candidato.date() 
    
    if mejor_turno_encontrado:
        # Devolver el turno, sus detalles, inicio/fin calculados y la fecha lógica del turno
        return mejor_turno_encontrado + (shift_logical_date,) 
    else:
        return (None, None, None, None, None) # Devolver None para la fecha lógica si no se encuentra turno

# --- 4. Función Principal para Calcular Horas Extras ---
def calcular_horas_extra(df_registros: pd.DataFrame, lugares_trabajo_normalizados: list, tolerancia_minutos: int):
    # Filtrar registros relevantes y normalizar la portería y PuntoMarcacion
    df_filtrado = df_registros[
        (df_registros['PORTERIA_NORMALIZED'].isin(lugares_trabajo_normalizados)) &
        (df_registros['PuntoMarcacion'].isin(['ent', 'sal']))
    ].copy() # Usar .copy() para evitar SettingWithCopyWarning

    if df_filtrado.empty:
        return pd.DataFrame()

    # Aplica la función para determinar el turno y la fecha lógica de la jornada
    # Esto es clave para manejar los turnos nocturnos que cruzan la medianoche
    df_filtrado[['TURNO', 'DETALLES_TURNO', 'INICIO_TURNO_CALCULADO', 'FIN_TURNO_CALCULADO', 'SHIFT_LOGICAL_DATE']] = \
        df_filtrado['FECHA_HORA_PROCESADA'].apply(
            lambda x: pd.Series(obtener_turno_para_registro(x, tolerancia_minutos))
        )
    
    # Eliminar registros que no pudieron ser asociados a ningún turno
    df_filtrado.dropna(subset=['SHIFT_LOGICAL_DATE'], inplace=True)
    df_filtrado['SHIFT_LOGICAL_DATE'] = df_filtrado['SHIFT_LOGICAL_DATE'].astype(str).astype('datetime64[ns]').dt.date # Asegurar tipo de fecha

    # Ordenar por trabajador y fecha/hora para un procesamiento secuencial
    df_filtrado.sort_values(by=['COD_TRABAJADOR', 'FECHA_HORA_PROCESADA'], inplace=True)

    resultados = []

    # Agrupar por trabajador y por la fecha lógica del turno para consolidar entradas y salidas de una jornada
    for (codigo_trabajador, shift_logical_date), grupo in df_filtrado.groupby(['COD_TRABAJADOR', 'SHIFT_LOGICAL_DATE']):
        nombre_trabajador = grupo['NOMBRE'].iloc[0]
        entradas = grupo[grupo['PuntoMarcacion'] == 'ent']
        salidas = grupo[grupo['PuntoMarcacion'] == 'sal']

        if entradas.empty or salidas.empty:
            continue

        primera_entrada_hora_real = entradas['FECHA_HORA_PROCESADA'].min()
        ultima_salida_hora_real = salidas['FECHA_HORA_PROCESADA'].max()

        # Re-inferir el turno basado en la primera entrada real del grupo
        # Esto es para obtener los detalles correctos del turno (inicio, fin programados) para esta jornada
        nombre_turno, detalles_turno, inicio_turno_programado_calc, fin_turno_programado_calc, _ = \
            obtener_turno_para_registro(primera_entrada_hora_real, tolerancia_minutos)
        
        if nombre_turno is None: # Si por alguna razón no se puede inferir el turno con la entrada real, saltar
            continue

        # Asegurarse de que la salida real sea posterior al inicio calculado del turno
        # Para evitar cálculos negativos si la salida es atípica o un error
        if ultima_salida_hora_real <= inicio_turno_programado_calc:
            horas_trabajadas_td = timedelta(0)
        else:
            horas_trabajadas_td = ultima_salida_hora_real - inicio_turno_programado_calc
        
        horas_trabajadas_hrs = horas_trabajadas_td.total_seconds() / 3600

        duracion_estandar_hrs = detalles_turno["duracion_hrs"]
        horas_extra = max(0, horas_trabajadas_hrs - duracion_estandar_hrs)

        if horas_extra < 0.5: # Umbral de 30 minutos (0.5 horas) para considerar horas extra
            horas_extra = 0.0

        resultados.append({
            'NOMBRE': nombre_trabajador,
            'COD_TRABAJADOR': codigo_trabajador,
            'FECHA': shift_logical_date, # Usamos la fecha lógica del turno como la fecha del reporte
            'Dia_Semana': shift_logical_date.strftime('%A'),
            'TURNO': nombre_turno,
            'Inicio_Turno_Programado': inicio_turno_programado_calc.strftime("%H:%M:%S"),
            'Fin_Turno_Programado': fin_turno_programado_calc.strftime("%H:%M:%S"),
            'Duracion_Turno_Programado_Hrs': duracion_estandar_hrs,
            'ENTRADA_AJUSTADA': primera_entrada_hora_real.strftime("%Y-%m-%d %H:%M:%S"),
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
            # Combinar FECHA y HORA en un solo campo datetime
            df_registros['FECHA_HORA_PROCESADA'] = df_registros.apply(
                lambda row: datetime.combine(row['FECHA'].date(), datetime.strptime(str(row['HORA']), "%H:%M").time()), axis=1
            )
            df_registros['PORTERIA_NORMALIZED'] = df_registros['PORTERIA'].astype(str).str.strip().str.lower()
            df_registros['PuntoMarcacion'] = df_registros['PuntoMarcacion'].astype(str).str.strip().str.lower().replace({'entrada': 'ent', 'salida': 'sal'})
            
            # Es importante ordenar ANTES de calcular las fechas lógicas de turno para asegurar coherencia
            df_registros.sort_values(by=['COD_TRABAJADOR', 'FECHA_HORA_PROCESADA'], inplace=True)
            df_registros.reset_index(drop=True, inplace=True)

            st.success("Archivo cargado y pre-procesado con éxito.")

            # Ejecutar el cálculo
            st.subheader("Resultados del Cálculo")
            df_resultados_diarios = calcular_horas_extra(df_registros.copy(), LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS, TOLERANCIA_INFERENCIA_MINUTOS)

            if not df_resultados_diarios.empty:
                # Filtrar turnos diurnos y nocturnos para la presentación y exportación
                turnos_nocturnos_nombres = ["Turno 3 LV", "Turno 3 SAB"]
                
                df_resultados_diurnos_extras = df_resultados_diarios[
                    (df_resultados_diarios['HORAS_EXTRA_HRS'] > 0) & 
                    (~df_resultados_diarios['TURNO'].isin(turnos_nocturnos_nombres))
                ].copy()

                df_resultados_nocturnos_extras = df_resultados_diarios[
                    (df_resultados_diarios['HORAS_EXTRA_HRS'] > 0) & 
                    (df_resultados_diarios['TURNO'].isin(turnos_nocturnos_nombres))
                ].copy()

                # Reporte de Horas Extra Diarias (Diurnas)
                if not df_resultados_diurnos_extras.empty:
                    st.write("### Reporte Horas Extra Diarias (Turnos Diurnos)")
                    st.dataframe(df_resultados_diurnos_extras)
                else:
                    st.info("No se encontraron horas extras diarias en turnos diurnos para reportar.")

                # Reporte de Horas Extra Diarias (Nocturnas)
                if not df_resultados_nocturnos_extras.empty:
                    st.write("### Reporte Horas Extra Diarias (Turnos Nocturnos)")
                    st.dataframe(df_resultados_nocturnos_extras)
                else:
                    st.info("No se encontraron horas extras diarias en turnos nocturnos para reportar.")

                # Crear un buffer de Excel en memoria para ambos reportes en hojas separadas
                if not df_resultados_diurnos_extras.empty or not df_resultados_nocturnos_extras.empty:
                    excel_buffer_diario_nocturno = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer_diario_nocturno, engine='openpyxl') as writer:
                        if not df_resultados_diurnos_extras.empty:
                            df_resultados_diurnos_extras.to_excel(writer, sheet_name='Horas_Extra_Diurnas', index=False)
                        if not df_resultados_nocturnos_extras.empty:
                            df_resultados_nocturnos_extras.to_excel(writer, sheet_name='Horas_Extra_Nocturnas', index=False)
                    excel_buffer_diario_nocturno.seek(0) # Regresar al inicio del buffer

                    st.download_button(
                        label="Descargar Reporte Horas Extra Diarias (Diurnas y Nocturnas - Excel)",
                        data=excel_buffer_diario_nocturno,
                        file_name="reporte_horas_extra_diarias_diurnas_nocturnas.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.info("No hay horas extras (diurnas ni nocturnas) para generar el reporte combinado.")
            else:
                st.warning("No se pudieron calcular horas extras. Asegúrate de que el archivo Excel tenga los datos y formatos correctos.")

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}. Asegúrate de que el archivo es un Excel válido y la hoja 'BaseDatos Modificada' existe.")

st.markdown("---")
st.caption("Somos NOEL DE CORAZÓN ❤️ ")
