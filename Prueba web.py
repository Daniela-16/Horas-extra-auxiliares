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
        "Turno 3 LV": {"inicio": "21:40:00", "fin": "05:40:00", "duracion_hrs": 8}, # Turno nocturno
    },
    "SAB": { # Sábados
        "Turno 1 SAB": {"inicio": "05:40:00", "fin": "11:40:00", "duracion_hrs": 6},
        "Turno 2 SAB": {"inicio": "11:40:00", "fin": "17:40:00", "duracion_hrs": 6},
        "Turno 3 SAB": {"inicio": "21:40:00", "fin": "05:40:00", "duracion_hrs": 8}, # Turno nocturno
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
TOLERANCIA_INFERENCIA_MINUTOS = 30 # Tolerancia en minutos para inferir el turno
JORNADA_SEMANAL_ESTANDAR = timedelta(hours=46) # Jornada estándar semanal

# --- 3. Función para determinar el turno y sus horas de inicio/fin ajustadas ---
def obtener_turno_para_registro(fecha_hora_registro: datetime, tolerancia_minutos: int):
    """
    Determina el turno al que pertenece un registro de entrada, considerando
    turnos diurnos y nocturnos que cruzan la medianoche.

    Args:
        fecha_hora_registro (datetime): Fecha y hora del registro de entrada.
        tolerancia_minutos (int): Minutos de tolerancia antes y después del inicio/fin del turno
                                  para considerar que un registro pertenece a un turno.

    Returns:
        tuple: (nombre_turno, detalles_turno, inicio_turno_calculado, fin_turno_calculado)
               Retorna (None, None, None, None) si no se encuentra un turno adecuado.
    """
    dia_de_semana = fecha_hora_registro.weekday() # 0 para Lunes, 6 para Domingo
    tipo_dia = "LV" if dia_de_semana < 5 else "SAB"

    mejor_turno_encontrado = None
    # Inicializar con una diferencia grande para asegurar que el primer turno válido sea seleccionado
    min_diferencia_tiempo = timedelta(days=999)

    # Iterar sobre los turnos definidos para el tipo de día (LV o SAB)
    for nombre_turno, detalles_turno in TURNOS[tipo_dia].items():
        # Convertir las horas de inicio y fin del turno a objetos time
        hora_inicio_turno_obj = datetime.strptime(detalles_turno["inicio"], "%H:%M:%S").time()
        hora_fin_turno_obj = datetime.strptime(detalles_turno["fin"], "%H:%M:%S").time()
        
        # Generar posibles fechas y horas de inicio del turno
        # Candidato 1: El turno inicia en la misma fecha del registro
        candidato_inicio_mismo_dia = fecha_hora_registro.replace(
            hour=hora_inicio_turno_obj.hour,
            minute=hora_inicio_turno_obj.minute,
            second=hora_inicio_turno_obj.second,
            microsecond=0
        )
        
        candidatos_fecha_hora_inicio_turno = [candidato_inicio_mismo_dia]
        
        # Para turnos nocturnos (fin_turno_obj es anterior a inicio_turno_obj),
        # también se considera que el turno pudo haber iniciado el día anterior al registro.
        if hora_inicio_turno_obj > hora_fin_turno_obj: # Es un turno nocturno
            candidato_inicio_dia_anterior = (fecha_hora_registro - timedelta(days=1)).replace(
                hour=hora_inicio_turno_obj.hour,
                minute=hora_inicio_turno_obj.minute,
                second=hora_inicio_turno_obj.second,
                microsecond=0
            )
            candidatos_fecha_hora_inicio_turno.append(candidato_inicio_dia_anterior)

        # Evaluar cada candidato de inicio de turno
        for inicio_candidato in candidatos_fecha_hora_inicio_turno:
            # Calcular la fecha y hora de fin del turno para este candidato de inicio
            fin_candidato = inicio_candidato.replace(
                hour=hora_fin_turno_obj.hour,
                minute=hora_fin_turno_obj.minute,
                second=hora_fin_turno_obj.second,
                microsecond=0
            )
            
            # Si es un turno nocturno, el fin del turno ocurre al día siguiente
            if hora_inicio_turno_obj > hora_fin_turno_obj:
                fin_candidato += timedelta(days=1)
            
            # Definir la ventana de tiempo del turno con tolerancia
            # El registro debe caer dentro de esta ventana para ser considerado parte del turno
            shift_start_with_tolerance = inicio_candidato - timedelta(minutes=tolerancia_minutos)
            shift_end_with_tolerance = fin_candidato + timedelta(minutes=tolerancia_minutos)
            
            # Verificar si el registro está dentro de la ventana de tolerancia del turno
            if shift_start_with_tolerance <= fecha_hora_registro <= shift_end_with_tolerance:
                # Si está dentro, calcular la diferencia absoluta entre el registro y el inicio del turno
                # para encontrar el turno más cercano.
                diferencia_tiempo = abs(fecha_hora_registro - inicio_candidato)

                # Si es el primer turno encontrado o este es más cercano, actualizar el mejor turno
                if mejor_turno_encontrado is None or diferencia_tiempo < min_diferencia_tiempo:
                    mejor_turno_encontrado = (nombre_turno, detalles_turno, inicio_candidato, fin_candidato)
                    min_diferencia_tiempo = diferencia_tiempo
    
    return mejor_turno_encontrado if mejor_turno_encontrado else (None, None, None, None)

# --- 4. Función Principal para Calcular Horas Extras ---
def calcular_horas_extra(df_registros: pd.DataFrame, lugares_trabajo_normalizados: list, tolerancia_minutos: int):
    """
    Calcula las horas extras de los trabajadores basándose en sus registros
    de entrada y salida y la definición de turnos.

    Args:
        df_registros (pd.DataFrame): DataFrame con los registros de los empleados.
        lugares_trabajo_normalizados (list): Lista de lugares de trabajo principales normalizados.
        tolerancia_minutos (int): Minutos de tolerancia para la inferencia de turnos.

    Returns:
        pd.DataFrame: DataFrame con los resultados de las horas extras.
    """
    # Filtrar registros por lugares de trabajo principales y tipos de marcación válidos ('ent', 'sal')
    df_filtrado = df_registros[
        (df_registros['PORTERIA_NORMALIZED'].isin(lugares_trabajo_normalizados)) &
        (df_registros['PuntoMarcacion'].isin(['ent', 'sal']))
    ].sort_values(by=['COD_TRABAJADOR', 'FECHA_HORA_PROCESADA']) # Ordenar para asegurar la secuencia correcta

    if df_filtrado.empty:
        return pd.DataFrame()

    resultados = [] # Lista para almacenar los resultados procesados
    
    # Procesar registros por cada trabajador individualmente
    for codigo_trabajador, employee_df in df_filtrado.groupby('COD_TRABAJADOR'):
        nombre_trabajador = employee_df['NOMBRE'].iloc[0] # El nombre del trabajador es el mismo para todos sus registros
        
        current_entry_record = None # Almacena el último registro de entrada no emparejado
        
        # Iterar a través de los registros del empleado de forma secuencial para encontrar pares de entrada/salida
        for idx, row in employee_df.iterrows():
            if row['PuntoMarcacion'] == 'ent':
                # Si se encuentra una entrada, la guardamos como el inicio de una posible nueva sesión
                current_entry_record = row
            elif row['PuntoMarcacion'] == 'sal':
                # Si se encuentra una salida y hay un registro de entrada previo no emparejado
                if current_entry_record is not None:
                    entrada_time = current_entry_record['FECHA_HORA_PROCESADA']
                    salida_time_original = row['FECHA_HORA_PROCESADA'] # Hora de salida original del registro

                    # Ajustar la hora de salida para el cálculo de duración si es un turno nocturno
                    # (la salida es numéricamente anterior pero en el día siguiente)
                    salida_time_for_duration = salida_time_original
                    if salida_time_for_duration < entrada_time:
                        salida_time_for_duration += timedelta(days=1)

                    # Determinar el turno basado en la hora de entrada de este par
                    nombre_turno, detalles_turno, inicio_turno_calculado, fin_turno_calculado = \
                        obtener_turno_para_registro(entrada_time, tolerancia_minutos)
                    
                    # Si no se pudo asignar un turno estándar al registro de entrada, se ignora este par
                    if nombre_turno is None:
                        current_entry_record = None # Reiniciar para el siguiente par
                        continue

                    # Calcular las horas trabajadas reales para este par de entrada-salida
                    total_worked_duration_td = salida_time_for_duration - entrada_time
                    total_worked_hours = total_worked_duration_td.total_seconds() / 3600.0

                    # Obtener la duración estándar del turno asignado
                    duracion_estandar_hrs = detalles_turno["duracion_hrs"]
                    
                    # Calcular horas extras: tiempo trabajado menos duración estándar del turno
                    horas_extra = max(0, total_worked_hours - duracion_estandar_hrs)

                    # Aplicar el umbral de 30 minutos (0.5 horas) para considerar horas extras
                    if horas_extra < 0.5:
                        horas_extra = 0.0

                    # Añadir los resultados a la lista
                    resultados.append({
                        'NOMBRE': nombre_trabajador,
                        'COD_TRABAJADOR': codigo_trabajador,
                        'FECHA_REGISTRO_ENTRADA': entrada_time.strftime('%Y-%m-%d'), # Fecha del registro de entrada
                        'HORA_ENTRADA_REAL': entrada_time.strftime('%H:%M:%S'),      # Hora del registro de entrada
                        'FECHA_REGISTRO_SALIDA': salida_time_original.strftime('%Y-%m-%d'), # Fecha del registro de salida
                        'HORA_SALIDA_REAL': salida_time_original.strftime('%H:%M:%S'),      # Hora del registro de salida
                        'Dia_Semana_Entrada': entrada_time.strftime('%A'),           # Día de la semana de la entrada
                        'TURNO_ASIGNADO': nombre_turno,                              # Nombre del turno asignado
                        'Inicio_Turno_Programado_Calculado': inicio_turno_calculado.strftime("%Y-%m-%d %H:%M:%S"), # Inicio calculado del turno
                        'Fin_Turno_Programado_Calculado': fin_turno_calculado.strftime("%Y-%m-%d %H:%M:%S"),     # Fin calculado del turno
                        'Duracion_Turno_Programado_Hrs': duracion_estandar_hrs,      # Duración estándar del turno
                        'HORAS_TRABAJADAS_PAR_ENTRADA_SALIDA_HRS': round(total_worked_hours, 2), # Horas trabajadas para este par
                        'HORAS_EXTRA_HRS': round(horas_extra, 2),                    # Horas extras calculadas
                        'HORAS_EXTRA_ENTERAS_HRS': int(horas_extra),                 # Horas extras enteras
                        'MINUTOS_EXTRA_CONVERTIDOS': round((horas_extra - int(horas_extra)) * 60, 2) # Minutos extras
                    })
                    
                    current_entry_record = None # Reiniciar el registro de entrada para el siguiente par
                # Si se encuentra una 'sal' pero no hay 'ent' previo, se ignora (registro anómalo)
    return pd.DataFrame(resultados)

# --- Interfaz de usuario de Streamlit ---
st.set_page_config(page_title="Calculadora de Horas Extra", layout="wide")
st.title("📊 Calculadora de Horas Extra")
st.write("Sube tu archivo de Excel para calcular las horas extra de tus trabajadores.")

uploaded_file = st.file_uploader("Sube un archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Lee el archivo en memoria
        # Asegúrate de que la hoja se llama 'BaseDatos Modificada' como se espera
        df_registros = pd.read_excel(uploaded_file, sheet_name='BaseDatos Modificada')

        columnas_requeridas = ['COD_TRABAJADOR', 'APUNTADOR', 'DESC_APUNTADOR', 'NOMBRE', 'FECHA', 'HORA', 'PORTERIA', 'PuntoMarcacion']
        if not all(col in df_registros.columns for col in columnas_requeridas):
            st.error(f"ERROR: Faltan columnas requeridas en la hoja 'BaseDatos Modificada'. Asegúrate de que existan: {', '.join(columnas_requeridas)}")
        else:
            # Preparación de datos: normalización y combinación de fecha/hora
            df_registros['FECHA'] = pd.to_datetime(df_registros['FECHA'])
            df_registros['HORA'] = df_registros['HORA'].astype(str) # Asegurar que la hora es string
            # Combinar FECHA y HORA en un solo Timestamp para procesamiento
            df_registros['FECHA_HORA_PROCESADA'] = pd.to_datetime(df_registros['FECHA'].dt.strftime('%Y-%m-%d') + ' ' + df_registros['HORA'])
            
            # Normalizar la columna 'PORTERIA' a minúsculas y sin espacios
            df_registros['PORTERIA_NORMALIZED'] = df_registros['PORTERIA'].astype(str).str.strip().str.lower()
            # Normalizar 'PuntoMarcacion' a 'ent' o 'sal'
            df_registros['PuntoMarcacion'] = df_registros['PuntoMarcacion'].astype(str).str.strip().str.lower().replace({'entrada': 'ent', 'salida': 'sal'})
            
            # Ordenar por trabajador y fecha/hora para asegurar el emparejamiento correcto de entradas y salidas
            df_registros.sort_values(by=['COD_TRABAJADOR', 'FECHA_HORA_PROCESADA'], inplace=True)
            df_registros.reset_index(drop=True, inplace=True)

            st.success("Archivo cargado y pre-procesado con éxito.")

            # Ejecutar el cálculo de horas extras
            st.subheader("Resultados del Cálculo")
            df_resultados_diarios = calcular_horas_extra(df_registros.copy(), LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS, TOLERANCIA_INFERENCIA_MINUTOS)

            if not df_resultados_diarios.empty:
                # Filtrar solo los registros con horas extras > 0 para el reporte diario
                df_resultados_diarios_filtrado_extras = df_resultados_diarios[df_resultados_diarios['HORAS_EXTRA_HRS'] > 0].copy()

                if not df_resultados_diarios_filtrado_extras.empty:
                    st.write("### Reporte Horas Extra Diarias por Par Entrada-Salida")
                    st.dataframe(df_resultados_diarios_filtrado_extras)

                    # Opción para descargar el reporte diario
                    excel_buffer_diario = io.BytesIO()
                    df_resultados_diarios_filtrado_extras.to_excel(excel_buffer_diario, index=False, engine='openpyxl')
                    excel_buffer_diario.seek(0) # Volver al inicio del buffer para la descarga

                    st.download_button(
                        label="Descargar Reporte Horas Extra Diarias (Excel)",
                        data=excel_buffer_diario,
                        file_name="reporte_horas_extra_diarias.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.info("No se encontraron horas extras diarias para reportar.")

                # Generar Resumen Semanal de Horas Trabajadas y Horas Extra
                df_resumen_semanal = pd.DataFrame()
                # Calcular el inicio de la semana (Lunes) para cada fecha de entrada
                df_resultados_diarios['Semana_Inicio'] = df_resultados_diarios['FECHA_REGISTRO_ENTRADA'].apply(lambda x: pd.to_datetime(x) - timedelta(days=pd.to_datetime(x).weekday()))

                # Agrupar por trabajador y semana para sumar las horas trabajadas
                df_resumen_semanal = df_resultados_diarios.groupby(['COD_TRABAJADOR', 'NOMBRE', 'Semana_Inicio']).agg(
                    Horas_Trabajadas_Calculadas_Semana_Hrs=('HORAS_TRABAJADAS_PAR_ENTRADA_SALIDA_HRS', 'sum')
                ).reset_index()

                # Asegurar que 'Semana_Inicio' es de tipo datetime para operaciones posteriores si fuera necesario
                df_resumen_semanal['Semana_Inicio'] = pd.to_datetime(df_resumen_semanal['Semana_Inicio'])

                # Calcular horas extras semanales (si exceden la jornada estándar)
                df_resumen_semanal['Horas_Extra_Semanales_Hrs'] = (
                    df_resumen_semanal['Horas_Trabajadas_Calculadas_Semana_Hrs'] - (JORNADA_SEMANAL_ESTANDAR.total_seconds() / 3600)
                ).apply(lambda x: max(0, round(x, 2))) # Asegurar que no haya valores negativos

                # Filtrar solo las semanas con horas extras semanales > 0
                df_resumen_semanal = df_resumen_semanal[df_resumen_semanal['Horas_Extra_Semanales_Hrs'] > 0].copy()

                # Formatear la columna 'Semana_Inicio' para la visualización
                df_resumen_semanal['Semana_Inicio'] = df_resumen_semanal['Semana_Inicio'].dt.strftime('%Y-%m-%d')
                # Redondear las horas trabajadas totales de la semana
                df_resumen_semanal['Horas_Trabajadas_Calculadas_Semana_Hrs'] = round(df_resumen_semanal['Horas_Trabajadas_Calculadas_Semana_Hrs'], 2)

                # Reordenar columnas para el reporte final
                df_resumen_semanal = df_resumen_semanal[[
                    'COD_TRABAJADOR', 'NOMBRE', 'Semana_Inicio',
                    'Horas_Trabajadas_Calculadas_Semana_Hrs', 'Horas_Extra_Semanales_Hrs',
                ]]

                if not df_resumen_semanal.empty:
                    st.write("### Resumen Horas Extra Semanal")
                    st.dataframe(df_resumen_semanal)

                    # Opción para descargar el resumen semanal
                    excel_buffer_semanal = io.BytesIO()
                    df_resumen_semanal.to_excel(excel_buffer_semanal, index=False, engine='openpyxl')
                    excel_buffer_semanal.seek(0)

                    st.download_button(
                        label="Descargar Resumen Horas Extra Semanal (Excel)",
                        data=excel_buffer_semanal,
                        file_name="resumen_horas_extra_semanal.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.info("No se encontraron horas extras semanales para reportar.")

            else:
                st.warning("No se pudieron calcular horas extras. Asegúrate de que el archivo Excel tenga los datos y formatos correctos y que haya pares de entrada/salida válidos.")

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}. Asegúrate de que el archivo es un Excel válido y la hoja 'BaseDatos Modificada' existe.")

st.markdown("---")
st.caption("Somos NOEL DE CORAZÓN ❤️ ")

