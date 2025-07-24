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
        "Turno 3 LV": {"inicio": "21:40:00", "fin": "05:40:00", "duracion_hrs": 8}, # Turno Nocturno
    },
    "SAB": { # Sábados
        "Turno 1 SAB": {"inicio": "05:40:00", "fin": "11:40:00", "duracion_hrs": 6},
        "Turno 2 SAB": {"inicio": "11:40:00", "fin": "17:40:00", "duracion_hrs": 6},
        "Turno 3 SAB": {"inicio": "21:40:00", "fin": "05:40:00", "duracion_hrs": 6}, # Turno Nocturno
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
TOLERANCIA_INFERENCIA_MINUTOS = 30 # Se mantiene en 30 minutos, ya que el pre-procesamiento corrige la fecha
JORNADA_SEMANAL_ESTANDAR = timedelta(hours=46)

# --- Función para ajustar la fecha de salidas mal registradas para turnos nocturnos ---
def adjust_misdated_night_shift_exits(df: pd.DataFrame):
    """
    Ajusta la fecha de los registros de 'salida' que aparecen incorrectamente en el mismo
    día calendario que una 'entrada' de turno nocturno posterior, pero lógicamente
    corresponden al día siguiente.

    Esto maneja el escenario:
    Registro 1: Salida, 21/06/2025 05:46:00
    Registro 2: Entrada, 21/06/2025 21:05:00

    En este caso, la fecha del Registro 1 se ajustará a 22/06/2025 05:46:00.
    """
    df_adjusted = df.copy()
    
    # Iterar a través de los registros de cada trabajador
    for cod_trabajador, group in df.groupby('COD_TRABAJADOR'):
        # Asegurarse de que el grupo esté ordenado cronológicamente para la detección de patrones
        group_sorted = group.sort_values(by='FECHA_HORA_PROCESADA').reset_index(drop=True)
        
        for i in range(len(group_sorted) - 1): # Iterar hasta el penúltimo registro
            current_record = group_sorted.iloc[i]
            next_record = group_sorted.iloc[i+1]
            
            # Condición 1: El registro actual es una 'salida' en la madrugada (antes de las 7 AM)
            is_early_morning_exit = (current_record['PuntoMarcacion'] == 'sal' and 
                                     current_record['FECHA_HORA_PROCESADA'].time() < time(7, 0, 0))
            
            # Condición 2: El siguiente registro es una 'entrada' para un turno nocturno
            # más tarde en el *mismo día calendario* del registro de salida actual.
            is_same_day_night_entry = (next_record['PuntoMarcacion'] == 'ent' and
                                       next_record['FECHA_HORA_PROCESADA'].date() == current_record['FECHA_HORA_PROCESADA'].date() and
                                       next_record['FECHA_HORA_PROCESADA'].hour >= 20) # Asumiendo entradas de turno nocturno desde las 8 PM
            
            if is_early_morning_exit and is_same_day_night_entry:
                # Si se cumplen estas condiciones, significa que esta salida de madrugada
                # es probablemente el final del turno que *comienza* con la próxima entrada nocturna.
                # Por lo tanto, necesitamos avanzar la fecha de este registro de 'salida' en un día.
                
                # Usar .loc con el índice original para modificar el DataFrame principal df_adjusted
                original_index = current_record.name # Obtener el índice original del DataFrame principal
                df_adjusted.loc[original_index, 'FECHA_HORA_PROCESADA'] += timedelta(days=1)
    
    # Después de todos los ajustes, reordenar el DataFrame completo para asegurar
    # el orden cronológico correcto para el cálculo de turnos.
    return df_adjusted.sort_values(by=['COD_TRABAJADOR', 'FECHA_HORA_PROCESADA']).reset_index(drop=True)

# --- 3. Función para determinar el turno y sus horas de inicio/fin ajustadas ---
def obtener_turno_para_registro(fecha_hora_registro: datetime, tolerancia_minutos: int):
    """
    Determina el turno al que pertenece un registro de hora, ajustando para turnos nocturnos.
    Prioriza turnos nocturnos del día anterior para registros de salida de madrugada.
    """
    dia_de_semana = fecha_hora_registro.weekday() # 0=Lunes, 6=Domingo
    
    # Determina el tipo de día para el registro actual y el día anterior
    tipo_dia_actual = "LV" if dia_de_semana < 5 else "SAB"
    dia_anterior = fecha_hora_registro - timedelta(days=1)
    tipo_dia_anterior = "LV" if dia_anterior.weekday() < 5 else "SAB"

    tolerancia = timedelta(minutes=tolerancia_minutos)
    
    # Lista para almacenar todos los turnos posibles a los que podría pertenecer el registro
    posibles_turnos = [] # Almacena {nombre_turno, detalles, inicio_real, fin_real, diferencia_a_inicio, es_turno_nocturno}

    # 1. Evaluar turnos que inician en el día calendario del registro
    for nombre_turno, detalles_turno in TURNOS[tipo_dia_actual].items():
        hora_inicio_turno_obj = datetime.strptime(detalles_turno["inicio"], "%H:%M:%S").time()
        hora_fin_turno_obj = datetime.strptime(detalles_turno["fin"], "%H:%M:%S").time()
        
        # Construye el inicio y fin real del turno basándose en la fecha del registro
        inicio_real_candidato = fecha_hora_registro.replace(
            hour=hora_inicio_turno_obj.hour,
            minute=hora_inicio_turno_obj.minute,
            second=hora_inicio_turno_obj.second
        )
        fin_real_candidato = inicio_real_candidato.replace(
            hour=hora_fin_turno_obj.hour,
            minute=hora_fin_turno_obj.minute,
            second=hora_fin_turno_obj.second
        )
        es_turno_nocturno = False
        if hora_inicio_turno_obj > hora_fin_turno_obj: # Si es un turno nocturno (ej: 21:00 a 05:00)
            fin_real_candidato += timedelta(days=1) # El fin real es al día siguiente
            es_turno_nocturno = True

        # Verifica si el registro cae dentro de la ventana de este turno (inicio a fin, con tolerancia)
        if (inicio_real_candidato - tolerancia <= fecha_hora_registro <= fin_real_candidato + tolerancia):
            posibles_turnos.append({
                'nombre_turno': nombre_turno,
                'detalles': detalles_turno,
                'inicio_real': inicio_real_candidato,
                'fin_real': fin_real_candidato,
                'diferencia_a_inicio': abs(fecha_hora_registro - inicio_real_candidato),
                'es_turno_nocturno': es_turno_nocturno
            })

    # 2. Evaluar turnos nocturnos que comenzaron el día ANTERIOR y terminan en la madrugada del día del registro.
    # Esto es crucial para registros de salida como el ejemplo (05:46 AM, fecha del día del registro).
    # Solo se aplica si la hora del registro es temprano en la mañana (ej. antes de las 7 AM).
    if fecha_hora_registro.time() < time(7, 0, 0): 
        for nombre_turno, detalles_turno in TURNOS[tipo_dia_anterior].items():
            hora_inicio_turno_obj = datetime.strptime(detalles_turno["inicio"], "%H:%M:%S").time()
            hora_fin_turno_obj = datetime.strptime(detalles_turno["fin"], "%H:%M:%S").time()

            if hora_inicio_turno_obj > hora_fin_turno_obj: # Solo consideramos turnos nocturnos del día anterior
                # El inicio real de este turno sería el día anterior
                inicio_real_candidato = (fecha_hora_registro - timedelta(days=1)).replace(
                    hour=hora_inicio_turno_obj.hour,
                    minute=hora_inicio_turno_obj.minute,
                    second=hora_inicio_turno_obj.second
                )
                # El fin real de este turno sería en la madrugada del día del registro
                fin_real_candidato = inicio_real_candidato.replace(
                    hour=hora_fin_turno_obj.hour,
                    minute=hora_fin_turno_obj.minute,
                    second=hora_fin_turno_obj.second
                ) + timedelta(days=1) 

                if (inicio_real_candidato - tolerancia <= fecha_hora_registro <= fin_real_candidato + tolerancia):
                    posibles_turnos.append({
                        'nombre_turno': nombre_turno,
                        'detalles': detalles_turno,
                        'inicio_real': inicio_real_candidato,
                        'fin_real': fin_real_candidato,
                        'diferencia_a_inicio': abs(fecha_hora_registro - inicio_real_candidato),
                        'es_turno_nocturno': True
                    })
    
    if not posibles_turnos:
        return None, None, None, None

    # Lógica de Priorización para seleccionar el mejor turno entre los posibles:
    # 1. Priorizar específicamente turnos nocturnos que comenzaron el día anterior
    #    si el registro es una marcación en la madrugada (antes de las 7 AM)
    #    y el registro cae dentro de la tolerancia del FINAL de ese turno nocturno.
    #    Esto resuelve el caso de las marcaciones de salida a primera hora del día siguiente.
    for turno in posibles_turnos:
        if turno['es_turno_nocturno'] and \
           fecha_hora_registro.time() < time(7,0,0) and \
           turno['inicio_real'].date() == (fecha_hora_registro - timedelta(days=1)).date():
            # Si el registro está dentro de la tolerancia de la HORA DE FINALIZACIÓN del turno nocturno del día anterior
            if abs(fecha_hora_registro - turno['fin_real']) <= tolerancia:
                return (turno['nombre_turno'], turno['detalles'], turno['inicio_real'], turno['fin_real'])

    # 2. Si no se encontró un turno nocturno prioritario del día anterior,
    #    buscar entre los turnos donde el registro *cae dentro* de su rango de inicio a fin.
    #    De entre ellos, elegir el que tenga la menor diferencia absoluta al inicio del turno.
    mejores_candidatos_internos = []
    for turno in posibles_turnos:
        if turno['inicio_real'] <= fecha_hora_registro <= turno['fin_real']:
            mejores_candidatos_internos.append(turno)

    if mejores_candidatos_internos:
        # Entre los que caen dentro del rango, elegir el que tiene la menor diferencia absoluta al inicio del turno.
        best_candidate = min(mejores_candidatos_internos, key=lambda x: x['diferencia_a_inicio'])
        return (best_candidate['nombre_turno'], best_candidate['detalles'], 
                best_candidate['inicio_real'], best_candidate['fin_real'])

    # 3. Si el registro no cae dentro de ningún turno (pero está cerca con tolerancia),
    #    buscar el turno más cercano a cualquiera de sus límites (inicio o fin).
    #    Esto captura registros que están justo antes o justo después de un turno.
    best_candidate = None
    min_dist_to_boundary = timedelta.max
    
    for turno in posibles_turnos:
        dist_to_start = abs(fecha_hora_registro - turno['inicio_real'])
        dist_to_end = abs(fecha_hora_registro - turno['fin_real'])
        current_dist = min(dist_to_start, dist_to_end)
        
        if current_dist < min_dist_to_boundary:
            min_dist_to_boundary = current_dist
            best_candidate = turno
            
    if best_candidate:
        return (best_candidate['nombre_turno'], best_candidate['detalles'], 
                best_candidate['inicio_real'], best_candidate['fin_real'])
    
    return None, None, None, None


# --- 4. Función Principal para Calcular Horas Extras ---
def calcular_horas_extra(df_registros: pd.DataFrame, lugares_trabajo_normalizados: list, tolerancia_minutos: int):
    # Filtra los registros para incluir solo los lugares de trabajo principales y tipos de marcación válidos
    df_filtrado = df_registros[
        (df_registros['PORTERIA_NORMALIZED'].isin(lugares_trabajo_normalizados)) &
        (df_registros['PuntoMarcacion'].isin(['ent', 'sal']))
    ].sort_values(by=['COD_TRABAJADOR', 'FECHA_HORA_PROCESADA']) # Asegura orden cronológico por trabajador

    if df_filtrado.empty:
        return pd.DataFrame()

    # Paso crucial: Asignar a cada registro la fecha de inicio de su turno inferido.
    # Esto permite agrupar correctamente los turnos nocturnos que cruzan la medianoche,
    # ya que todos los registros de un mismo turno nocturno (entrada y salida) se asignarán a la misma fecha de inicio.
    df_filtrado['INFERRED_SHIFT_START_DATETIME'] = pd.NaT # Inicializar con Not a Time (datetime nulo de pandas)
    df_filtrado['INFERRED_SHIFT_END_DATETIME'] = pd.NaT
    df_filtrado['INFERRED_SHIFT_NAME'] = None
    df_filtrado['INFERRED_SHIFT_DURATION'] = None
    
    # Itera sobre cada registro para determinar su turno y la fecha de inicio inferida del mismo.
    # Esto es necesario para que cada registro individual "sepa" a qué turno lógico pertenece,
    # incluso si su fecha de registro original es diferente a la fecha de inicio del turno.
    for idx, row in df_filtrado.iterrows():
        nombre_turno, detalles_turno, inicio_turno_calculado, fin_turno_calculado = \
            obtener_turno_para_registro(row['FECHA_HORA_PROCESADA'], tolerancia_minutos)
        
        if nombre_turno is not None:
            df_filtrado.at[idx, 'INFERRED_SHIFT_START_DATETIME'] = inicio_turno_calculado
            df_filtrado.at[idx, 'INFERRED_SHIFT_END_DATETIME'] = fin_turno_calculado
            df_filtrado.at[idx, 'INFERRED_SHIFT_NAME'] = nombre_turno
            df_filtrado.at[idx, 'INFERRED_SHIFT_DURATION'] = detalles_turno['duracion_hrs']

    # Elimina registros que no pudieron ser asignados a ningún turno (ej: fuera de todo horario de trabajo conocido)
    df_filtrado = df_filtrado.dropna(subset=['INFERRED_SHIFT_START_DATETIME']).copy()
    
    if df_filtrado.empty:
        return pd.DataFrame()

    resultados = []

    # Ahora, agrupar por el trabajador y la fecha de inicio de turno inferida.
    # Esto asegura que todos los registros de un turno (incluyendo los nocturnos que abarcan dos días)
    # se procesen como una única jornada laboral lógica.
    for (codigo_trabajador, shift_start_dt), grupo in df_filtrado.groupby(['COD_TRABAJADOR', 'INFERRED_SHIFT_START_DATETIME']):
        
        nombre_trabajador = grupo['NOMBRE'].iloc[0]
        # La fecha base para el resultado será la fecha de inicio inferida del turno
        fecha_dia_base = shift_start_dt.date() 
        
        # Obtener la primera entrada y la última salida REALES registradas para este grupo de turno inferido
        primera_entrada_hora_real = grupo[grupo['PuntoMarcacion'] == 'ent']['FECHA_HORA_PROCESADA'].min()
        ultima_salida_hora_real = grupo[grupo['PuntoMarcacion'] == 'sal']['FECHA_HORA_PROCESADA'].max()

        # Si falta una entrada o salida para este turno inferido, se salta el cálculo
        if pd.isna(primera_entrada_hora_real) or pd.isna(ultima_salida_hora_real):
            continue

        # Obtener los detalles del turno programado (deberían ser consistentes dentro del grupo de turno inferido)
        shift_name = grupo['INFERRED_SHIFT_NAME'].iloc[0]
        shift_scheduled_start = grupo['INFERRED_SHIFT_START_DATETIME'].iloc[0]
        shift_scheduled_end = grupo['INFERRED_SHIFT_END_DATETIME'].iloc[0]
        shift_duration_hrs = grupo['INFERRED_SHIFT_DURATION'].iloc[0]
        
        # Calcular las horas trabajadas reales como la diferencia entre la última salida y la primera entrada.
        # Esto representa la duración total de la presencia del trabajador en este turno lógico.
        horas_trabajadas_td = ultima_salida_hora_real - primera_entrada_hora_real
        horas_trabajadas_hrs = horas_trabajadas_td.total_seconds() / 3600

        # Calcular horas extras: (horas trabajadas reales - duración estándar del turno programado)
        horas_extra = max(0, horas_trabajadas_hrs - shift_duration_hrs)

        if horas_extra < 0.5: # Umbral de 30 minutos para considerar horas extras
            horas_extra = 0.0

        resultados.append({
            'NOMBRE': nombre_trabajador,
            'COD_TRABAJADOR': codigo_trabajador,
            'FECHA': fecha_dia_base, # Fecha del inicio real del turno
            'Dia_Semana': fecha_dia_base.strftime('%A'),
            'TURNO': shift_name,
            'Inicio_Turno_Programado': shift_scheduled_start.strftime("%H:%M:%S"),
            'Fin_Turno_Programado': shift_scheduled_end.strftime("%H:%M:%S"),
            'Duracion_Turno_Programado_Hrs': shift_duration_hrs,
            'ENTRADA_AJUSTADA': primera_entrada_hora_real.strftime("%Y-%m-%d %H:%M:%S"), # La hora real de la primera entrada para este turno
            'SALIDA_REAL': ultima_salida_hora_real.strftime("%Y-%m-%d %H:%M:%S"), # La hora real de la última salida para este turno
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
            # Combina FECHA y HORA para crear un solo campo datetime para cada registro de marcación
            df_registros['FECHA_HORA_PROCESADA'] = pd.to_datetime(df_registros['FECHA'].dt.strftime('%Y-%m-%d') + ' ' + df_registros['HORA'])
            df_registros['PORTERIA_NORMALIZED'] = df_registros['PORTERIA'].astype(str).str.strip().str.lower()
            df_registros['PuntoMarcacion'] = df_registros['PuntoMarcacion'].astype(str).str.strip().str.lower().replace({'entrada': 'ent', 'salida': 'sal'})
            
            # --- Aplica el pre-procesamiento para corregir las fechas de salida de turnos nocturnos ---
            df_registros = adjust_misdated_night_shift_exits(df_registros)
            
            # Re-ordenar por si el ajuste de fechas cambió el orden cronológico general
            df_registros.sort_values(by=['COD_TRABAJADOR', 'FECHA_HORA_PROCESADA'], inplace=True)
            df_registros.reset_index(drop=True, inplace=True)

            st.success("Archivo cargado y pre-procesado con éxito.")

            # Ejecutar el cálculo
            st.subheader("Resultados del Cálculo")
            df_resultados_diarios = calcular_horas_extra(df_registros.copy(), LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS, TOLERANCIA_INFERENCIA_MINUTOS)

            if not df_resultados_diarios.empty:
                st.write("### Reporte Horas Extra Diarias (Solo registros con horas extra)")
                df_resultados_diarios_filtrado_extras = df_resultados_diarios[df_resultados_diarios['HORAS_EXTRA_HRS'] > 0].copy()

                if not df_resultados_diarios_filtrado_extras.empty:
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
                    st.info("No se encontraron horas extras diarias que superen el umbral de 30 minutos.")

                # --- Sección para ver TODOS los turnos diarios calculados (con y sin horas extra) ---
                if not df_resultados_diarios.empty:
                    st.markdown("---")
                    with st.expander("Ver todos los turnos diarios calculados (incluyendo sin horas extra)"):
                        st.write("Aquí puedes ver todos los turnos que se lograron inferir para cada trabajador, incluso si no generaron horas extra.")
                        st.dataframe(df_resultados_diarios)
                        
                        excel_buffer_all_daily = io.BytesIO()
                        df_resultados_diarios.to_excel(excel_buffer_all_daily, index=False, engine='openpyxl')
                        excel_buffer_all_daily.seek(0)
                        st.download_button(
                            label="Descargar Todos los Turnos Diarios Calculados (Excel)",
                            data=excel_buffer_all_daily,
                            file_name="todos_los_turnos_diarios_calculados.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                else:
                    st.info("No se pudieron calcular turnos diarios para ningún registro válido después del pre-procesamiento.")


                # Generar Resumen Semanal
                st.markdown("---")
                st.write("### Resumen Horas Extra Semanal")
                df_resumen_semanal = pd.DataFrame()
                # La columna 'FECHA' en df_resultados_diarios ya es la fecha de inicio inferida del turno,
                # lo cual es correcto para agrupar por semana.
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
                st.warning("No se pudieron calcular horas extras. Asegúrate de que el archivo Excel tenga los datos y formatos correctos y que haya registros válidos de entrada/salida en lugares de trabajo principales.")

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}. Asegúrate de que el archivo es un Excel válido y la hoja 'BaseDatos Modificada' existe.")

st.markdown("---")
st.caption("Somos NOEL DE CORAZÓN ❤️ ")
