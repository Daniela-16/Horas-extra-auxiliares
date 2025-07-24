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
TOLERANCIA_INFERENCIA_MINUTOS = 30
JORNADA_SEMANAL_ESTANDAR = timedelta(hours=46)

# --- 3. Función para determinar el turno y sus horas de inicio/fin ajustadas y la fecha efectiva del turno ---
def obtener_turno_para_registro(fecha_hora_registro: datetime, tolerancia_minutos: int):
    """
    Determina el turno al que corresponde un registro dado,
    ajustando las horas de inicio y fin del turno para que reflejen
    correctamente los turnos nocturnos que cruzan la medianoche.
    """
    dia_de_semana = fecha_hora_registro.weekday()
    tipo_dia = "LV" if dia_de_semana < 5 else "SAB"

    mejor_turno_info = None # Formato: (nombre_turno, detalles_turno, inicio_turno_efectivo, fin_turno_efectivo)
    min_diferencia_tiempo = timedelta(days=999)
    fecha_turno_asociada = None # La fecha de inicio del turno lógico (para agrupar)

    for nombre_turno, detalles_turno in TURNOS[tipo_dia].items():
        hora_inicio_turno_obj = datetime.strptime(detalles_turno["inicio"], "%H:%M:%S").time()
        hora_fin_turno_obj = datetime.strptime(detalles_turno["fin"], "%H:%M:%S").time()

        # Generar candidatos para el inicio y fin del turno basados en la fecha del registro
        # Candidato 1: El turno comienza el mismo día del registro
        candidato_inicio_1 = fecha_hora_registro.replace(
            hour=hora_inicio_turno_obj.hour,
            minute=hora_inicio_turno_obj.minute,
            second=hora_inicio_turno_obj.second
        )
        candidato_fin_1 = candidato_inicio_1.replace(
            hour=hora_fin_turno_obj.hour,
            minute=hora_fin_turno_obj.minute,
            second=hora_fin_turno_obj.second
        )
        if hora_inicio_turno_obj > hora_fin_turno_obj: # Si es turno nocturno (inicio > fin), la salida es al día siguiente
            candidato_fin_1 += timedelta(days=1)

        # Candidato 2: Si es un turno nocturno, el registro actual podría pertenecer a un turno que comenzó el día anterior
        candidato_inicio_2 = None
        candidato_fin_2 = None
        if hora_inicio_turno_obj > hora_fin_turno_obj: # Solo relevante para turnos nocturnos
            candidato_inicio_2 = (fecha_hora_registro - timedelta(days=1)).replace( # El turno empezó el día anterior
                hour=hora_inicio_turno_obj.hour,
                minute=hora_inicio_turno_obj.minute,
                second=hora_inicio_turno_obj.second
            )
            candidato_fin_2 = candidato_inicio_2.replace(
                hour=hora_fin_turno_obj.hour,
                minute=hora_fin_turno_obj.minute,
                second=hora_fin_turno_obj.second
            )
            candidato_fin_2 += timedelta(days=1) # La salida de este turno también sería al día siguiente del inicio

        possible_shifts = []
        possible_shifts.append((candidato_inicio_1, candidato_fin_1))
        if candidato_inicio_2:
            possible_shifts.append((candidato_inicio_2, candidato_fin_2))

        # Iterar sobre los candidatos a turno para encontrar el que mejor se ajusta
        for inicio_actual, fin_actual in possible_shifts:
            # Verifica si el registro cae dentro de la ventana de tolerancia de este turno
            if (inicio_actual - timedelta(minutes=tolerancia_minutos) <= fecha_hora_registro <= fin_actual + timedelta(minutes=tolerancia_minutos)):
                # Calcula la diferencia de tiempo entre el registro y el inicio del turno candidato
                # para encontrar el más cercano.
                diferencia_tiempo = abs(fecha_hora_registro - inicio_actual)

                if mejor_turno_info is None or diferencia_tiempo < min_diferencia_tiempo:
                    mejor_turno_info = (nombre_turno, detalles_turno, inicio_actual, fin_actual)
                    min_diferencia_tiempo = diferencia_tiempo
                    # La fecha de turno asociada es siempre la fecha de inicio del turno programado
                    fecha_turno_asociada = inicio_actual.date()

    if mejor_turno_info:
        # Retorna el nombre del turno, sus detalles, las fechas/horas programadas ajustadas, y la fecha efectiva del turno
        return mejor_turno_info[0], mejor_turno_info[1], mejor_turno_info[2], mejor_turno_info[3], fecha_turno_asociada
    else:
        # Si no se encontró ningún turno que coincida, se devuelve None
        return None, None, None, None, None

# --- 4. Función Principal para Calcular Horas Extras ---
def calcular_horas_extra(df_registros: pd.DataFrame, lugares_trabajo_normalizados: list, tolerancia_minutos: int):
    """
    Calcula las horas extras diarias y semanales a partir de un DataFrame de registros de marcaciones,
    utilizando la lógica de turnos y corrigiendo las fechas para turnos nocturnos.

    Args:
        df_registros (pd.DataFrame): DataFrame con los registros de marcaciones ya pre-procesados
                                     (con FECHA_HORA_PROCESADA corregida).
        lugares_trabajo_normalizados (list): Lista de lugares de trabajo válidos normalizados.
        tolerancia_minutos (int): Tolerancia en minutos para la inferencia de turnos.

    Returns:
        pd.DataFrame: DataFrame con los resultados de horas extras diarias.
    """
    # El filtrado por portería y tipo de marcación ya se realizó en el pre-procesamiento
    df_filtrado = df_registros.sort_values(by=['COD_TRABAJADOR', 'FECHA_HORA_PROCESADA'])

    if df_filtrado.empty:
        return pd.DataFrame()

    # Aplica la función de detección de turno a cada registro para obtener su fecha de turno efectiva
    # y los tiempos programados correspondientes.
    df_filtrado['NOMBRE_TURNO_DETECTADO'] = None
    df_filtrado['DETALLES_TURNO_DETECTADO'] = None
    df_filtrado['INICIO_TURNO_PROGRAMADO_DETECTADO'] = pd.NaT
    df_filtrado['FIN_TURNO_PROGRAMADO_DETECTADO'] = pd.NaT
    df_filtrado['FECHA_TURNO_EFECTIVA'] = pd.NaT # Esta es la columna clave para agrupar lógicamente los turnos

    for index, row in df_filtrado.iterrows():
        nombre_turno, detalles_turno, inicio_prog, fin_prog, fecha_asociada = \
            obtener_turno_para_registro(row['FECHA_HORA_PROCESADA'], tolerancia_minutos)

        if nombre_turno:
            df_filtrado.at[index, 'NOMBRE_TURNO_DETECTADO'] = nombre_turno
            df_filtrado.at[index, 'DETALLES_TURNO_DETECTADO'] = detalles_turno
            df_filtrado.at[index, 'INICIO_TURNO_PROGRAMADO_DETECTADO'] = inicio_prog
            df_filtrado.at[index, 'FIN_TURNO_PROGRAMADO_DETECTADO'] = fin_prog
            df_filtrado.at[index, 'FECHA_TURNO_EFECTIVA'] = fecha_asociada

    # Elimina registros que no pudieron ser asociados con ningún turno válido
    df_filtrado.dropna(subset=['FECHA_TURNO_EFECTIVA'], inplace=True)

    if df_filtrado.empty:
        return pd.DataFrame()

    resultados = []

    # Agrupa por trabajador y la nueva "fecha de turno efectiva" para consolidar los registros de cada turno lógico.
    for (codigo_trabajador, fecha_turno_efectiva), grupo in df_filtrado.groupby(['COD_TRABAJADOR', 'FECHA_TURNO_EFECTIVA']):
        nombre_trabajador = grupo['NOMBRE'].iloc[0]
        entradas = grupo[grupo['PuntoMarcacion'] == 'ent']
        salidas = grupo[grupo['PuntoMarcacion'] == 'sal']

        if entradas.empty or salidas.empty:
            continue

        primera_entrada_hora_real = entradas['FECHA_HORA_PROCESADA'].min()
        ultima_salida_hora_real = salidas['FECHA_HORA_PROCESADA'].max()

        # Vuelve a determinar el turno y sus límites programados usando la *primera entrada real* del grupo.
        # Esto es crucial para obtener el 'inicio_turno_programado' correcto para el cálculo del periodo.
        nombre_turno, detalles_turno, inicio_turno_programado, fin_turno_programado, _ = \
             obtener_turno_para_registro(primera_entrada_hora_real, tolerancia_minutos)

        if nombre_turno is None:
            # Si la primera entrada del grupo no coincide con un turno, se salta (caso excepcional, ya filtrado en teoría)
            continue

        # Evita cálculos si la última salida real es anterior o igual a la primera entrada real
        if ultima_salida_hora_real <= primera_entrada_hora_real:
            continue

        # Calcula las horas trabajadas desde el inicio programado del turno hasta la salida real
        horas_trabajadas_td = ultima_salida_hora_real - inicio_turno_programado
        horas_trabajadas_hrs = horas_trabajadas_td.total_seconds() / 3600

        duracion_estandar_hrs = detalles_turno["duracion_hrs"]
        horas_extra = max(0, horas_trabajadas_hrs - duracion_estandar_hrs)

        if horas_extra < 0.5: # Umbral de 30 minutos (menos de 30 minutos no se considera extra)
            horas_extra = 0.0

        resultados.append({
            'NOMBRE': nombre_trabajador,
            'COD_TRABAJADOR': codigo_trabajador,
            'FECHA_TURNO_EFECTIVA': fecha_turno_efectiva, # La fecha de inicio efectiva del turno
            'Dia_Semana': fecha_turno_efectiva.strftime('%A'),
            'TURNO': nombre_turno,
            'Inicio_Turno_Programado': inicio_turno_programado.strftime("%Y-%m-%d %H:%M:%S"), # Incluye fecha para turnos nocturnos
            'Fin_Turno_Programado': fin_turno_programado.strftime("%Y-%m-%d %H:%M:%S"),     # Incluye fecha para turnos nocturnos
            'Duracion_Turno_Programado_Hrs': duracion_estandar_hrs,
            'ENTRADA_REAL': primera_entrada_hora_real.strftime("%Y-%m-%d %H:%M:%S"), # Hora real de entrada
            'SALIDA_REAL': ultima_salida_hora_real.strftime("%Y-%m-%d %H:%M:%S"),   # Hora real de salida
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
            # NORMALIZACIÓN Y PRE-PROCESAMIENTO INICIAL DE DATOS CRUDOS
            df_registros['FECHA'] = pd.to_datetime(df_registros['FECHA'])
            df_registros['HORA_STR'] = df_registros['HORA'].astype(str) # Se mantiene la hora como string para combinar
            df_registros['PuntoMarcacion'] = df_registros['PuntoMarcacion'].astype(str).str.strip().str.lower().replace({'entrada': 'ent', 'salida': 'sal'})
            df_registros['PORTERIA_NORMALIZED'] = df_registros['PORTERIA'].astype(str).str.strip().str.lower()

            # Filtrar registros válidos por portería y tipo de marcación antes de cualquier lógica compleja.
            df_procesar = df_registros[
                (df_registros['PORTERIA_NORMALIZED'].isin(LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS)) &
                (df_registros['PuntoMarcacion'].isin(['ent', 'sal']))
            ].copy()

            if df_procesar.empty:
                st.warning("No se encontraron registros válidos para procesar después de filtrar por portería y tipo de marcación.")
            else:
                # Combinar FECHA y HORA en un datetime para obtener un timestamp "crudo" tal como está en el Excel.
                df_procesar['FECHA_HORA_RAW'] = df_procesar.apply(
                    lambda row: datetime.combine(row['FECHA'].date(), datetime.strptime(row['HORA_STR'], "%H:%M:%S").time()), axis=1
                )

                # --- Lógica de Emparejamiento y Corrección de Fecha para Turnos Nocturnos ---
                # Este es el paso crucial para corregir las fechas de salida de turnos nocturnos
                # cuando la base de datos original registra entrada y salida en el mismo día.
                processed_records_list = [] # Lista para acumular los registros con FECHA_HORA_PROCESADA corregida

                # Agrupar por trabajador para procesar sus marcaciones secuencialmente
                for cod_trabajador, group_df in df_procesar.groupby('COD_TRABAJADOR'):
                    # Ordenar los registros del empleado por el timestamp crudo para asegurar el orden cronológico
                    group_sorted = group_df.sort_values(by='FECHA_HORA_RAW').reset_index(drop=True)

                    pending_entries = [] # Pila para mantener las entradas no emparejadas para este trabajador

                    for idx in range(len(group_sorted)):
                        current_rec = group_sorted.loc[idx].copy() # Usar .copy() para evitar SettingWithCopyWarning

                        if current_rec['PuntoMarcacion'] == 'ent':
                            pending_entries.append(current_rec) # Añadir la entrada a la pila
                        elif current_rec['PuntoMarcacion'] == 'sal':
                            # Intentar emparejar esta salida con la entrada pendiente más reciente
                            if pending_entries:
                                # Tomar la última entrada de la pila (la más reciente)
                                prev_ent_rec = pending_entries.pop()

                                entrada_dt_raw = prev_ent_rec['FECHA_HORA_RAW']
                                salida_dt_raw = current_rec['FECHA_HORA_RAW']
                                salida_dt_corrected = salida_dt_raw # Por defecto, la salida corregida es la cruda

                                # CONDICIÓN CLAVE PARA DETECTAR Y CORREGIR TURNOS NOCTURNOS MAL REGISTRADOS
                                # Si la entrada y la salida están en la misma fecha RAW
                                # Y la entrada es tarde (ej. después de las 8 PM)
                                # Y la salida es temprano (ej. antes de las 8 AM)
                                if (entrada_dt_raw.date() == salida_dt_raw.date() and
                                    entrada_dt_raw.time() >= time(20, 0, 0) and # Entrada a partir de las 8 PM
                                    salida_dt_raw.time() < time(8, 0, 0)):     # Salida antes de las 8 AM
                                    # Esto indica que es un turno nocturno donde la fecha de salida necesita corrección
                                    salida_dt_corrected = salida_dt_raw + timedelta(days=1)

                                # Asignar los timestamps procesados (potencialmente corregidos) a las nuevas columnas
                                prev_ent_rec['FECHA_HORA_PROCESADA'] = entrada_dt_raw
                                current_rec['FECHA_HORA_PROCESADA'] = salida_dt_corrected

                                processed_records_list.append(prev_ent_rec)
                                processed_records_list.append(current_rec)
                            else:
                                # Si no hay entradas pendientes para esta salida, mantener su timestamp raw
                                current_rec['FECHA_HORA_PROCESADA'] = current_rec['FECHA_HORA_RAW']
                                processed_records_list.append(current_rec)
                        # Cualquier otro tipo de marcación (que debería ser filtrada previamente, pero por seguridad)
                        else:
                            current_rec['FECHA_HORA_PROCESADA'] = current_rec['FECHA_HORA_RAW']
                            processed_records_list.append(current_rec)

                    # Añadir cualquier entrada que haya quedado sin emparejar al final de la jornada del trabajador
                    for rec in pending_entries:
                        rec['FECHA_HORA_PROCESADA'] = rec['FECHA_HORA_RAW']
                        processed_records_list.append(rec)

                # Convertir la lista de registros procesados a un DataFrame
                df_final = pd.DataFrame(processed_records_list)

                if df_final.empty:
                    st.warning("No se pudieron procesar registros de entrada/salida válidos después de la corrección de fechas.")
                    st.stop() # Detener la ejecución si no hay datos finales

                # Re-ordenar el DataFrame final por trabajador y por la nueva FECHA_HORA_PROCESADA corregida.
                df_final.sort_values(by=['COD_TRABAJADOR', 'FECHA_HORA_PROCESADA'], inplace=True)
                df_final.reset_index(drop=True, inplace=True)

                # Eliminar columnas temporales que ya no se necesitan
                df_final = df_final.drop(columns=['FECHA_HORA_RAW', 'FECHA', 'HORA_STR'], errors='ignore')

                st.success("Archivo cargado y pre-procesado con éxito. Fechas de turnos nocturnos ajustadas.")

                # --- Debugging: Mostrar cómo se ven los registros para el trabajador 71329 después de la corrección ---
                # Esto es para ayudar a verificar visualmente que las fechas de salida de los turnos nocturnos se han corregido
                # Puedes quitar esta sección una vez que confirmes que funciona correctamente.
                if 71329 in df_procesar['COD_TRABAJADOR'].unique():
                    st.write(f"### Vista de Registros Procesados para el Trabajador 71329 (DEBUG)")
                    debug_df = df_final[df_final['COD_TRABAJADOR'] == 71329].copy()
                    if not debug_df.empty:
                        # Muestra las columnas relevantes para la depuración
                        st.dataframe(debug_df[['COD_TRABAJADOR', 'FECHA_HORA_PROCESADA', 'PuntoMarcacion']])
                    else:
                        st.info("No hay registros para el trabajador 71329 en el conjunto de datos final.")
                # --- Fin Debugging ---

                # Ejecutar el cálculo de horas extra con el DataFrame ya corregido
                st.subheader("Resultados del Cálculo")
                df_resultados_diarios = calcular_horas_extra(df_final.copy(), LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS, TOLERANCIA_INFERENCIA_MINUTOS)

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
                    # Asegúrate de usar 'FECHA_TURNO_EFECTIVA' para el cálculo de la semana
                    df_resultados_diarios['Semana_Inicio'] = df_resultados_diarios['FECHA_TURNO_EFECTIVA'].apply(lambda x: x - timedelta(days=x.weekday()))

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
