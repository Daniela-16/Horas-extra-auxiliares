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
# Los turnos se definen con horas de inicio y fin.
# Para el turno nocturno (Turno 3), la hora de fin es menor que la de inicio,
# lo que indica que cruza la medianoche.
TURNOS = {
    "LV": { # Lunes a Viernes
        "Turno 1 LV": {"inicio": "05:40:00", "fin": "13:40:00", "duracion_hrs": 8},
        "Turno 2 LV": {"inicio": "13:40:00", "fin": "21:40:00", "duracion_hrs": 8},
        "Turno 3 LV": {"inicio": "21:40:00", "fin": "05:40:00", "duracion_hrs": 8}, # Cruza la medianoche
    },
    "SAB": { # Sábados
        "Turno 1 SAB": {"inicio": "05:40:00", "fin": "11:40:00", "duracion_hrs": 6},
        "Turno 2 SAB": {"inicio": "11:40:00", "fin": "17:40:00", "duracion_hrs": 6},
        "Turno 3 SAB": {"inicio": "21:40:00", "fin": "05:40:00", "duracion_hrs": 8}, # Cruza la medianoche
    }
}

# --- 2. Configuración General ---
# Lista de lugares de trabajo principales para filtrar los registros.
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
TOLERANCIA_INFERENCIA_MINUTOS = 30 # Tolerancia en minutos para asociar un registro a un turno
JORNADA_SEMANAL_ESTANDAR = timedelta(hours=46) # Jornada semanal estándar para cálculo de horas extra semanales

# --- 3. Función para determinar el turno y sus horas de inicio/fin ajustadas ---
# Esta función toma un registro de fecha y hora (entrada o salida) y determina
# a qué turno programado pertenece, considerando turnos diurnos y nocturnos.
def obtener_turno_para_registro(fecha_hora_registro: datetime, tolerancia_minutos: int):
    # Determina si es día de semana o sábado para aplicar la configuración de turnos correcta.
    dia_de_semana = fecha_hora_registro.weekday()
    tipo_dia = "LV" if dia_de_semana < 5 else "SAB"

    best_match_info = None # Almacenará (nombre_turno, detalles_turno, inicio_turno_calculado, fin_turno_calculado, fecha_inicio_turno_efectiva)
    min_diferencia_para_registro = timedelta(days=999) # La diferencia más pequeña entre el registro y el inicio programado del turno.

    # Itera sobre todos los turnos definidos para el tipo de día (LV o SAB).
    for nombre_turno, detalles_turno in TURNOS[tipo_dia].items():
        hora_inicio_turno_obj = datetime.strptime(detalles_turno["inicio"], "%H:%M:%S").time()
        hora_fin_turno_obj = datetime.strptime(detalles_turno["fin"], "%H:%M:%S").time()

        # Genera posibles fechas y horas de inicio para el turno, relativas a la fecha del registro.
        # Un registro puede pertenecer a un turno que empezó el mismo día o el día anterior (para turnos nocturnos).
        possible_shift_starts = []

        # Opción 1: El turno comienza el mismo día del registro.
        possible_shift_starts.append(
            datetime.combine(fecha_hora_registro.date(), hora_inicio_turno_obj)
        )

        # Opción 2: El turno comenzó el día anterior (crucial para turnos nocturnos y salidas de madrugada).
        possible_shift_starts.append(
            datetime.combine((fecha_hora_registro - timedelta(days=1)).date(), hora_inicio_turno_obj)
        )

        for shift_start_dt_candidate in possible_shift_starts:
            # Calcula la fecha y hora de fin para este turno candidato.
            shift_end_dt_candidate = datetime.combine(shift_start_dt_candidate.date(), hora_fin_turno_obj)

            # Ajusta la fecha de fin si el turno cruza la medianoche.
            # Esto ocurre si la hora de inicio es posterior a la hora de fin (ej. 21:40 -> 05:40).
            if hora_inicio_turno_obj > hora_fin_turno_obj:
                shift_end_dt_candidate += timedelta(days=1)
            # También para robustez si por alguna definición el fin fuera antes del inicio en un mismo día.
            elif shift_end_dt_candidate < shift_start_dt_candidate:
                 shift_end_dt_candidate += timedelta(days=1)

            # Define la ventana de tolerancia alrededor del turno.
            # Un registro se considera parte de un turno si cae dentro de esta ventana:
            # (inicio del turno - tolerancia) <= registro <= (fin del turno + tolerancia).
            window_start_dt = shift_start_dt_candidate - timedelta(minutes=tolerancia_minutos)
            window_end_dt = shift_end_dt_candidate + timedelta(minutes=tolerancia_minutos)

            # Si el registro actual está dentro de esta ventana, este turno es un candidato.
            if window_start_dt <= fecha_hora_registro <= window_end_dt:
                # Calcula la diferencia absoluta entre el registro y el inicio programado del turno.
                # Se busca el turno cuyo inicio programado sea el más cercano al registro,
                # para priorizar la asignación si hay solapamientos por la tolerancia.
                current_diff = abs(fecha_hora_registro - shift_start_dt_candidate)

                # Si es el primer turno encontrado o este es más cercano al registro, lo selecciona como el mejor.
                if best_match_info is None or current_diff < min_diferencia_para_registro:
                    best_match_info = (nombre_turno, detalles_turno, shift_start_dt_candidate, shift_end_dt_candidate, shift_start_dt_candidate.date())
                    min_diferencia_para_registro = current_diff

    # Retorna la información del mejor turno encontrado o valores nulos si no se encontró ninguno.
    return best_match_info if best_match_info else (None, None, None, None, None)

# --- 4. Función Principal para Calcular Horas Extras ---
# Esta función procesa el DataFrame de registros y calcula las horas extra,
# emparejando entradas y salidas de forma cronológica por empleado.
def calcular_horas_extra(df_registros: pd.DataFrame, lugares_trabajo_normalizados: list, tolerancia_minutos: int):
    # Filtra los registros por lugares de trabajo principales y por tipo de marcación (entrada/salida).
    df_filtrado = df_registros[
        (df_registros['PORTERIA_NORMALIZED'].isin(lugares_trabajo_normalizados)) &
        (df_registros['PuntoMarcacion'].isin(['ent', 'sal']))
    ].sort_values(by=['COD_TRABAJADOR', 'FECHA_HORA_PROCESADA']) # Ordena por trabajador y cronológicamente.

    if df_filtrado.empty:
        return pd.DataFrame()

    resultados = [] # Lista para almacenar los resultados de las horas extra calculadas.

    # Procesa los registros de cada empleado de forma individual.
    for codigo_trabajador, df_empleado in df_filtrado.groupby('COD_TRABAJADOR'):
        nombre_trabajador = df_empleado['NOMBRE'].iloc[0]
        # `unmatched_entries` almacena las entradas que aún no han sido emparejadas con una salida.
        # Formato: (fecha_hora_entrada_real, fecha_inicio_turno_efectiva, nombre_turno, detalles_turno, inicio_turno_calculado, fin_turno_calculado)
        unmatched_entries = []

        # Itera sobre cada registro (entrada o salida) del empleado en orden cronológico.
        for _, row in df_empleado.iterrows():
            punto_marcacion = row['PuntoMarcacion']
            fecha_hora_actual = row['FECHA_HORA_PROCESADA']

            if punto_marcacion == 'ent':
                # Para un registro de entrada, se determina a qué turno pertenece.
                nombre_turno, detalles_turno, inicio_turno_calculado, fin_turno_calculado, effective_shift_date = \
                    obtener_turno_para_registro(fecha_hora_actual, tolerancia_minutos)

                if nombre_turno is not None:
                    # Si se encuentra un turno, se añade la entrada a la lista de no emparejadas.
                    unmatched_entries.append((fecha_hora_actual, effective_shift_date, nombre_turno, detalles_turno, inicio_turno_calculado, fin_turno_calculado))
                # Se asegura que las entradas no emparejadas estén siempre ordenadas por su hora real,
                # para emparejar con la entrada más antigua disponible.
                unmatched_entries.sort(key=lambda x: x[0])

            elif punto_marcacion == 'sal':
                best_entry_idx = -1
                # Se usa una diferencia de tiempo muy grande para inicializar,
                # que se reducirá a medida que se encuentre la mejor entrada.
                best_entry_time_diff_to_shift_start = timedelta(days=999)

                # Busca una entrada adecuada para emparejar con esta salida de la lista de `unmatched_entries`.
                for i, (entry_time, entry_shift_date, entry_shift_name, entry_shift_details, entry_shift_start_calc, entry_shift_end_calc) in enumerate(unmatched_entries):
                    # La entrada debe ser anterior a la salida.
                    if entry_time >= fecha_hora_actual:
                        continue

                    # La duración entre la entrada y la salida debe ser razonable para un turno.
                    # Por ejemplo, no más de 16 horas y no menos de 5 minutos para evitar marcaciones erróneas.
                    time_diff_actual = fecha_hora_actual - entry_time
                    if time_diff_actual > timedelta(hours=16) or time_diff_actual < timedelta(minutes=5):
                        continue

                    # Validación crucial: La salida debe caer dentro del período de tolerancia del turno asociado a la entrada.
                    # Esto es fundamental para asegurar que la entrada y la salida pertenecen al mismo turno lógico.
                    exit_tolerance_window_start = entry_shift_start_calc - timedelta(minutes=tolerancia_minutos)
                    exit_tolerance_window_end = entry_shift_end_calc + timedelta(minutes=tolerancia_minutos)

                    if not (exit_tolerance_window_start <= fecha_hora_actual <= exit_tolerance_window_end):
                         continue

                    # Si múltiples entradas son candidatas para esta salida, se prefiere aquella cuya hora de entrada real
                    # sea más cercana a la hora de inicio programada del turno que le fue asignado.
                    current_entry_time_diff = abs(entry_time - entry_shift_start_calc)

                    if best_entry_idx == -1 or current_entry_time_diff < best_entry_time_diff_to_shift_start:
                        best_entry_time_diff_to_shift_start = current_entry_time_diff
                        best_entry_idx = i

                # Si se encontró un par de entrada-salida adecuado.
                if best_entry_idx != -1:
                    # Se extrae la información de la entrada emparejada y se remueve de la lista de no emparejadas.
                    entry_time, effective_shift_date, nombre_turno, detalles_turno, inicio_turno_calculado, fin_turno_calculado = unmatched_entries.pop(best_entry_idx)

                    # Calcula las horas trabajadas: desde el inicio *programado* del turno hasta la salida *real*.
                    horas_trabajadas_td = fecha_hora_actual - inicio_turno_calculado
                    horas_trabajadas_hrs = horas_trabajadas_td.total_seconds() / 3600

                    duracion_estandar_hrs = detalles_turno["duracion_hrs"]
                    horas_extra = max(0, horas_trabajadas_hrs - duracion_estandar_hrs) # Asegura que las horas extra no sean negativas.

                    # Aplica un umbral de 30 minutos (0.5 horas) para considerar horas extra.
                    # Si es menos de 30 minutos, se considera 0 horas extra.
                    if horas_extra < 0.5:
                        horas_extra = 0.0

                    # Añade los resultados al listado final.
                    resultados.append({
                        'NOMBRE': nombre_trabajador,
                        'COD_TRABAJADOR': codigo_trabajador,
                        'FECHA': effective_shift_date, # La fecha efectiva del turno (el día en que el turno comenzó).
                        'Dia_Semana': effective_shift_date.strftime('%A'),
                        'TURNO': nombre_turno,
                        'Inicio_Turno_Programado': inicio_turno_calculado.strftime("%H:%M:%S"),
                        'Fin_Turno_Programado': fin_turno_calculado.strftime("%H:%M:%S"),
                        'Duracion_Turno_Programado_Hrs': duracion_estandar_hrs,
                        'ENTRADA_AJUSTADA': inicio_turno_calculado.strftime("%Y-%m-%d %H:%M:%S"), # Inicio del turno programado con su fecha correcta.
                        'SALIDA_REAL': fecha_hora_actual.strftime("%Y-%m-%d %H:%M:%S"), # La hora de salida real del empleado.
                        'HORAS_TRABAJADAS_CALCULADAS_HRS': round(horas_trabajadas_hrs, 2),
                        'HORAS_EXTRA_HRS': round(horas_extra, 2),
                        'HORAS_EXTRA_ENTERAS_HRS': int(horas_extra),
                        'MINUTOS_EXTRA_CONVERTIDOS': round((horas_extra - int(horas_extra)) * 60, 2)
                    })
    # Retorna un DataFrame con todos los resultados calculados.
    return pd.DataFrame(resultados)

# --- Interfaz de usuario de Streamlit ---
# Configuración de la página de Streamlit.
st.set_page_config(page_title="Calculadora de Horas Extra", layout="wide")
st.title("📊 Calculadora de Horas Extra")
st.write("Sube tu archivo de Excel para calcular las horas extra de tus trabajadores.")

# Componente para subir el archivo de Excel.
uploaded_file = st.file_uploader("Sube un archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Lee el archivo de Excel. Se espera que la hoja de datos se llame 'BaseDatos Modificada'.
        df_registros = pd.read_excel(uploaded_file, sheet_name='BaseDatos Modificada')

        # Verifica que todas las columnas requeridas existan en el archivo cargado.
        columnas_requeridas = ['COD_TRABAJADOR', 'APUNTADOR', 'DESC_APUNTADOR', 'NOMBRE', 'FECHA', 'HORA', 'PORTERIA', 'PuntoMarcacion']
        if not all(col in df_registros.columns for col in columnas_requeridas):
            st.error(f"ERROR: Faltan columnas requeridas en la hoja 'BaseDatos Modificada'. Asegúrate de que existan: {', '.join(columnas_requeridas)}")
        else:
            # Preprocesamiento de los datos:
            # Convierte la columna 'FECHA' a formato datetime, asumiendo formato día/mes/año.
            df_registros['FECHA'] = pd.to_datetime(df_registros['FECHA'], dayfirst=True)
            # Convierte la columna 'HORA' a string y luego combina con 'FECHA' para crear 'FECHA_HORA_PROCESADA'.
            # Se usa errors='coerce' para convertir a NaT (Not a Time) cualquier valor que no pueda ser parseado,
            # lo que ayuda a evitar errores y a limpiar los datos.
            df_registros['HORA'] = df_registros['HORA'].astype(str)
            df_registros['FECHA_HORA_PROCESADA'] = pd.to_datetime(df_registros['FECHA'].dt.strftime('%Y-%m-%d') + ' ' + df_registros['HORA'], errors='coerce')

            # Manejo de errores: Si hay valores nulos después de la conversión de FECHA_HORA_PROCESADA,
            # significa que algunas fechas/horas no eran válidas y se eliminan.
            if df_registros['FECHA_HORA_PROCESADA'].isnull().any():
                st.warning("Algunas combinaciones de FECHA y HORA no pudieron ser procesadas y serán ignoradas. Por favor, revisa el formato de tus datos.")
                df_registros.dropna(subset=['FECHA_HORA_PROCESADA'], inplace=True)

            # Normaliza las columnas 'PORTERIA' y 'PuntoMarcacion' a minúsculas y sin espacios extra,
            # y reemplaza 'entrada'/'salida' con 'ent'/'sal'.
            df_registros['PORTERIA_NORMALIZED'] = df_registros['PORTERIA'].astype(str).str.strip().str.lower()
            df_registros['PuntoMarcacion'] = df_registros['PuntoMarcacion'].astype(str).str.strip().str.lower().replace({'entrada': 'ent', 'salida': 'sal'})
            # Ordena los registros por trabajador y cronológicamente para el procesamiento.
            df_registros.sort_values(by=['COD_TRABAJADOR', 'FECHA_HORA_PROCESADA'], inplace=True)
            df_registros.reset_index(drop=True, inplace=True)

            st.success("Archivo cargado y pre-procesado con éxito.")

            # Ejecuta la función principal para calcular las horas extra.
            st.subheader("Resultados del Cálculo")
            df_resultados_diarios = calcular_horas_extra(df_registros.copy(), LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS, TOLERANCIA_INFERENCIA_MINUTOS)

            if not df_resultados_diarios.empty:
                # Muestra el reporte de horas extra diarias (solo donde hay horas extra).
                df_resultados_diarios_filtrado_extras = df_resultados_diarios[df_resultados_diarios['HORAS_EXTRA_HRS'] > 0].copy()

                if not df_resultados_diarios_filtrado_extras.empty:
                    st.write("### Reporte Horas Extra Diarias")
                    st.dataframe(df_resultados_diarios_filtrado_extras)

                    # Permite al usuario descargar el reporte diario como un archivo Excel.
                    excel_buffer_diario = io.BytesIO()
                    df_resultados_diarios_filtrado_extras.to_excel(excel_buffer_diario, index=False, engine='openpyxl')
                    excel_buffer_diario.seek(0) # Regresar al inicio del buffer para la descarga.

                    st.download_button(
                        label="Descargar Reporte Horas Extra Diarias (Excel)",
                        data=excel_buffer_diario,
                        file_name="reporte_horas_extra_diarias.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.info("No se encontraron horas extras diarias para reportar.")

                # Generar Resumen Semanal de horas trabajadas para calcular extras semanales.
                # Asegura que 'FECHA' es datetime para calcular el inicio de la semana.
                df_resultados_diarios['FECHA'] = pd.to_datetime(df_resultados_diarios['FECHA'])
                # Calcula el inicio de la semana para cada registro (lunes de esa semana).
                df_resultados_diarios['Semana_Inicio'] = df_resultados_diarios['FECHA'].apply(lambda x: x - timedelta(days=x.weekday()))

                # Agrupa por trabajador y semana para sumar las horas trabajadas.
                df_resumen_semanal = df_resultados_diarios.groupby(['COD_TRABAJADOR', 'NOMBRE', 'Semana_Inicio']).agg(
                    Horas_Trabajadas_Calculadas_Semana_Hrs=('HORAS_TRABAJADAS_CALCULADAS_HRS', 'sum')
                ).reset_index()

                df_resumen_semanal['Semana_Inicio'] = pd.to_datetime(df_resumen_semanal['Semana_Inicio'])

                # Calcula las horas extra semanales restando la jornada estándar.
                df_resumen_semanal['Horas_Extra_Semanales_Hrs'] = (
                    df_resumen_semanal['Horas_Trabajadas_Calculadas_Semana_Hrs'] - (JORNADA_SEMANAL_ESTANDAR.total_seconds() / 3600)
                ).apply(lambda x: max(0, round(x, 2))) # Asegura que las horas extra no sean negativas y redondea.

                # Filtra y muestra solo los registros con horas extra semanales.
                df_resumen_semanal = df_resumen_semanal[df_resumen_semanal['Horas_Extra_Semanales_Hrs'] > 0].copy()

                # Formatea la columna de inicio de semana y redondea las horas para el reporte final.
                df_resumen_semanal['Semana_Inicio'] = df_resumen_semanal['Semana_Inicio'].dt.strftime('%Y-%m-%d')
                df_resumen_semanal['Horas_Trabajadas_Calculadas_Semana_Hrs'] = round(df_resumen_semanal['Horas_Trabajadas_Calculadas_Semana_Hrs'], 2)

                # Reordena las columnas para el reporte final.
                df_resumen_semanal = df_resumen_semanal[[
                    'COD_TRABAJADOR', 'NOMBRE', 'Semana_Inicio',
                    'Horas_Trabajadas_Calculadas_Semana_Hrs', 'Horas_Extra_Semanales_Hrs',
                ]]

                if not df_resumen_semanal.empty:
                    st.write("### Resumen Horas Extra Semanal")
                    st.dataframe(df_resumen_semanal)

                    # Permite al usuario descargar el resumen semanal como un archivo Excel.
                    excel_buffer_semanal = io.BytesIO()
                    df_resumen_semanal.to_excel(excel_buffer_semanal, index=False, engine='openpyxl')
                    excel_buffer_semanal.seek(0) # Regresar al inicio del buffer para la descarga.

                    st.download_button(
                        label="Descargar Resumen Horas Extra Semanal (Excel)",
                        data=excel_buffer_semanal,
                        file_name="resumen_horas_extra_semanal.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.info("No se encontraron horas extras semanales para reportar.")

            else:
                st.warning("No se pudieron calcular horas extras. Asegúrate de que el archivo Excel tenga los datos y formatos correctos y que haya registros de entrada y salida válidos.")

    except Exception as e:
        # Captura y muestra cualquier error que ocurra durante el procesamiento del archivo.
        st.error(f"Ocurrió un error al procesar el archivo: {e}. Asegúrate de que el archivo es un Excel válido y la hoja 'BaseDatos Modificada' existe y tiene el formato correcto.")

st.markdown("---")
st.caption("Somos NOEL DE CORAZÓN ❤️ ")

