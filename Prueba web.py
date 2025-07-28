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
# Define los horarios de inicio, fin y duración para los turnos diurnos y nocturnos.
TURNOS = {
    "LV": { # Lunes a Viernes
        "Turno 1 LV": {"inicio": "05:40:00", "fin": "13:40:00", "duracion_hrs": 8},
        "Turno 2 LV": {"inicio": "13:40:00", "fin": "21:40:00", "duracion_hrs": 8},
        "Turno 3 LV": {"inicio": "21:40:00", "fin": "05:40:00", "duracion_hrs": 8}, # Turno nocturno que cruza la medianoche
    },
    "SAB": { # Sábado
        "Turno 1 SAB": {"inicio": "05:40:00", "fin": "11:40:00", "duracion_hrs": 6},
        "Turno 2 SAB": {"inicio": "11:40:00", "fin": "17:40:00", "duracion_hrs": 6},
        "Turno 3 SAB": {"inicio": "21:40:00", "fin": "05:40:00", "duracion_hrs": 8}, # Turno nocturno que cruza la medianoche
    },
    "DOM": { # Domingo
        "Turno 1 DOM": {"inicio": "05:40:00", "fin": "11:40:00", "duracion_hrs": 6},
        "Turno 2 DOM": {"inicio": "11:40:00", "fin": "17:40:00", "duracion_hrs": 6},
        "Turno 3 DOM": {"inicio": "22:40:00", "fin": "05:40:00", "duracion_hrs": 7}, # Turno nocturno que cruza la medianoche
    }
}

# --- 2. Configuración General ---
# Lista de lugares de trabajo principales que son relevantes para el cálculo de horas.
LUGARES_TRABAJO_PRINCIPAL = [
    "NOEL_MDE_OFIC_PRODUCCION_ENT", "NOEL_MDE_OFIC_PRODUCCION_SAL",
    "NOEL_MDE_MR_TUNEL_VIENTO_1_ENT", "NOEL_MDE_MR_MEZCLAS_ENT",
    "NOEL_MDE_MR_WAFER_RCH_CREMAS_ENT", "NOEL_MDE_MR_HORNO_11_ENT",
    "NOEL_MDE_MR_MEZCLAS_SAL", "NOEL_MDE_MR_HORNO_6-8-9_SAL",
    "NOEL_MDE_MR_SERVICIOS_2_SAL", "NOEL_MDE_MR_TUNEL_VIENTO_2_ENT",
    "NOEL_MDE_MR_SERVICIOS_2_ENT", "NOEL_MDE_MR_HORNO_1-3_ENT",
    "NOEL_MDE_MR_WAFER_RCH_CREMAS_SAL", "NOEL_MDE_MR_HORNO_6-8-9_ENT",
    "NOEL_MDE_MR_HORNO_11_SAL", "NOEL_MDE_MR_HORNOS_ENT",
    "NOEL_MDE_MR_HORNO_2-12_ENT", "NOEL_MDE_ING_MEN_CREMAS_ENT",
    "NOEL_MDE_ING_MEN_CREMAS_SAL", "NOEL_MDE_ING_MEN_ALERGENOS_ENT",
    "NOEL_MDE_MR_HORNO_4-5_ENT", "NOEL_MDE_ESENCIAS_2_SAL",
    "NOEL_MDE_ESENCIAS_1_ENT", "NOEL_MDE_ESENCIAS_1_SAL",
    "NOEL_MDE_MR_HORNO_6-8-9_SAL_2", "NOEL_MDE_MR_ASPIRACION_ENT",
    "NOEL_MDE_ING_MENORES_1_SAL", "NOEL_MDE_ING_MENORES_2_ENT",
    "NOEL_MDE_ING_MENORES_2_SAL", "NOEL_MDE_MR_HORNO_1-3_SAL",
    "NOEL_MDE_MR_HORNO_18_ENT", "NOEL_MDE_MR_HORNO_18_SAL",
    "NOEL_MDE_MR_HORNOS_SAL", "NOEL_MDE_ING_MENORES_1_ENT",
    "NOEL_MDE_MR_HORNO_7-10_SAL", "NOEL_MDE_MR_HORNO_7-10_ENT"
]
# Normaliza los nombres de los lugares de trabajo para facilitar comparaciones (minúsculas, sin espacios extra).
LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS = [lugar.strip().lower() for lugar in LUGARES_TRABAJO_PRINCIPAL]

# Tolerancia en minutos para inferir si una marcación está cerca del inicio/fin de un turno.
TOLERANCIA_INFERENCIA_MINUTOS = 50
# Límite máximo de horas que una salida puede exceder el fin de turno programado.
MAX_EXCESO_SALIDA_HRS = 3
# Hora de corte para determinar la 'fecha clave de turno' para turnos nocturnos.
# Las marcaciones antes de esta hora se asocian al día de turno anterior.
# Se ajusta a 06:00:00 para asegurar que salidas de turnos nocturnos (hasta 05:40)
# sean correctamente asignadas al día de turno anterior.
HORA_CORTE_NOCTURNO = datetime.strptime("06:00:00", "%H:%M:%S").time()

# --- 3. Obtener turno basado en fecha y hora ---
def obtener_turno_para_registro(fecha_hora_evento: datetime, tolerancia_minutos: int):
    """
    Identifica el turno programado más probable al que pertenece una marcación de evento.
    Maneja turnos que cruzan la medianoche.

    Parámetros:
    - fecha_hora_evento (datetime): La fecha y hora de la marcación (usualmente la entrada).
    - tolerancia_minutos (int): Minutos de flexibilidad alrededor del inicio/fin del turno.

    Retorna:
    - tupla (nombre_turno, info_turno_dict, inicio_turno_programado, fin_turno_programado)
      Si no se encuentra un turno, retorna (None, None, None, None).
    """
    dia_semana = fecha_hora_evento.weekday() # 0=Lunes, 6=Domingo
    tipo_dia = "LV" if dia_semana < 5 "SAB" if dia_semana = 5 "DOM" if dia_semana = 6 

    mejor_turno = None
    menor_diferencia = timedelta(days=999) # Inicializa con una diferencia muy grande

    # Itera sobre los turnos definidos para el tipo de día (LV o SAB)
    for nombre_turno, info_turno in TURNOS[tipo_dia].items():
        hora_inicio = datetime.strptime(info_turno["inicio"], "%H:%M:%S").time()
        hora_fin = datetime.strptime(info_turno["fin"], "%H:%M:%S").time()

        # Prepara posibles fechas de inicio del turno
        # El primer candidato es el turno que empieza el mismo día de la marcación
        candidatos_inicio = [fecha_hora_evento.replace(hour=hora_inicio.hour, minute=hora_inicio.minute, second=hora_inicio.second)]

        # Si el turno es nocturno (ej. 21:00 a 05:00), la marcación podría ser del día siguiente.
        # Añade un candidato de inicio de turno para el día anterior.
        if hora_inicio > hora_fin:
            candidatos_inicio.append((fecha_hora_evento - timedelta(days=1)).replace(hour=hora_inicio.hour, minute=hora_inicio.minute, second=hora_inicio.second))

        # Evalúa cada posible inicio de turno
        for inicio_posible_turno in candidatos_inicio:
            # Calcula el fin de turno correspondiente al inicio_posible_turno
            fin_posible_turno = inicio_posible_turno.replace(hour=hora_fin.hour, minute=hora_fin.minute, second=hora_fin.second)
            if hora_inicio > hora_fin:
                fin_posible_turno += timedelta(days=1) # Ajusta el fin para turnos nocturnos

            # Verifica si la fecha_hora_evento cae dentro del rango del turno con tolerancia
            rango_inicio = inicio_posible_turno - timedelta(minutes=tolerancia_minutos)
            rango_fin = fin_posible_turno + timedelta(minutes=tolerancia_minutos)

            if not (rango_inicio <= fecha_hora_evento <= rango_fin):
                continue # Si no está dentro del rango con tolerancia, salta a la siguiente opción

            # Calcula la diferencia absoluta entre la marcación y el inicio programado del turno
            diferencia = abs(fecha_hora_evento - inicio_posible_turno)

            # Si es el primer turno válido encontrado o es una mejor coincidencia
            if mejor_turno is None or diferencia < menor_diferencia:
                mejor_turno = (nombre_turno, info_turno, inicio_posible_turno, fin_posible_turno)
                menor_diferencia = diferencia

    # Retorna el mejor turno encontrado o None si no hubo coincidencias
    return mejor_turno if mejor_turno else (None, None, None, None)

# --- 4. Calculo de horas ---
def calcular_turnos(df: pd.DataFrame, lugares_normalizados: list, tolerancia_minutos: int):
    """
    Procesa un DataFrame de marcaciones para calcular horas trabajadas y horas extra por empleado y 'día de turno'.

    Parámetros:
    - df (pd.DataFrame): DataFrame con marcaciones preprocesadas, incluyendo 'FECHA_CLAVE_TURNO'.
    - lugares_normalizados (list): Lista de porterías válidas (normalizadas).
    - tolerancia_minutos (int): Tolerancia para la inferencia de turnos.

    Retorna:
    - pd.DataFrame: Con los resultados de horas trabajadas y extra.
    """
    # Filtra las marcaciones a solo los lugares principales y tipos 'ent'/'sal'
    df = df[(df['PORTERIA_NORMALIZADA'].isin(lugares_normalizados)) & (df['TIPO_MARCACION'].isin(['ent', 'sal']))]
    # Ordena para asegurar que las marcaciones estén en orden cronológico por trabajador
    df.sort_values(by=['ID_TRABAJADOR', 'FECHA_HORA'], inplace=True)

    if df.empty:
        return pd.DataFrame() # Retorna un DataFrame vacío si no hay datos para procesar

    resultados = [] # Lista para almacenar los resultados calculados

    # Agrupa por ID de trabajador y la NUEVA 'FECHA_CLAVE_TURNO'.
    # Esto permite que los turnos nocturnos que cruzan la medianoche se agrupen correctamente.
    for (id_trabajador, fecha_clave_turno), grupo in df.groupby(['ID_TRABAJADOR', 'FECHA_CLAVE_TURNO']):
        nombre = grupo['NOMBRE'].iloc[0] # El nombre del trabajador (asumiendo que es el mismo en el grupo)
        entradas = grupo[grupo['TIPO_MARCACION'] == 'ent'] # Marcaciones de entrada del grupo
        salidas = grupo[grupo['TIPO_MARCACION'] == 'sal'] # Marcaciones de salida del grupo

        # Regla 1: Si no hay entradas o salidas, se ignora el grupo (jornada incompleta)
        if entradas.empty or salidas.empty:
            continue

        # Obtiene la primera entrada y la última salida real del grupo de marcaciones
        entrada_real = entradas['FECHA_HORA'].min()
        salida_real = salidas['FECHA_HORA'].max()

        # Regla 2: Valida la consistencia básica de las marcaciones y una duración mínima de jornada
        # Si la salida es antes o igual a la entrada, o la duración total es menor a 5 horas, se ignora.
        if salida_real <= entrada_real or (salida_real - entrada_real) < timedelta(hours=5):
            continue

        # Regla 3: Intenta asignar un turno programado a la jornada
        # La función obtener_turno_para_registro maneja la lógica de turnos nocturnos.
        turno_nombre, info_turno, inicio_turno, fin_turno = obtener_turno_para_registro(entrada_real, tolerancia_minutos)
        if turno_nombre is None:
            continue # Si no se puede asignar un turno, se ignora el grupo

        # Regla 4: Valida que la salida real no exceda un límite razonable del fin de turno programado
        # Esto ayuda a filtrar errores de marcación (ej. olvido de marcar salida)
        if salida_real > fin_turno + timedelta(hours=MAX_EXCESO_SALIDA_HRS):
            continue

        # --- INICIO DE LOS CAMBIOS PARA HORAS TRABAJADAS Y EXTRA ---
        # Definir el inicio efectivo para el cálculo de horas trabajadas y extra:
        # SIEMPRE será el inicio programado del turno, independientemente de la entrada real.
        inicio_efectivo_calculo = inicio_turno # <-- Este es el cambio principal
        
        # Calcular la duración sobre la cual se aplicará la lógica de horas trabajadas y extra
        duracion_efectiva_calculo = salida_real - inicio_efectivo_calculo
        horas_trabajadas = round(duracion_efectiva_calculo.total_seconds() / 3600, 2) # Horas trabajadas desde la hora ajustada
        
        horas_turno = info_turno["duracion_hrs"] # Duración programada del turno asignado

        # Las horas extra son la duración efectiva trabajada menos la duración del turno, nunca negativa
        horas_extra = max(0, round(horas_trabajadas - horas_turno, 2))
        # --- FIN DE LOS CAMBIOS PARA HORAS TRABAJADAS Y EXTRA ---

        # Añade los resultados a la lista
        resultados.append({
            'NOMBRE': nombre,
            'ID_TRABAJADOR': id_trabajador,
            'FECHA': fecha_clave_turno, # Usa la fecha clave de turno para el reporte
            'Dia_Semana': fecha_clave_turno.strftime('%A'), # Día de la semana de la fecha clave de turno
            'TURNO': turno_nombre,
            'Inicio_Turno_Programado': inicio_turno.strftime("%H:%M:%S"),
            'Fin_Turno_Programado': fin_turno.strftime("%H:%M:%S"),
            'Duracion_Turno_Programado_Hrs': horas_turno,
            'ENTRADA_REAL': entrada_real.strftime("%Y-%m-%d %H:%M:%S"), # Muestra la entrada real (sin cambiar)
            'SALIDA_REAL': salida_real.strftime("%Y-%m-%d %H:%M:%S"),
            'Horas_Trabajadas': horas_trabajadas, # Ahora muestra las horas calculadas desde la hora ajustada
            'Horas_Extra': horas_extra,
            'Horas_Extra_Enteras': int(horas_extra),
            'Minutos_Extra': round((horas_extra - int(horas_extra)) * 60)
        })

    return pd.DataFrame(resultados) # Retorna los resultados como un DataFrame

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
            # Preprocesamiento inicial de columnas
            df_raw['FECHA'] = pd.to_datetime(df_raw['FECHA'])
            df_raw['HORA'] = df_raw['HORA'].astype(str)
            df_raw['FECHA_HORA'] = pd.to_datetime(df_raw['FECHA'].dt.strftime('%Y-%m-%d') + ' ' + df_raw['HORA'])
            df_raw['PORTERIA_NORMALIZADA'] = df_raw['PORTERIA'].astype(str).str.strip().str.lower()
            df_raw['TIPO_MARCACION'] = df_raw['PuntoMarcacion'].astype(str).str.strip().str.lower().replace({'entrada': 'ent', 'salida': 'sal'})
            df_raw.rename(columns={'COD_TRABAJADOR': 'ID_TRABAJADOR'}, inplace=True)

            # --- LÓGICA: Asignar Fecha Clave de Turno para el agrupamiento ---
            # Esta función determina a qué 'día de turno' pertenece una marcación,
            # lo que es crucial para turnos nocturnos que cruzan la medianoche.
            def asignar_fecha_clave_turno(row):
                fecha_original = row['FECHA_HORA'].date()
                hora_marcacion = row['FECHA_HORA'].time()

                # Si la marcación es antes de HORA_CORTE_NOCTURNO (ej. 6 AM),
                # se asume que forma parte del turno que inició el día anterior.
                if hora_marcacion < HORA_CORTE_NOCTURNO:
                    return fecha_original - timedelta(days=1)
                else:
                    return fecha_original
            
            # Aplica la función para crear la nueva columna en el DataFrame
            df_raw['FECHA_CLAVE_TURNO'] = df_raw.apply(asignar_fecha_clave_turno, axis=1)
            # --- FIN LÓGICA DE FECHA CLAVE ---

            st.success("Archivo cargado y preprocesado con éxito.")
            # Llama a la función principal de cálculo con el DataFrame modificado
            df_resultado = calcular_turnos(df_raw.copy(), LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS, TOLERANCIA_INFERENCIA_MINUTOS)

            if not df_resultado.empty:
                st.subheader("Resultados de las horas extra")
                st.dataframe(df_resultado)

                # Prepara el DataFrame para descarga en formato Excel
                buffer_excel = io.BytesIO()
                df_resultado.to_excel(buffer_excel, index=False, engine='openpyxl')
                buffer_excel.seek(0)

                # Botón de descarga para el usuario
                st.download_button(
                    label="Descargar Reporte de Horas extra (Excel)",
                    data=buffer_excel,
                    file_name="reporte_horas_extra.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("No se pudieron asignar turnos o hubo inconsistencias en los registros que cumplieran los criterios de cálculo. Revisa tus datos y las reglas del sistema (por ejemplo, duración mínima de 5 horas o la hora de corte para turnos nocturnos).")

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}. Asegúrate de que la hoja se llama 'BaseDatos Modificada' y que tiene todas las columnas requeridas.")

st.markdown("---")
st.caption("Somos NOEL DE CORAZÓN ❤️")
