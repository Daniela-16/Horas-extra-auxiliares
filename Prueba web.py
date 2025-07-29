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
HORA_CORTE_NOCTURNO = datetime.strptime("08:00:00", "%H:%M:%S").time()

# --- 3. Obtener turno basado en fecha y hora ---
# AHORA toma un parámetro adicional: fecha_clave_turno_reporte
def obtener_turno_para_registro(fecha_hora_evento: datetime, fecha_clave_turno_reporte: datetime.date, tolerancia_minutos: int):
    """
    Identifica el turno programado más probable al que pertenece una marcación de evento.
    Maneja turnos que cruzan la medianoche.

    Parámetros:
    - fecha_hora_evento (datetime): La fecha y hora de la marcación (usualmente la entrada).
    - fecha_clave_turno_reporte (datetime.date): La fecha lógica del turno (FECHA_CLAVE_TURNO)
                                                  usada para determinar el tipo de día (LV, SAB, DOM).
    - tolerancia_minutos (int): Minutos de flexibilidad alrededor del inicio/fin del turno.

    Retorna:
    - tupla (nombre_turno, info_turno_dict, inicio_turno_programado, fin_turno_programado)
      Si no se encuentra un turno, retorna (None, None, None, None).
    """
    # Determinamos el tipo de día usando la FECHA_CLAVE_TURNO_REPORTE, no la fecha_hora_evento directamente
    dia_semana_clave = fecha_clave_turno_reporte.weekday() # 0=Lunes, 6=Domingo
    
    if dia_semana_clave < 5: # Lunes a Viernes
        tipo_dia = "LV"
    elif dia_semana_clave == 5: # Sábado
        tipo_dia = "SAB"
    else: # dia_semana_clave == 6 (Domingo)
        tipo_dia = "DOM"

    # Asegurarse de que el tipo_dia exista en TURNOS para evitar KeyError
    if tipo_dia not in TURNOS:
        return (None, None, None, None)

    mejor_turno = None
    menor_diferencia = timedelta(days=999) # Inicializa con una diferencia muy grande

    # Itera sobre los turnos definidos para el tipo de día (LV, SAB o DOM)
    for nombre_turno, info_turno in TURNOS[tipo_dia].items():
        hora_inicio = datetime.strptime(info_turno["inicio"], "%H:%M:%S").time()
        hora_fin = datetime.strptime(info_turno["fin"], "%H:%M:%S").time()

        # Prepara posibles fechas de inicio del turno
        # El primer candidato de inicio debe ser la FECHA_CLAVE_TURNO_REPORTE
        candidatos_inicio = [datetime.combine(fecha_clave_turno_reporte, hora_inicio)]

        # Si el turno es nocturno (ej. 21:00 a 05:00), la marcación podría ser del día siguiente.
        # En este caso, la fecha_clave_turno_reporte ya es el día de inicio lógico.
        # No necesitamos el día anterior aquí para el inicio_posible_turno
        # Si la hora de inicio es posterior a la hora de la marcación y la hora de fin es anterior, significa
        # que el turno nocturno puede empezar en fecha_clave_turno_reporte y terminar al día siguiente.
        # La lógica de los candidatos_inicio está bien, porque la marcación se puede dar en el día siguiente.

        # Evalúa cada posible inicio de turno. En este contexto, el principal inicio_posible_turno
        # es la combinación de la fecha_clave_turno_reporte con la hora de inicio del turno.
        for inicio_posible_turno in candidatos_inicio:
            # Calcula el fin de turno correspondiente al inicio_posible_turno
            fin_posible_turno = inicio_posible_turno.replace(hour=hora_fin.hour, minute=hora_fin.minute, second=hora_fin.second)
            if hora_inicio > hora_fin: # Si es un turno nocturno
                fin_posible_turno += timedelta(days=1) # El fin ocurre al día siguiente

            # Verifica si la fecha_hora_evento (marcación real) cae dentro del rango del turno con tolerancia
            # Aquí es donde se compara la marcación real con el rango del turno programado.
            rango_inicio = inicio_posible_turno - timedelta(minutes=tolerancia_minutos)
            rango_fin = fin_posible_turno + timedelta(minutes=tolerancia_minutos)

            # Para turnos nocturnos, la fecha_hora_evento puede ser del día siguiente
            # Es importante que el rango_inicio pueda extenderse al día anterior si la marcación es muy temprana.
            # Y el rango_fin pueda extenderse al día siguiente si la marcación es muy tardía.
            # La marcación (fecha_hora_evento) para un turno nocturno con fecha_clave_turno_reporte = Domingo
            # podría ser Lunes 00:00:00 como en tu ejemplo.
            # Entonces, el inicio_posible_turno para Turno 3 DOM (22:40-05:40) sería Domingo 22:40.
            # Y el fin_posible_turno sería Lunes 05:40.
            # El rango para la marcación 00:00 del lunes debe ser (Domingo 22:40 - 50min) a (Lunes 05:40 + 50min)
            # Esto debería funcionar con la lógica actual si el 'fecha_hora_evento' cae dentro de este rango.

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

        porteria_entrada = entradas[entradas['FECHA_HORA'] == entrada_real]['PORTERIA'].iloc[0] if not entradas.empty else None
        porteria_salida = salidas[salidas['FECHA_HORA'] == salida_real]['PORTERIA'].iloc[0] if not salidas.empty else None

        # Regla 2: Valida la consistencia básica de las marcaciones y una duración mínima de jornada
        # Si la salida es antes o igual a la entrada, o la duración total es menor a 5 horas, se ignora.
        if salida_real <= entrada_real or (salida_real - entrada_real) < timedelta(hours=5):
            continue

        # Regla 3: Intenta asignar un turno programado a la jornada
        # ¡IMPORTANTE CAMBIO AQUÍ! Pasamos fecha_clave_turno_reporte a la función
        turno_nombre, info_turno, inicio_turno, fin_turno = obtener_turno_para_registro(entrada_real, fecha_clave_turno, tolerancia_minutos)
        if turno_nombre is None:
            continue # Si no se puede asignar un turno, se ignora el grupo

        # Regla 4: Valida que la salida real no exceda un límite razonable del fin de turno programado
        # Esto ayuda a filtrar errores de marcación (ej. olvido de marcar salida)
        if salida_real > fin_turno + timedelta(hours=MAX_EXCESO_SALIDA_HRS):
            continue

        # --- INICIO DE LOS CAMBIOS PARA HORAS TRABAJADAS Y EXTRA ---
        # Definir el inicio efectivo para el cálculo de horas trabajadas y extra:
        # SIEMPRE será el inicio programado del turno, independientemente de la entrada real.
        inicio_efectivo_calculo = inicio_turno 
        
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
            'PORTERIA_ENTRADA': porteria_entrada,
            'SALIDA_REAL': salida_real.strftime("%Y-%m-%d %H:%M:%S"),
            'PORTERIA_SALIDA': porteria_salida,
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
                tipo_marcacion = row['TIPO_MARCACION'] # 'ent' o 'sal'

                # Si la marcación es una SALIDA y su hora es antes de HORA_CORTE_NOCTURNO,
                # entonces esa salida pertenece al turno que inició el día anterior.
                if tipo_marcacion == 'sal' and hora_marcacion < HORA_CORTE_NOCTURNO:
                    return fecha_original - timedelta(days=1)
                # Para ENTRADAS, o SALIDAS que son después de HORA_CORTE_NOCTURNO,
                # la fecha clave es la fecha de la marcación misma.
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
