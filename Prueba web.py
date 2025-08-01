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
        "Turno 3 LV": {"inicio": "21:40:00", "fin": "05:40:00", "duracion_hrs": 8}, # Turno nocturno
    },
    "SAB": { # Sábado
        "Turno 1 SAB": {"inicio": "05:40:00", "fin": "11:40:00", "duracion_hrs": 6},
        "Turno 2 SAB": {"inicio": "11:40:00", "fin": "17:40:00", "duracion_hrs": 6},
        "Turno 3 SAB": {"inicio": "21:40:00", "fin": "05:40:00", "duracion_hrs": 8}, # Turno nocturno
    },
    "DOM": { # Domingo
        "Turno 1 DOM": {"inicio": "05:40:00", "fin": "11:40:00", "duracion_hrs": 6},
        "Turno 2 DOM": {"inicio": "11:40:00", "fin": "17:40:00", "duracion_hrs": 6},
        "Turno 3 DOM": {"inicio": "22:40:00", "fin": "05:40:00", "duracion_hrs": 7}, # Turno nocturno
    }
}

# --- 2. Configuración General ---

LUGARES_TRABAJO = [
    "NOEL_MDE_OFIC_PRODUCCION_ENT",
    "NOEL_MDE_OFIC_PRODUCCION_SAL",
    "NOEL_MDE_MR_TUNEL_VIENTO_1_ENT",
    "NOEL_MDE_MR_MEZCLAS_ENT",
    "NOEL_MDE_ING_MEN_CREMAS_ENT",
    "NOEL_MDE_ING_MEN_CREMAS_SAL",
    "NOEL_MDE_MR_HORNO_6-8-9_ENT",
    "NOEL_MDE_MR_SERVICIOS_2_ENT",
    "NOEL_MDE_RECURSOS_HUMANOS_ENT",
    "NOEL_MDE_RECURSOS_HUMANOS_SAL",
    "NOEL_MDE_ESENCIAS_2_SAL",
    "NOEL_MDE_ESENCIAS_1_SAL",
    "NOEL_MDE_ING_MENORES_2_ENT",
    "NOEL_MDE_MR_HORNO_18_ENT",
    "NOEL_MDE_MR_WAFER_RCH_CREMAS_ENT",
    "NOEL_MDE_MR_HORNO_6-8-9_SAL",
    "NOEL_MDE_TORNIQUETE_SORTER_ENT",
    "NOEL_MDE_TORNIQUETE_SORTER_SAL",
    "NOEL_MDE_MR_MEZCLAS_SAL",
    "NOEL_MDE_MR_TUNEL_VIENTO_2_ENT",
    "NOEL_MDE_MR_HORNO_7-10_ENT",
    "NOEL_MDE_MR_HORNO_11_ENT",
    "NOEL_MDE_MR_WAFER_RCH_CREMAS_SAL",
    "NOEL_MDE_MR_HORNO_2-4-5_SAL",
    "NOEL_MDE_MR_HORNO_4-5_ENT",
    "NOEL_MDE_MR_HORNO_18_SAL",
    "NOEL_MDE_MR_HORNO_1-3_SAL",
    "NOEL_MDE_MR_HORNO_1-3_ENT",
    "NOEL_MDE_CONTROL_BUHLER_ENT",
    "NOEL_MDE_CONTROL_BUHLER_SAL",
    "NOEL_MDE_ING_MEN_ALERGENOS_ENT",
    "NOEL_MDE_ING_MENORES_2_SAL",
    "NOEL_MDE_MR_SERVICIOS_2_SAL",
    "NOEL_MDE_MR_HORNO_11_SAL",
    "NOEL_MDE_MR_HORNO_7-10_SAL",
    "NOEL_MDE_MR_HORNO_2-12_ENT",
    "NOEL_MDE_TORNIQUETE_PATIO_SAL",
    "NOEL_MDE_TORNIQUETE_PATIO_ENT",
    "NOEL_MDE_ESENCIAS_1_ENT",
    "NOEL_MDE_ING_MENORES_1_SAL",
    "NOEL_MDE_MOLINETE_BODEGA_EXT_SAL",
    "NOEL_MDE_PRINCIPAL_ENT",
    "NOEL_MDE_ING_MENORES_1_ENT",
    "NOEL_MDE_MR_HORNOS_SAL",
    "NOEL_MDE_MR_HORNO_6-8-9_SAL_2",
    "NOEL_MDE_PRINCIPAL_SAL",
    "NOEL_MDE_MR_ASPIRACION_ENT",
    "NOEL_MDE_MR_HORNO_2-12_SAL",
    "NOEL_MDE_MR_HORNOS_ENT",
    "NOEL_MDE_MR_HORNO_4-5_SAL",
    "NOEL_MDE_ING_MEN_ALERGENOS_SAL"
]

# Normaliza los nombres de los lugares de trabajo (minúsculas, sin espacios extra).
LUGARES_NORM = [lugar.strip().lower() for lugar in LUGARES_TRABAJO]

# Tolerancia en minutos para inferir si una marcación está cerca del inicio/fin de un turno.
TOLERANCIA_TURNO = 50

# Límite máximo de horas que una salida puede exceder el fin de turno programado.
MAX_SALIDA_HORAS = 3

# Hora de corte para determinar la 'fecha clave de turno' para turnos nocturnos.
# Las marcaciones antes de esta hora se asocian al día de turno anterior.
HORA_CORTE_NOCTURNO = datetime.strptime("08:00:00", "%H:%M:%S").time()

# Constante para la tolerancia de llegada tarde
LLEGADA_TARDE = 40

# --- 3. Obtener turno basado en fecha y hora ---

def obtener_turno_por_evento(dt_evento: datetime, fecha_clave_turno: datetime.date, tolerancia_min: int):
    """
    Parámetros:
    - dt_evento (datetime): La fecha y hora de la marcación (usualmente la entrada).
    - fecha_clave_turno (datetime.date): La fecha lógica del turno (FECHA_CLAVE_TURNO)
                                         usada para determinar el tipo de día (LV, SAB, DOM).
    - tolerancia_min (int): Minutos de flexibilidad alrededor del inicio/fin del turno.

    Retorna:
    - tupla (nombre_turno, info_turno_dict, inicio_turno_programado, fin_turno_programado)
      Si no se encuentra un turno, retorna (None, None, None, None).
    """
    # Determina el tipo de día usando la fecha_clave_turno, que es la fecha de entrada
    num_dia_semana = fecha_clave_turno.weekday() # 0=Lunes, 6=Domingo

    if num_dia_semana < 5: # Lunes a Viernes
        tipo_dia = "LV"
    elif num_dia_semana == 5: # Sábado
        tipo_dia = "SAB"
    else: # num_dia_semana == 6 (Domingo)
        tipo_dia = "DOM"

    # Asegurarse de que el tipo_dia exista en TURNOS para evitar KeyError
    if tipo_dia not in TURNOS:
        return (None, None, None, None)

    mejor_coincidencia = None
    min_diferencia = timedelta(days=999) # Inicializa dif grande

    # Itera sobre el diccionario de turnos definidos para el tipo de día (LV, SAB o DOM)
    for nombre_turno, info_turno in TURNOS[tipo_dia].items():
        hora_inicio = datetime.strptime(info_turno["inicio"], "%H:%M:%S").time()
        hora_fin = datetime.strptime(info_turno["fin"], "%H:%M:%S").time()

        # Prepara la posible fecha de inicio del turno
        inicio_posible = datetime.combine(fecha_clave_turno, hora_inicio)

        # Calcula el fin de turno correspondiente
        fin_posible = inicio_posible.replace(hour=hora_fin.hour, minute=hora_fin.minute, second=hora_fin.second)

        # Si es turno nocturno
        if hora_inicio > hora_fin:
            fin_posible += timedelta(days=1) # El fin ocurre al día siguiente

        # Compara la marcación real con el rango del turno programado.
        rango_inicio = inicio_posible - timedelta(minutes=tolerancia_min)
        rango_fin = fin_posible + timedelta(minutes=tolerancia_min)

        # Si no está dentro del rango con tolerancia, salta a la siguiente opción
        if not (rango_inicio <= dt_evento <= rango_fin):
            continue

        # Calcula la diferencia absoluta entre la marcación y el inicio programado del turno
        diferencia = abs(dt_evento - inicio_posible)

        # Actualiza las variables de menor diferencia y de mejor turno
        if mejor_coincidencia is None or diferencia < min_diferencia:
            mejor_coincidencia = (nombre_turno, info_turno, inicio_posible, fin_posible)
            min_diferencia = diferencia

    # Retorna el mejor turno encontrado o None si no hubo coincidencias
    return mejor_coincidencia if mejor_coincidencia else (None, None, None, None)

# --- 4. Calculo de horas ---

def calcular_horas(df: pd.DataFrame, lugares_norm: list, tolerancia_turnos_min: int, llegada_tarde_min: int):
    """
    Parámetros:
    - df (pd.DataFrame): DataFrame con marcaciones preprocesadas, incluyendo 'FECHA_CLAVE_TURNO'.
    - lugares_norm (list): Lista de porterías válidas (normalizadas).
    - tolerancia_turnos_min (int): Tolerancia para la inferencia de turnos.
    - llegada_tarde_min (int): Minutos de gracia para considerar una llegada tarde.

    Retorna:
    - pd.DataFrame: Con los resultados de horas trabajadas y extra.
    """
    # Filtra las marcaciones por los lugares principales y tipos 'ent'/'sal'
    df = df[(df['PORTERIA_NORMALIZADA'].isin(lugares_norm)) & (df['TIPO_MARCACION'].isin(['ent', 'sal']))]

    # Ordena para asegurar que las marcaciones estén en orden cronológico por trabajador
    df.sort_values(by=['ID_TRABAJADOR', 'FECHA_HORA'], inplace=True)

    if df.empty:
        return pd.DataFrame() # Retorna un DataFrame vacío si no hay datos para procesar

    resultados = [] # Lista para almacenar los resultados calculados

    # Agrupa por ID de trabajador y por fecha clave de turno
    # Esto permite que los turnos nocturnos que cruzan la medianoche se agrupen en un mismo "listado"
    for (id_trabajador, fecha_clave), grupo in df.groupby(['ID_TRABAJADOR', 'FECHA_CLAVE_TURNO']):
        nombre = grupo['NOMBRE'].iloc[0] #sera el mismo en todo el grupo debido a que se ordenó por nombre
        entradas = grupo[grupo['TIPO_MARCACION'] == 'ent'] # Marcaciones de entrada del grupo
        salidas = grupo[grupo['TIPO_MARCACION'] == 'sal'] # Marcaciones de salida del grupo

        # Regla 1: Si no hay entradas o salidas, se ignora el grupo
        if entradas.empty or salidas.empty:
            continue

        # Obtiene la primera entrada y la última salida real del grupo de marcaciones
        entrada_real = entradas['FECHA_HORA'].min()
        salida_real = salidas['FECHA_HORA'].max()

        porteria_entrada = entradas[entradas['FECHA_HORA'] == entrada_real]['PORTERIA'].iloc[0] if not entradas.empty else None
        porteria_salida = salidas[salidas['FECHA_HORA'] == salida_real]['PORTERIA'].iloc[0] if not salidas.empty else None

        # Regla 2: Si la salida es antes o igual a la entrada, o la duración total es menor a 4 horas, se ignora.
        if salida_real <= entrada_real or (salida_real - entrada_real) < timedelta(hours=4):
            continue

        # Regla 3: Intenta asignar un turno programado a la jornada
        turno_info = obtener_turno_por_evento(entrada_real, fecha_clave, tolerancia_turnos_min)
        if turno_info is None:
            continue # Si no se puede asignar un turno, se ignora el grupo
        
        nombre_turno, info_turno, inicio_turno, fin_turno = turno_info

        # Regla 4: Valida que la salida real no exceda un límite razonable del fin de turno programado
        if salida_real > fin_turno + timedelta(hours=MAX_SALIDA_HORAS):
            continue

        # --- Lógica de cálculo de horas basada en la nueva regla de llegada tarde ---
        inicio_efectivo = inicio_turno
        es_tarde = False

        if entrada_real > inicio_turno:
            diferencia_entrada = entrada_real - inicio_turno
            if diferencia_entrada > timedelta(minutes=llegada_tarde_min):
                inicio_efectivo = entrada_real
                es_tarde = True

        # Calcular la duración sobre la cual se aplicará la lógica de horas trabajadas y extra
        duracion_efectiva = salida_real - inicio_efectivo
        horas_trabajadas = round(duracion_efectiva.total_seconds() / 3600, 2) # Horas trabajadas desde la hora ajustada

        horas_turno = info_turno["duracion_hrs"] # Duración programada del turno asignado

        # Las horas extra son la duración efectiva trabajada menos la duración del turno, nunca negativa
        horas_extra = max(0, round(horas_trabajadas - horas_turno, 2))

        # Añade los resultados a la lista
        resultados.append({
            'NOMBRE': nombre,
            'ID_TRABAJADOR': id_trabajador,
            'FECHA': fecha_clave, # Usa la fecha clave de turno para el reporte
            'Dia_Semana': fecha_clave.strftime('%A'), # Día de la semana de la fecha clave de turno
            'TURNO': nombre_turno,
            'Inicio_Turno_Programado': inicio_turno.strftime("%H:%M:%S"),
            'Fin_Turno_Programado': fin_turno.strftime("%H:%M:%S"),
            'Duracion_Turno_Programado_Hrs': horas_turno, # Corregido: se usa horas_turno aquí
            'ENTRADA_REAL': entrada_real.strftime("%Y-%m-%d %H:%M:%S"), # Muestra la entrada real (sin cambiar)
            'PORTERIA_ENTRADA': porteria_entrada,
            'SALIDA_REAL': salida_real.strftime("%Y-%m-%d %H:%M:%S"),
            'PORTERIA_SALIDA': porteria_salida,
            'Horas_Trabajadas': horas_trabajadas, # Ahora muestra las horas calculadas desde la hora ajustada
            'Horas_Extra': horas_extra,
            'Horas_Extra_Enteras': int(horas_extra),
            'Minutos_Extra': round((horas_extra - int(horas_extra)) * 60),
            'Estado_llegada': es_tarde
        })

    return pd.DataFrame(resultados) # Retorna los resultados como un DataFrame

# --- Interfaz Streamlit ---
st.set_page_config(page_title="Calculadora de Horas Extra", layout="wide")
st.title("📊 Calculadora de Horas Extra")
st.write("Sube tu archivo de Excel para calcular las horas extra del personal.")

archivo_subido = st.file_uploader("Sube un archivo Excel (.xlsx)", type=["xlsx"])

if archivo_subido is not None:
    try:
        df_bruto = pd.read_excel(archivo_subido, sheet_name='BaseDatos Modificada')

        columnas_requeridas = ['COD_TRABAJADOR', 'APUNTADOR', 'DESC_APUNTADOR', 'NOMBRE', 'FECHA', 'HORA', 'PORTERIA', 'PuntoMarcacion']
        if not all(col in df_bruto.columns for col in columnas_requeridas):
            st.error(f"ERROR: Faltan columnas requeridas: {', '.join(columnas_requeridas)}")
        else:
            # Preprocesamiento inicial de columnas
            df_bruto['FECHA'] = pd.to_datetime(df_bruto['FECHA'])
            df_bruto['HORA'] = df_bruto['HORA'].astype(str)

            # Función para asegurar que la hora tenga el formato HH:MM:SS
            def estandarizar_formato_hora(cadena_hora):
                parts = cadena_hora.split(':')
                if len(parts) == 2: # El formato es HH:MM, se añaden ':00' para los segundos
                    return f"{cadena_hora}:00"
                elif len(parts) == 3: # El formato ya es HH:MM:SS
                    return cadena_hora
                else: # Manejar formatos inesperados, se retorna la cadena original
                    return cadena_hora

            df_bruto['HORA'] = df_bruto['HORA'].apply(estandarizar_formato_hora)
            
            df_bruto['FECHA_HORA'] = pd.to_datetime(df_bruto['FECHA'].dt.strftime('%Y-%m-%d') + ' ' + df_bruto['HORA'])
            df_bruto['PORTERIA_NORMALIZADA'] = df_bruto['PORTERIA'].astype(str).str.strip().str.lower()
            df_bruto['TIPO_MARCACION'] = df_bruto['PuntoMarcacion'].astype(str).str.strip().str.lower().replace({'entrada': 'ent', 'salida': 'sal'})
            df_bruto.rename(columns={'COD_TRABAJADOR': 'ID_TRABAJADOR'}, inplace=True)

            # --- Lógica: Asignar Fecha Clave de Turno para el agrupamiento ---
            # Esta función determina a qué 'día de turno' pertenece una marcación,
            # lo que es crucial para turnos nocturnos que cruzan la medianoche.
            def asignar_fecha_clave_turno(fila):
                fecha_original = fila['FECHA_HORA'].date()
                hora_marcacion = fila['FECHA_HORA'].time()
                tipo_marcacion = fila['TIPO_MARCACION'] # 'ent' o 'sal'

                # Si la marcación es una SALIDA y su hora es antes de HORA_CORTE_NOCTURNO,
                # entonces esa salida pertenece al turno que inició el día anterior.
                if tipo_marcacion == 'sal' and hora_marcacion < HORA_CORTE_NOCTURNO:
                    return fecha_original - timedelta(days=1)
                # Para ENTRADAS, o SALIDAS que son después de HORA_CORTE_NOCTURNO,
                # la fecha clave es la fecha de la marcación misma.
                else:
                    return fecha_original

            df_bruto['FECHA_CLAVE_TURNO'] = df_bruto.apply(asignar_fecha_clave_turno, axis=1)

            st.success("Archivo cargado y preprocesado con éxito.")

            df_resultado = calcular_horas(df_bruto.copy(), LUGARES_NORM, TOLERANCIA_TURNO, LLEGADA_TARDE)

            if not df_resultado.empty:
                df_resultado['Estado_llegada'] = df_resultado['Estado_llegada'].map({True: 'Tarde', False: 'A tiempo'})

                st.subheader("Resultados de las horas extra")
                st.dataframe(df_resultado)

                buffer_excel = io.BytesIO()
                with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
                    df_para_exportar = df_resultado.copy()
                    df_para_exportar.to_excel(writer, sheet_name='Reporte Horas Extra', index=False)

                    workbook = writer.book
                    worksheet = writer.sheets['Reporte Horas Extra']
                    
                    formato_llegada_tarde = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

                    try:
                        indice_columna_entrada = df_para_exportar.columns.get_loc('ENTRADA_REAL')
                    except KeyError:
                        indice_columna_entrada = -1

                    if indice_columna_entrada != -1:
                        for idx, estado_llegada in enumerate(df_para_exportar['Estado_llegada']):
                            if estado_llegada == 'Tarde':
                                worksheet.write(idx + 1, indice_columna_entrada, df_para_exportar.iloc[idx]['ENTRADA_REAL'], formato_llegada_tarde)
                            else:
                                worksheet.write(idx + 1, indice_columna_entrada, df_para_exportar.iloc[idx]['ENTRADA_REAL'])

                buffer_excel.seek(0)

                st.download_button(
                    label="Descargar Reporte de Horas extra (Excel)",
                    data=buffer_excel,
                    file_name="reporte_horas_extra.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("No se pudieron asignar turnos o hubo inconsistencias en los registros que cumplieran los criterios de cálculo. Revisa tus datos y las reglas del sistema.")

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}. Asegúrate de que la hoja se llama 'BaseDatos Modificada' y que tiene todas las columnas requeridas.")

st.markdown("---")
st.caption("Somos NOEL DE CORAZÓN ❤️")
