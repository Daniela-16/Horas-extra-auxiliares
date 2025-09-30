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

LUGARES_TRABAJO_PRINCIPAL = [
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
    "NOEL_MDE_ING_MEN_ALERGENOS_SAL",
    "NOEL_MDE_MR_WAFER_RCH_CREMAS_ENT",
    "NOEL_MDE_MR_WAFER_RCH_CREMAS_SAL",
    
]

# Normaliza los nombres de los lugares de trabajo (minúsculas, sin espacios extra).
LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS = [lugar.strip().lower() for lugar in LUGARES_TRABAJO_PRINCIPAL]

# Tolerancia en minutos para inferir si una marcación está cerca del inicio/fin de un turno.
TOLERANCIA_INFERENCIA_MINUTOS = 50

# Límite máximo de horas que una salida puede exceder el fin de turno programado.
MAX_EXCESO_SALIDA_HRS = 3

# Hora de corte para determinar la 'fecha clave de turno' para turnos nocturnos.
# Las marcaciones antes de esta hora se asocian al día de turno anterior.
HORA_CORTE_NOCTURNO = datetime.strptime("08:00:00", "%H:%M:%S").time()

# Constante para la tolerancia de llegada tarde
TOLERANCIA_LLEGADA_TARDE_MINUTOS = 40

# --- 3. Obtener turno basado en fecha y hora ---

def obtener_turno_para_registro(fecha_hora_evento: datetime, fecha_clave_turno_reporte: datetime.date, tolerancia_minutos: int):

    """
    Parámetros:
    - fecha_hora_evento (datetime): La fecha y hora de la marcación (usualmente la entrada).
    - fecha_clave_turno_reporte (datetime.date): La fecha lógica del turno (FECHA_CLAVE_TURNO)
                                                 usada para determinar el tipo de día (LV, SAB, DOM).
    - tolerancia_minutos (int): Minutos de flexibilidad alrededor del inicio/fin del turno.

    Retorna:
    - tupla (nombre_turno, info_turno_dict, inicio_turno_programado, fin_turno_programado)
      Si no se encuentra un turno, retorna (None, None, None, None).
    """

    # Determina el tipo de día usando la FECHA_CLAVE_TURNO_REPORTE, que es la fecha de entrada
    
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
    menor_diferencia = timedelta(days=999) # Inicializa dif grande

    # Itera sobre el diccionario de turnos definidos para el tipo de día (LV, SAB o DOM)
    
    for nombre_turno, info_turno in TURNOS[tipo_dia].items():
        hora_inicio = datetime.strptime(info_turno["inicio"], "%H:%M:%S").time()
        hora_fin = datetime.strptime(info_turno["fin"], "%H:%M:%S").time()

        # Prepara posibles fechas de inicio del turno
        # El primer candidato de inicio debe ser la FECHA_CLAVE_TURNO_REPORTE
        candidatos_inicio = [datetime.combine(fecha_clave_turno_reporte, hora_inicio)]


        # el principal inicio_posible_turno la fecha_clave_turno_reporte con la hora de inicio del turno.

        for inicio_posible_turno in candidatos_inicio:

            # Calcula el fin de turno correspondiente al inicio_posible_turno

            fin_posible_turno = inicio_posible_turno.replace(hour=hora_fin.hour, minute=hora_fin.minute, second=hora_fin.second)

            #si es turno nocturno

            if hora_inicio > hora_fin:
                fin_posible_turno += timedelta(days=1) # El fin ocurre al día siguiente


            # Compara la marcación real con el rango del turno programado.

            rango_inicio = inicio_posible_turno - timedelta(minutes=tolerancia_minutos)
            rango_fin = fin_posible_turno + timedelta(minutes=tolerancia_minutos)

            # Si no está dentro del rango con tolerancia, salta a la siguiente opción

            if not (rango_inicio <= fecha_hora_evento <= rango_fin):
                continue

            # Calcula la diferencia absoluta entre la marcación y el inicio programado del turno
            diferencia = abs(fecha_hora_evento - inicio_posible_turno)

            # actualiza las variables de menor diferencia y de mejor turno
            if mejor_turno is None or diferencia < menor_diferencia:
                mejor_turno = (nombre_turno, info_turno, inicio_posible_turno, fin_posible_turno)
                menor_diferencia = diferencia

    # Retorna el mejor turno encontrado o None si no hubo coincidencias
    return mejor_turno if mejor_turno else (None, None, None, None)

# --- 4. Calculo de horas ---

def calcular_turnos(df: pd.DataFrame, lugares_normalizados: list, tolerancia_minutos: int, tolerancia_llegada_tarde: int):

    """
    Parámetros:
    - df (pd.DataFrame): DataFrame con marcaciones preprocesadas, incluyendo 'FECHA_CLAVE_TURNO'.
    - lugares_normalizados (list): Lista de porterías válidas (normalizadas).
    - tolerancia_minutos (int): Tolerancia para la inferencia de turnos.
    - tolerancia_llegada_tarde (int): Minutos de gracia para considerar una llegada tarde.

    Retorna:
    - pd.DataFrame: Con los resultados de horas trabajadas y extra (incluyendo horas extra = 0).
    """

    # Filtra las marcaciones por los lugares principales y tipos 'ent'/'sal'

    df = df[(df['PORTERIA_NORMALIZADA'].isin(lugares_normalizados)) & (df['TIPO_MARCACION'].isin(['ent', 'sal']))]

    # Ordena para asegurar que las marcaciones estén en orden cronológico por trabajador

    df.sort_values(by=['ID_TRABAJADOR', 'FECHA_HORA'], inplace=True)

    if df.empty:
        return pd.DataFrame() # Retorna un DataFrame vacío si no hay datos para procesar

    resultados = [] # Lista para almacenar los resultados calculados

    # Agrupa por ID de trabajador y por fecha clave turno

    # Esto permite que los turnos nocturnos que cruzan la medianoche se agrupen en un mismo "listado"

    for (id_trabajador, fecha_clave_turno), grupo in df.groupby(['ID_TRABAJADOR', 'FECHA_CLAVE_TURNO']):

        nombre = grupo['NOMBRE'].iloc[0] #sera el mismo en todo el grupo debido a que se ordenó por nombre
        entradas = grupo[grupo['TIPO_MARCACION'] == 'ent'] # Marcaciones de entrada del grupo
        salidas = grupo[grupo['TIPO_MARCACION'] == 'sal'] # Marcaciones de salida del grupo
        
        # OBTENER TODAS LAS MARCAS DE ENTRADA Y SALIDA PARA EL REPORTE COMPLETO
        
        todas_entradas = ' | '.join(entradas['FECHA_HORA'].dt.strftime('%H:%M:%S').tolist())
        todas_salidas = ' | '.join(salidas['FECHA_HORA'].dt.strftime('%H:%M:%S').tolist())
        
        # Regla 1:
        # Si no hay entradas o salidas, se ignora el grupo

        if entradas.empty or salidas.empty:
            continue

        # Obtiene la primera entrada y la última salida real del grupo de marcaciones
        entrada_real = entradas['FECHA_HORA'].min()
        salida_real = salidas['FECHA_HORA'].max()

        porteria_entrada = entradas[entradas['FECHA_HORA'] == entrada_real]['PORTERIA'].iloc[0] if not entradas.empty else None
        porteria_salida = salidas[salidas['FECHA_HORA'] == salida_real]['PORTERIA'].iloc[0] if not salidas.empty else None

        # Regla 2:
        # Si la salida es antes o igual a la entrada, o la duración total es menor a 4 horas, se ignora.

        if salida_real <= entrada_real or (salida_real - entrada_real) < timedelta(hours=4):
            continue

        # Regla 3:
        # Intenta asignar un turno programado a la jornada

        turno_nombre, info_turno, inicio_turno, fin_turno = obtener_turno_para_registro(entrada_real, fecha_clave_turno, tolerancia_minutos)
        if turno_nombre is None:
            continue # Si no se puede asignar un turno, se ignora el grupo

        # Regla 4:
        # Valida que la salida real no exceda un límite razonable del fin de turno programado

        if salida_real > fin_turno + timedelta(hours=MAX_EXCESO_SALIDA_HRS):
            continue

        # --- Lógica de cálculo de horas basada en la nueva regla de llegada tarde ---
        inicio_efectivo_calculo = inicio_turno
        llegada_tarde_flag = False

        if entrada_real > inicio_turno:
            diferencia_entrada = entrada_real - inicio_turno
            if diferencia_entrada > timedelta(minutes=tolerancia_llegada_tarde):
                inicio_efectivo_calculo = entrada_real
                llegada_tarde_flag = True


        # Calcular la duración sobre la cual se aplicará la lógica de horas trabajadas y extra

        duracion_efectiva_calculo = salida_real - inicio_efectivo_calculo
        horas_trabajadas = round(duracion_efectiva_calculo.total_seconds() / 3600, 2) # Horas trabajadas desde la hora ajustada

        horas_turno = info_turno["duracion_hrs"] # Duración programada del turno asignado

        # Las horas extra son la duración efectiva trabajada menos la duración del turno, nunca negativa

        horas_extra = max(0, round(horas_trabajadas - horas_turno, 2))


        # Añade los resultados a la lista (SE INCLUYEN TODOS LOS REGISTROS VÁLIDOS AQUÍ)
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
            # Se usa el mismo cálculo de horas extra, pero se deja en el reporte
            'Horas_Extra_Enteras': int(horas_extra),
            'Minutos_Extra': round((horas_extra - int(horas_extra)) * 60),
            'Llegada_Tarde_Mas_40_Min': llegada_tarde_flag, # Nueva columna para indicar llegada tarde
            # Columnas agregadas para el reporte completo de marcaciones:
            'Todas_Marcaciones_Entrada_Hrs': todas_entradas,
            'Todas_Marcaciones_Salida_Hrs': todas_salidas,
        })

    return pd.DataFrame(resultados) # Retorna los resultados como un DataFrame

# --- Interfaz Streamlit ---
st.set_page_config(page_title="Reporte de Marcaciones", layout="wide")
st.title("📊 Reporte Completo de Marcaciones y Horas Trabajadas")
st.write("Sube tu archivo de Excel para obtener el reporte de marcaciones válidas (incluyendo horas extra).")

archivo_excel = st.file_uploader("Sube un archivo Excel (.xlsx)", type=["xlsx"])

if archivo_excel is not None:
    try:
        df_raw = pd.read_excel(archivo_excel, sheet_name='BaseDatos Modificada')

        columnas = ['COD_TRABAJADOR', 'APUNTADOR', 'DESC_APUNTADOR', 'NOMBRE', 'FECHA', 'HORA', 'PORTERIA', 'PuntoMarcacion']
        if not all(col in df_raw.columns for col in columnas):
            st.error(f"ERROR: Faltan columnas requeridas: {', '.join(columnas)}")
        else:
            # Preprocesamiento inicial de columnas

            df_raw['FECHA'] = pd.to_datetime(df_raw['FECHA'], errors='coerce')
            
            # --- AJUSTE CLAVE DE MANEJO DE HORA (Más robusto) ---
            # 1. Limpiar y convertir la columna HORA a string
            df_raw['HORA_STR'] = df_raw['HORA'].astype(str).str.strip()

            # 2. Función para estandarizar el formato de hora (añadir segundos si faltan)
            def standardize_time_format(time_str):
                parts = time_str.split(':')
                if len(parts) == 2:
                    return f"{time_str}:00"
                return time_str

            df_raw['HORA_STANDARDIZED'] = df_raw['HORA_STR'].apply(standardize_time_format)

            # 3. Combinar FECHA y HORA de manera más segura usando Series y to_datetime
            # Asegura que solo se intenten combinar valores válidos de FECHA
            fecha_valida_mask = df_raw['FECHA'].notna()
            
            df_raw.loc[fecha_valida_mask, 'FECHA_HORA'] = pd.to_datetime(
                df_raw.loc[fecha_valida_mask, 'FECHA'].dt.strftime('%Y-%m-%d') + ' ' + df_raw.loc[fecha_valida_mask, 'HORA_STANDARDIZED'],
                errors='coerce' # Si falla la conversión, será NaT
            )
            # Rellenar con NaT donde no se pudo combinar
            df_raw['FECHA_HORA'] = df_raw['FECHA_HORA'].fillna(pd.NaT)
            # ---------------------------------------------------

            df_raw['PORTERIA_NORMALIZADA'] = df_raw['PORTERIA'].astype(str).str.strip().str.lower()
            df_raw['TIPO_MARCACION'] = df_raw['PuntoMarcacion'].astype(str).str.strip().str.lower().replace({'entrada': 'ent', 'salida': 'sal'})
            df_raw.rename(columns={'COD_TRABAJADOR': 'ID_TRABAJADOR'}, inplace=True)
            
            # Filtrar filas donde la fecha/hora es válida
            df_raw = df_raw.dropna(subset=['FECHA_HORA'])

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

            # fin de la fecha clave

            st.success("Archivo cargado y preprocesado con éxito.")

            # Llama a la función principal de cálculo con el DataFrame modificado
            df_resultado = calcular_turnos(df_raw.copy(), LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS, TOLERANCIA_INFERENCIA_MINUTOS, TOLERANCIA_LLEGADA_TARDE_MINUTOS)

            if not df_resultado.empty:
                # Renombra las columnas de horas extra para el reporte
                df_resultado.rename(columns={
                    'Horas_Extra_Enteras': 'Horas',
                    'Minutos_Extra': 'Minutos',
                    'Llegada_Tarde_Mas_40_Min': 'Llegada_Tarde_Flag' # Columna auxiliar para formato Excel
                }, inplace=True)
                
                # 1. Prepara la columna booleana para el formato de Excel
                df_resultado['Llegada_Tarde_Bool'] = df_resultado['Llegada_Tarde_Flag']
                
                # 2. Mapea los valores True/False a 'Tarde'/'A tiempo' en la columna 'Estado_Llegada'
                df_resultado['Estado_Llegada'] = df_resultado['Llegada_Tarde_Flag'].map({True: 'Tarde', False: 'A tiempo'})
                
                # 3. Elimina la columna booleana original para la visualización y descarga principal
                df_resultado.drop(columns=['Llegada_Tarde_Flag'], inplace=True)


                st.subheader("Reporte Completo de Marcaciones y Horas Trabajadas")
                # Se muestra el DataFrame modificado en Streamlit, sin la columna booleana auxiliar
                st.dataframe(df_resultado.drop(columns=['Llegada_Tarde_Bool'])) 

                # Prepara el DataFrame para descarga en formato Excel
                buffer_excel = io.BytesIO()
                with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
                    
                    # El DataFrame a exportar es el resultado SIN la columna booleana auxiliar.
                    df_to_excel = df_resultado.drop(columns=['Llegada_Tarde_Bool']).copy() 
                    df_to_excel.to_excel(writer, sheet_name='Reporte Completo', index=False)

                    workbook = writer.book
                    worksheet = writer.sheets['Reporte Completo']

                    # Define a format for the orange background
                    orange_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

                    # Obtiene el índice de la columna 'ENTRADA_REAL' en el DataFrame exportado
                    try:
                        entrada_real_col_idx = df_to_excel.columns.get_loc('ENTRADA_REAL')
                    except KeyError:
                        entrada_real_col_idx = -1

                    if entrada_real_col_idx != -1:
                        
                        # USAMOS LA COLUMNA BOOLEANA GUARDADA PARA EL FORMATO
                        llegada_tarde_serie = df_resultado['Llegada_Tarde_Bool']
                        
                        # Itera a través de las filas para aplicar el formato. Empieza en la fila 1 (después de encabezados)
                        for row_num, is_late in enumerate(llegada_tarde_serie): 
                            excel_row = row_num + 1 # Fila de Excel (1-indexada)
                            entrada_valor = df_resultado.iloc[row_num]['ENTRADA_REAL']
                            
                            if is_late:
                                # Escribe el valor con el formato naranja
                                # Se asegura que el valor sea una cadena, que es el formato que sale de strftime
                                worksheet.write(excel_row, entrada_real_col_idx, str(entrada_valor), orange_format)
                            else:
                                # Escribe el valor sin formato
                                worksheet.write(excel_row, entrada_real_col_idx, str(entrada_valor))

                    buffer_excel.seek(0)

                    # Botón de descarga para el usuario
                    st.download_button(
                        label="Descargar Reporte Completo (Excel)",
                        data=buffer_excel,
                        file_name="Reporte_Marcaciones_y_Horas.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            else:
                st.warning("No se pudieron asignar turnos o hubo inconsistencias en los registros que cumplieran los criterios de validación (por ejemplo, duración mínima de 4 horas o la hora de corte para turnos nocturnos).")

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}. Asegúrate de que la hoja se llama 'BaseDatos Modificada' y que tiene todas las columnas requeridas.")

st.markdown("---")
st.caption("Somos NOEL DE CORAZÓN ❤️")



