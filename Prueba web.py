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
# NOTA IMPORTANTE: La FECHA CLAVE DE TURNO se basará en el día que INICIA el turno.
TURNOS = {
    "LV": { # Lunes a Viernes (0-4)
        "Turno 1 LV": {"inicio": "05:40:00", "fin": "13:40:00", "duracion_hrs": 8},
        "Turno 2 LV": {"inicio": "13:40:00", "fin": "21:40:00", "duracion_hrs": 8},
        "Turno 3 LV": {"inicio": "21:40:00", "fin": "05:40:00", "duracion_hrs": 8, "nocturno": True}, # Turno nocturno
    },
    "SAB": { # Sábado (5)
        "Turno 1 SAB": {"inicio": "05:40:00", "fin": "11:40:00", "duracion_hrs": 6},
        "Turno 2 SAB": {"inicio": "11:40:00", "fin": "17:40:00", "duracion_hrs": 6},
        "Turno 3 SAB": {"inicio": "21:40:00", "fin": "05:40:00", "duracion_hrs": 8, "nocturno": True}, # Turno nocturno
    },
    "DOM": { # Domingo (6)
        "Turno 1 DOM": {"inicio": "05:40:00", "fin": "11:40:00", "duracion_hrs": 6},
        "Turno 2 DOM": {"inicio": "11:40:00", "fin": "17:40:00", "duracion_hrs": 6},
        "Turno 3 DOM": {"inicio": "22:40:00", "fin": "05:40:00", "duracion_hrs": 7, "nocturno": True}, # Turno nocturno
    }
}

# --- 2. Configuración General ---

LUGARES_TRABAJO_PRINCIPAL = [
    # ... (Se mantiene la lista de lugares para ahorrar espacio en la respuesta)
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
    "NOEL_MDE_PORT_2_PEATONAL_1_ENT",
    "NOEL_MDE_TORN_PORTERIA_3_SAL",
    "NOEL_MDE_VEHICULAR_PORT_1_ENT",
    "NOEL_MDE_PORT_2_PEATONAL_1_SAL",
    "NOEL_MDE_PORT_2_PEATONAL_2_ENT",
    "NOEL_MDE_VEHICULAR_PORT_1_SAL",
    "NOEL_MDE_TORN_PORTERIA_3_ENT",
    "NOEL_MDE_PORT_2_PEATONAL_2_SAL",
    "NOEL_MDE_PORT_2_PEATONAL_3_SAL",
    "NOEL_MDE_PORT_2_PEATONAL_3_ENT",
    "NOEL_MDE_PORT_1_PEATONAL_1_ENT"
    
]

# Normaliza los nombres de los lugares de trabajo (minúsculas, sin espacios extra).
LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS = [lugar.strip().lower() for lugar in LUGARES_TRABAJO_PRINCIPAL]

# Tolerancia en minutos para inferir si una marcación está cerca del inicio/fin de un turno.
TOLERANCIA_INFERENCIA_MINUTOS = 50

# Límite máximo de horas que una salida puede exceder el fin de turno programado.
MAX_EXCESO_SALIDA_HRS = 3

# Hora de corte para determinar la 'fecha clave de turno' para turnos nocturnos.
# Las marcaciones de SALIDA antes de esta hora se asocian al día de turno anterior (si es nocturno).
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

    # Determina el tipo de día usando la FECHA_CLAVE_TURNO_REPORTE, que ahora se basa en la ENTRADA
    dia_semana_clave = fecha_clave_turno_reporte.weekday() # 0=Lunes, 6=Domingo

    if dia_semana_clave < 5: # Lunes a Viernes
        tipo_dia = "LV"
    elif dia_semana_clave == 5: # Sábado
        tipo_dia = "SAB"
    else: # dia_semana_clave == 6 (Domingo)
        tipo_dia = "DOM"

    if tipo_dia not in TURNOS:
        return (None, None, None, None)

    mejor_turno = None
    menor_diferencia = timedelta(days=999) # Inicializa dif grande

    # Itera sobre el diccionario de turnos definidos para el tipo de día (LV, SAB o DOM)
    for nombre_turno, info_turno in TURNOS[tipo_dia].items():
        hora_inicio = datetime.strptime(info_turno["inicio"], "%H:%M:%S").time()
        hora_fin = datetime.strptime(info_turno["fin"], "%H:%M:%S").time()
        es_nocturno = info_turno.get("nocturno", False)

        # El inicio del turno programado se fija en la FECHA_CLAVE_TURNO_REPORTE
        inicio_posible_turno = datetime.combine(fecha_clave_turno_reporte, hora_inicio)

        # Calcula el fin de turno. Si es nocturno, el fin será el día siguiente
        if es_nocturno:
            fin_posible_turno = datetime.combine(fecha_clave_turno_reporte + timedelta(days=1), hora_fin)
        else:
            fin_posible_turno = datetime.combine(fecha_clave_turno_reporte, hora_fin)

        # Compara la marcación real con el rango del turno programado (con tolerancia)
        rango_inicio = inicio_posible_turno - timedelta(minutes=tolerancia_minutos)
        
        # Para evitar que un turno posterior se "coma" la entrada de un turno diurno
        # el rango_fin solo debe extenderse hasta el inicio del siguiente turno (si aplica)
        # o un poco más allá de su fin programado.
        rango_fin = fin_posible_turno + timedelta(minutes=tolerancia_minutos)


        # Si la entrada real no cae dentro del rango de tolerancia del inicio del turno, se descarta.
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
    - pd.DataFrame: Con los resultados de horas trabajadas y extra.
    """

    # Mantener solo las marcaciones de lugares principales y tipos 'ent'/'sal'
    df_filtrado = df[(df['PORTERIA_NORMALIZADA'].isin(lugares_normalizados)) & (df['TIPO_MARCACION'].isin(['ent', 'sal']))].copy()

    # Ordena para asegurar que las marcaciones estén en orden cronológico por trabajador
    df_filtrado.sort_values(by=['ID_TRABAJADOR', 'FECHA_HORA'], inplace=True)

    if df_filtrado.empty:
        return pd.DataFrame() # Retorna un DataFrame vacío si no hay datos para procesar

    resultados = [] # Lista para almacenar los resultados calculados

    # Agrupa por ID de trabajador y por fecha clave turno
    for (id_trabajador, fecha_clave_turno), grupo in df_filtrado.groupby(['ID_TRABAJADOR', 'FECHA_CLAVE_TURNO']):

        nombre = grupo['NOMBRE'].iloc[0] # El nombre será el mismo en todo el grupo
        entradas = grupo[grupo['TIPO_MARCACION'] == 'ent'] # Marcaciones de entrada del grupo
        salidas = grupo[grupo['TIPO_MARCACION'] == 'sal'] # Marcaciones de salida del grupo

        # Obtiene la primera entrada y la última salida real del grupo de marcaciones
        # Si no hay entradas/salidas, el valor por defecto será None (o NaT para datetime)
        entrada_real = entradas['FECHA_HORA'].min() if not entradas.empty else pd.NaT
        salida_real = salidas['FECHA_HORA'].max() if not salidas.empty else pd.NaT

        porteria_entrada = entradas[entradas['FECHA_HORA'] == entrada_real]['PORTERIA'].iloc[0] if not entradas.empty and pd.notna(entrada_real) else 'Sin Entrada'
        porteria_salida = salidas[salidas['FECHA_HORA'] == salida_real]['PORTERIA'].iloc[0] if not salidas.empty and pd.notna(salida_real) else 'Sin Salida'

        # --- Intento de asignación de turno y validación de reglas para el CALCULO ---
        turno_nombre, info_turno, inicio_turno, fin_turno = (None, None, None, None)
        horas_trabajadas = 0.0
        horas_extra = 0.0
        llegada_tarde_flag = False
        estado_calculo = "No Calculado" # Nuevo estado para rastrear el por qué no se calculó

        if pd.notna(entrada_real) and pd.notna(salida_real):
            # Regla 2: Mínima duración para el cálculo (4 horas)
            if salida_real <= entrada_real or (salida_real - entrada_real) < timedelta(hours=4):
                estado_calculo = "Duración < 4h o Inconsistente"
            else:
                # Regla 3: Intenta asignar un turno programado a la jornada
                turno_nombre, info_turno, inicio_turno, fin_turno = obtener_turno_para_registro(entrada_real, fecha_clave_turno, tolerancia_minutos)
                
                if turno_nombre is None:
                    estado_calculo = "Turno No Asignado (Fuera de rango)"
                
                else:
                    # Regla 4: Valida que la salida real no exceda un límite razonable
                    if salida_real > fin_turno + timedelta(hours=MAX_EXCESO_SALIDA_HRS):
                        estado_calculo = "Salida Excede Límite"
                    else:
                        # --- Lógica de cálculo de horas basada en la regla de llegada tarde ---
                        inicio_efectivo_calculo = inicio_turno
                        
                        if entrada_real > inicio_turno:
                            diferencia_entrada = entrada_real - inicio_turno
                            if diferencia_entrada > timedelta(minutes=tolerancia_llegada_tarde):
                                # Se toma la entrada real como inicio efectivo
                                inicio_efectivo_calculo = entrada_real 
                                llegada_tarde_flag = True
                        
                        # Calcular la duración sobre la cual se aplicará la lógica de horas trabajadas y extra
                        duracion_efectiva_calculo = salida_real - inicio_efectivo_calculo
                        horas_trabajadas = round(duracion_efectiva_calculo.total_seconds() / 3600, 2)
                        
                        horas_turno = info_turno["duracion_hrs"]
                        
                        # Las horas extra son la duración efectiva trabajada menos la duración del turno, nunca negativa
                        horas_extra = max(0, round(horas_trabajadas - horas_turno, 2))
                        estado_calculo = "Calculado"

        elif pd.notna(entrada_real) and pd.isna(salida_real):
             estado_calculo = "Falta Salida"
        elif pd.isna(entrada_real) and pd.notna(salida_real):
             estado_calculo = "Falta Entrada"
        else:
             estado_calculo = "Sin Marcaciones Válidas"
             

        # --- Añade los resultados a la lista (Se reporta todo, con o sin cálculo) ---
        
        # Formateo condicional para el reporte
        ent_str = entrada_real.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(entrada_real) else 'N/A'
        sal_str = salida_real.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(salida_real) else 'N/A'
        inicio_str = inicio_turno.strftime("%H:%M:%S") if inicio_turno else 'N/A'
        fin_str = fin_turno.strftime("%H:%M:%S") if fin_turno else 'N/A'
        horas_turno_val = info_turno["duracion_hrs"] if info_turno else 0

        resultados.append({
            'NOMBRE': nombre,
            'ID_TRABAJADOR': id_trabajador,
            'FECHA': fecha_clave_turno, # Usa la fecha clave de turno para el reporte
            'Dia_Semana': fecha_clave_turno.strftime('%A'),
            'TURNO': turno_nombre if turno_nombre else 'N/A',
            'Inicio_Turno_Programado': inicio_str,
            'Fin_Turno_Programado': fin_str,
            'Duracion_Turno_Programado_Hrs': horas_turno_val,
            'ENTRADA_REAL': ent_str,
            'PORTERIA_ENTRADA': porteria_entrada,
            'SALIDA_REAL': sal_str,
            'PORTERIA_SALIDA': porteria_salida,
            'Horas_Trabajadas': horas_trabajadas,
            'Horas_Extra': horas_extra,
            'Horas': int(horas_extra),
            'Minutos': round((horas_extra - int(horas_extra)) * 60),
            'Llegada_Tarde_Mas_40_Min': llegada_tarde_flag,
            'Estado_Calculo': estado_calculo # Muestra el estado para saber por qué no se calculó
        })

    return pd.DataFrame(resultados)

# --- Interfaz Streamlit ---
st.set_page_config(page_title="Calculadora de Horas Extra", layout="wide")
st.title("📊 Calculadora de Horas Extra")
st.write("Sube tu archivo de Excel para calcular las horas extra del personal.")

archivo_excel = st.file_uploader("Sube un archivo Excel (.xlsx)", type=["xlsx"])

if archivo_excel is not None:
    try:
        # Se asume que la hoja se llama 'BaseDatos Modificada'
        df_raw = pd.read_excel(archivo_excel, sheet_name='BaseDatos Modificada')

        columnas = ['COD_TRABAJADOR', 'APUNTADOR', 'DESC_APUNTADOR', 'NOMBRE', 'FECHA', 'HORA', 'PORTERIA', 'PuntoMarcacion']
        if not all(col in df_raw.columns for col in columnas):
            st.error(f"ERROR: Faltan columnas requeridas: {', '.join(columnas)}. Asegúrate de que las columnas están escritas exactamente igual.")
        else:
            # Preprocesamiento inicial de columnas
            df_raw['FECHA'] = pd.to_datetime(df_raw['FECHA'])
            df_raw['HORA'] = df_raw['HORA'].astype(str)

            # Función para asegurar que la hora tenga el formato HH:MM:SS
            def standardize_time_format(time_str):
                parts = time_str.split(':')
                if len(parts) == 2:
                    return f"{time_str}:00"
                elif len(parts) == 3:
                    return time_str
                else:
                    return time_str

            df_raw['HORA'] = df_raw['HORA'].apply(standardize_time_format)
            
            df_raw['FECHA_HORA'] = pd.to_datetime(df_raw['FECHA'].dt.strftime('%Y-%m-%d') + ' ' + df_raw['HORA'])
            df_raw['PORTERIA_NORMALIZADA'] = df_raw['PORTERIA'].astype(str).str.strip().str.lower()
            df_raw['TIPO_MARCACION'] = df_raw['PuntoMarcacion'].astype(str).str.strip().str.lower().replace({'entrada': 'ent', 'salida': 'sal'})
            df_raw.rename(columns={'COD_TRABAJADOR': 'ID_TRABAJADOR'}, inplace=True)

            # --- LÓGICA: Asignar Fecha Clave de Turno para el agrupamiento ---
            
            # **CORRECCIÓN CLAVE:** La fecha clave debe anclarse al día de la ENTRADA,
            # pero ajustando la SALIDA si es nocturna y cae antes de la hora de corte.
            def asignar_fecha_clave_turno(row):
                fecha_original = row['FECHA_HORA'].date()
                hora_marcacion = row['FECHA_HORA'].time()
                tipo_marcacion = row['TIPO_MARCACION']

                # Para ENTRADAS: La fecha clave es la fecha de la marcación misma.
                if tipo_marcacion == 'ent':
                    return fecha_original
                
                # Para SALIDAS: Si la hora es ANTES de la HORA_CORTE_NOCTURNO, 
                # asumimos que pertenece al turno que inició el día anterior.
                # Esto es crucial para agrupar (ENTRADA 21:40 del día X, SALIDA 05:40 del día X+1)
                # bajo la FECHA_CLAVE_TURNO = Día X.
                elif tipo_marcacion == 'sal' and hora_marcacion < HORA_CORTE_NOCTURNO:
                    return fecha_original - timedelta(days=1)
                
                # Si es una SALIDA después de la hora de corte (o es mediodía/tarde), 
                # la fecha clave es el día de la marcación misma. Esto capturará
                # salidas muy tardías que quizás no tienen una entrada asociada del mismo día,
                # pero el filtro de calcular_turnos las dejará pasar.
                else:
                    return fecha_original

            # Aplica la función para crear la nueva columna en el DataFrame
            df_raw['FECHA_CLAVE_TURNO'] = df_raw.apply(asignar_fecha_clave_turno, axis=1)

            st.success("Archivo cargado y preprocesado con éxito.")

            # Llama a la función principal de cálculo con el DataFrame modificado
            df_resultado = calcular_turnos(df_raw.copy(), LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS, TOLERANCIA_INFERENCIA_MINUTOS, TOLERANCIA_LLEGADA_TARDE_MINUTOS)

            if not df_resultado.empty:
                # 1. Guarda la columna original de booleanos para el formato de Excel
                df_resultado['Llegada_Tarde'] = df_resultado['Llegada_Tarde_Mas_40_Min']

                # 2. Renombra la columna 'Llegada_Tarde_Mas_40_Min' a 'Estado_Llegada'
                # y se agrega la columna de estado de cálculo
                df_resultado.rename(columns={'Llegada_Tarde_Mas_40_Min': 'Estado_Llegada'}, inplace=True)

                # 3. Mapea los valores True/False a 'Tarde'/'A tiempo' en la columna 'Estado_Llegada'
                df_resultado['Estado_Llegada'] = df_resultado['Estado_Llegada'].map({True: 'Tarde', False: 'A tiempo'})


                st.subheader("Resultados de las horas extra")
                # Se muestra el DataFrame modificado en Streamlit, sin la columna auxiliar
                st.dataframe(df_resultado.drop(columns=['Llegada_Tarde']))

                # Prepara el DataFrame para descarga en formato Excel
                buffer_excel = io.BytesIO()
                with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
                    # Exporta el DataFrame sin la columna auxiliar al Excel
                    df_to_excel = df_resultado.drop(columns=['Llegada_Tarde']).copy()
                    df_to_excel.to_excel(writer, sheet_name='Reporte Horas Extra', index=False)

                    workbook = writer.book
                    worksheet = writer.sheets['Reporte Horas Extra']

                    # Define formatos
                    orange_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                    gray_format = workbook.add_format({'bg_color': '#D9D9D9'}) # Formato para no calculados
                    
                    # Columna 'ENTRADA_REAL' para el formato condicional de tardanza
                    try:
                        entrada_real_col_idx = df_to_excel.columns.get_loc('ENTRADA_REAL')
                    except KeyError:
                        entrada_real_col_idx = -1
                        
                    # Columna 'Estado_Calculo' para el formato condicional de filas no calculadas
                    try:
                        estado_calculo_col_idx = df_to_excel.columns.get_loc('Estado_Calculo')
                    except KeyError:
                        estado_calculo_col_idx = -1


                    # Itera a través de las filas para aplicar el formato
                    for row_num, row in df_resultado.iterrows():
                        excel_row = row_num + 1 # Fila en Excel (empezando en 1 después del encabezado)
                        
                        # 1. Aplicar formato para filas NO calculadas (gris)
                        if row['Estado_Calculo'] != "Calculado":
                            # Aplica el formato gris a toda la fila
                            worksheet.set_row(excel_row, None, gray_format)
                            
                        # 2. Aplicar formato para Llegada Tarde (Naranja en ENTRADA_REAL)
                        if row['Llegada_Tarde'] and row['Estado_Calculo'] == "Calculado":
                            if entrada_real_col_idx != -1:
                                # Sobreescribe la celda de ENTRADA_REAL con el formato naranja
                                worksheet.write(excel_row, entrada_real_col_idx, row['ENTRADA_REAL'], orange_format)
                        
                        # Escribir el valor de ENTRADA_REAL si no se aplicó formato naranja (para evitar perder el valor)
                        # Nota: Si se aplicó el formato gris a toda la fila, ya se escribió en gris, así que
                        # solo se escribe sin formato si no es tarde y no se aplicó el formato gris previamente.
                        # Este bloque de código es solo para asegurar que el valor se mantenga si no está en gris
                        # y no está en naranja (la línea 251 ya lo maneja). La corrección más limpia
                        # es dejar que la iteración del dataframe de excelwriter escriba los valores,
                        # y solo sobrescribir aquellos que necesitan formato.
                        
                        # Se ajusta la lógica de sobrescritura para evitar errores de tipo en el Excel
                        # (la línea 251-255 en el código original es la que requiere corrección)
                        
                        # Se recorren todas las columnas del df_to_excel para aplicar el formato por celda,
                        # ya que set_row aplica el formato solo a las celdas vacías o a las que se escriban después.
                        
                        for col_idx, col_name in enumerate(df_to_excel.columns):
                            value = row[col_name]
                            cell_format = None
                            
                            # Si no se calculó, la fila es gris.
                            if row['Estado_Calculo'] != "Calculado":
                                cell_format = gray_format
                            
                            # Si se calculó y es llegada tarde, ENTRADA_REAL es naranja.
                            elif col_name == 'ENTRADA_REAL' and row['Llegada_Tarde']:
                                cell_format = orange_format

                            # Escribir el valor con el formato correspondiente, o sin formato
                            if col_name in ['FECHA', 'ENTRADA_REAL', 'SALIDA_REAL']:
                                # Convertir objetos datetime.date/str a datetime para el writer
                                if isinstance(value, pd.Timestamp) or isinstance(value, datetime):
                                    value_to_write = value.strftime("%Y-%m-%d %H:%M:%S")
                                elif isinstance(value, pd.Timedelta):
                                    value_to_write = round(value.total_seconds() / 3600, 2)
                                else:
                                    value_to_write = str(value)

                                worksheet.write(excel_row, col_idx, value_to_write, cell_format)
                            else:
                                worksheet.write(excel_row, col_idx, value, cell_format)


                buffer_excel.seek(0)

                # Botón de descarga para el usuario
                st.download_button(
                    label="Descargar Reporte de Horas extra (Excel)",
                    data=buffer_excel,
                    file_name="Marcación_horas_extra.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("No se encontraron marcaciones válidas en las porterías principales o el archivo estaba vacío después del filtrado.")

    except KeyError as e:
        st.error(f"Error en el procesamiento de columnas. Asegúrate de que las columnas requeridas existen: {e}")
    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}. Asegúrate de que la hoja se llama 'BaseDatos Modificada' y que las celdas de fecha y hora tienen un formato compatible.")

st.markdown("---")
st.caption("Somos NOEL DE CORAZÓN ❤️")








