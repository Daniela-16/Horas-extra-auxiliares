# -*- coding: utf-8 -*-
"""
Calculadora de Horas Extra.
MODIFICACIÓN CLAVE: La asignación de turno prioriza la MENOR DISTANCIA ABSOLUTA
al inicio programado del turno, lo cual resuelve la superposición de horarios
entre marcaciones de desplazamiento (ej. 4:00 AM) y entradas de turnos diurnos (ej. 5:40 AM).
Se mantiene la restricción estricta para la entrada del turno nocturno.
"""

import pandas as pd
from datetime import datetime, timedelta
import streamlit as st
import io
import numpy as np

# --- CÓDIGOS DE TRABAJADORES PERMITIDOS (ACTUALIZADO) ---
# Se filtra el DataFrame de entrada para incluir SOLAMENTE los registros con estos ID.
CODIGOS_TRABAJADORES_FILTRO = [
    81169, 82911, 81515, 81744, 82728, 83617, 81594, 81215, 79114, 80531,
    71329, 82383, 79143, 80796, 80795, 79830, 80584, 81131, 79110, 80530,
    82236, 82645, 80532, 71332, 82441, 79030, 81020, 82724, 82406, 81953,
    81164, 81024, 81328, 81957, 80577, 14042, 82803, 80233, 83521, 82226,
    71337381, 82631, 82725, 83309, 81947, 82385, 80765, 82642, 1128268115,
    80526, 82979, 81240, 81873, 83320, 82617, 82243, 81948, 82954
]

# --- 1. Definición de los Turnos ---

TURNOS = {
    "LV": { # Lunes a Viernes (0-4)
        "Turno 1 LV": {"inicio": "05:40:00", "fin": "13:40:00", "duracion_hrs": 8},
        "Turno 2 LV": {"inicio": "13:40:00", "fin": "21:40:00", "duracion_hrs": 8},
        "Turno 4 LV": {"inicio": "7:00:00", "fin": "17:00:00", "duracion_hrs": 10},
        # Turno nocturno: Inicia un día y termina al día siguiente
        "Turno 3 LV": {"inicio": "21:40:00", "fin": "05:40:00", "duracion_hrs": 8, "nocturno": True},
        
    },
    "SAB": { # Sábado (5)
        "Turno 1 SAB": {"inicio": "05:40:00", "fin": "11:40:00", "duracion_hrs": 6},
        "Turno 2 SAB": {"inicio": "11:40:00", "fin": "17:40:00", "duracion_hrs": 6},
        "Turno 3 SAB": {"inicio": "21:40:00", "fin": "05:40:00", "duracion_hrs": 8, "nocturno": True},
    },
    "DOM": { # Domingo (6)
        "Turno 1 DOM": {"inicio": "05:40:00", "fin": "11:40:00", "duracion_hrs": 6},
        "Turno 2 DOM": {"inicio": "11:40:00", "fin": "17:40:00", "duracion_hrs": 6},
        # Turno nocturno de Domingo: Ligeramente más tarde que los días de semana
        "Turno 3 DOM": {"inicio": "22:40:00", "fin": "05:40:00", "duracion_hrs": 7, "nocturno": True},
    }
}

# --- 2. Configuración de Puntos de Marcación ---

# PRIORITY 1: Puestos de Trabajo
LUGARES_PUESTO_TRABAJO = [
    "NOEL_MDE_OFIC_PRODUCCION_ENT", "NOEL_MDE_OFIC_PRODUCCION_SAL", "NOEL_MDE_MR_TUNEL_VIENTO_1_ENT",
    "NOEL_MDE_MR_MEZCLAS_ENT", "NOEL_MDE_ING_MEN_CREMAS_ENT", "NOEL_MDE_ING_MEN_CREMAS_SAL",
    "NOEL_MDE_MR_HORNO_6-8-9_ENT", "NOEL_MDE_MR_SERVICIOS_2_ENT", "NOEL_MDE_RECURSOS_HUMANOS_ENT",
    "NOEL_MDE_RECURSOS_HUMANOS_SAL", "NOEL_MDE_ESENCIAS_2_SAL", "NOEL_MDE_ESENCIAS_1_SAL",
    "NOEL_MDE_ING_MENORES_2_ENT", "NOEL_MDE_MR_HORNO_18_ENT", "NOEL_MDE_MR_WAFER_RCH_CREMAS_ENT",
    "NOEL_MDE_MR_HORNO_6-8-9_SAL", "NOEL_MDE_TORNIQUETE_SORTER_ENT", "NOEL_MDE_TORNIQUETE_SORTER_SAL",
    "NOEL_MDE_MR_MEZCLAS_SAL", "NOEL_MDE_MR_TUNEL_VIENTO_2_ENT", "NOEL_MDE_MR_HORNO_7-10_ENT",
    "NOEL_MDE_MR_HORNO_11_ENT", "NOEL_MDE_MR_WAFER_RCH_CREMAS_SAL", "NOEL_MDE_MR_HORNO_2-4-5_SAL",
    "NOEL_MDE_MR_HORNO_4-5_ENT", "NOEL_MDE_MR_HORNO_18_SAL", "NOEL_MDE_MR_HORNO_1-3_SAL",
    "NOEL_MDE_MR_HORNO_1-3_ENT", "NOEL_MDE_CONTROL_BUHLER_ENT", "NOEL_MDE_CONTROL_BUHLER_SAL",
    "NOEL_MDE_ING_MEN_ALERGENOS_ENT", "NOEL_MDE_ING_MENORES_2_SAL", "NOEL_MDE_MR_SERVICIOS_2_SAL",
    "NOEL_MDE_MR_HORNO_11_SAL", "NOEL_MDE_MR_HORNO_7-10_SAL", "NOEL_MDE_MR_HORNO_2-12_ENT",
    "NOEL_MDE_TORNIQUETE_PATIO_SAL", "NOEL_MDE_TORNIQUETE_PATIO_ENT", "NOEL_MDE_ESENCIAS_1_ENT",
    "NOEL_MDE_ING_MENORES_1_SAL", "NOEL_MDE_MOLINETE_BODEGA_EXT_SAL", "NOEL_MDE_PRINCIPAL_ENT",
    "NOEL_MDE_ING_MENORES_1_ENT", "NOEL_MDE_MR_HORNOS_SAL", "NOEL_MDE_MR_HORNO_6-8-9_SAL_2",
    "NOEL_MDE_PRINCIPAL_SAL", "NOEL_MDE_MR_ASPIRACION_ENT", "NOEL_MDE_MR_HORNO_2-12_SAL",
    "NOEL_MDE_MR_HORNOS_ENT", "NOEL_MDE_MR_HORNO_4-5_SAL", "NOEL_MDE_ING_MEN_ALERGENOS_SAL",
    "NOEL_MDE_MR_WAFER_RCH_CREMAS_ENT", "NOEL_MDE_MR_WAFER_RCH_CREMAS_SAL",
    "NOEL_MDE_MR_MEZCLAS_ENT", "NOEL_MDE_OFIC_PRODUCCION_SAL", "NOEL_MDE_OFIC_PRODUCCION_ENT", "NOEL_MDE_MR_MEZCLAS_ENT"
]

# PRIORITY 2: Porterías
LUGARES_PORTERIA = [
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

LUGARES_PUESTO_TRABAJO_NORMALIZADOS = [lugar.strip().lower() for lugar in LUGARES_PUESTO_TRABAJO]
LUGARES_PORTERIA_NORMALIZADOS = [lugar.strip().lower() for lugar in LUGARES_PORTERIA]
LUGARES_COMBINADOS_NORMALIZADOS = LUGARES_PUESTO_TRABAJO_NORMALIZADOS + LUGARES_PORTERIA_NORMALIZADOS


MAX_EXCESO_SALIDA_HRS = 3
HORA_CORTE_NOCTURNO = datetime.strptime("08:00:00", "%H:%M:%S").time()

# --- CONSTANTES DE TOLERANCIA ---
TOLERANCIA_LLEGADA_TARDE_MINUTOS = 40
TOLERANCIA_ENTRADA_TEMPRANA_MINUTOS = 360 # 6 horas de adelanto
TOLERANCIA_ASIGNACION_TARDE_MINUTOS = 180 # 3 horas de margen para la asignación
UMBRAL_PAGO_ENTRADA_TEMPRANA_MINUTOS = 30
MIN_DURACION_ACEPTABLE_REAL_SALIDA_HRS = 1
UMBRAL_HORAS_EXTRA_RESALTAR = 30 / 60 

# --- 3. Obtener turno basado en fecha y hora ---

def buscar_turnos_posibles(fecha_clave: datetime.date):
    """Genera una lista de (nombre_turno, info, inicio_dt, fin_dt, fecha_clave_asignada) para un día."""
    dia_semana_clave = fecha_clave.weekday()

    if dia_semana_clave < 5: tipo_dia = "LV"
    elif dia_semana_clave == 5: tipo_dia = "SAB"
    else: tipo_dia = "DOM"

    turnos_dia = []
    if tipo_dia in TURNOS:
        for nombre_turno, info_turno in TURNOS[tipo_dia].items():
            hora_inicio = datetime.strptime(info_turno["inicio"], "%H:%M:%S").time()
            hora_fin = datetime.strptime(info_turno["fin"], "%H:%M:%S").time()
            es_nocturno = info_turno.get("nocturno", False)

            inicio_posible_turno = datetime.combine(fecha_clave, hora_inicio)

            if es_nocturno:
                fin_posible_turno = datetime.combine(fecha_clave + timedelta(days=1), hora_fin)
            else:
                fin_posible_turno = datetime.combine(fecha_clave, hora_fin)

            turnos_dia.append((nombre_turno, info_turno, inicio_posible_turno, fin_posible_turno, fecha_clave))
    return turnos_dia

def obtener_turno_para_registro(fecha_hora_evento: datetime, fecha_clave_turno_reporte: datetime.date):
    """
    Busca el turno programado más cercano a la marcación de entrada.
    Prioriza el turno con la menor distancia absoluta a la hora de inicio programada,
    resolviendo la ambigüedad de marcaciones de desplazamiento.
    
    Retorna: (nombre, info, inicio_turno, fin_turno, fecha_clave_final)
    """
    mejor_turno_data = None
    mejor_distancia_a_inicio = timedelta.max  

    # --- 1. Generar Candidatos de Turno (Día X y Día X-1) ---
    turnos_candidatos = buscar_turnos_posibles(fecha_clave_turno_reporte)
    hora_evento = fecha_hora_evento.time()
    
    # Si la marcación es temprano en la mañana, añadir candidatos nocturnos del día anterior (Día X - 1)
    if hora_evento < HORA_CORTE_NOCTURNO:
        fecha_clave_anterior = fecha_clave_turno_reporte - timedelta(days=1)
        turnos_candidatos.extend(buscar_turnos_posibles(fecha_clave_anterior))

    
    for nombre_turno, info_turno, inicio_posible_turno, fin_posible_turno, fecha_clave_asignada in turnos_candidatos:

        es_nocturno = info_turno.get("nocturno", False)
        
        # 1. El límite más temprano que aceptamos la entrada (6 horas antes por defecto)
        rango_inicio_temprano = inicio_posible_turno - timedelta(minutes=TOLERANCIA_ENTRADA_TEMPRANA_MINUTOS)
        
        # 2. El límite más tardío que aceptamos la entrada (Tolerancia amplia por defecto, 3 horas)
        rango_fin_tarde = inicio_posible_turno + timedelta(minutes=TOLERANCIA_ASIGNACION_TARDE_MINUTOS + 5)
        
        
        # --- RESTRICCIÓN CLAVE: EVITAR MARCADAS DE MADRUGADA COMO ENTRADA NOCTURNA (22:40) ---
        # Si es nocturno (T3), solo aceptamos entradas hasta 60 minutos después de la hora programada.
        # Esto filtra marcaciones de 2:00 AM que se confunden con el inicio de 22:40.
        if es_nocturno:
             rango_fin_tarde = inicio_posible_turno + timedelta(minutes=60)


        # Validar si el evento (la entrada) cae en esta ventana
        if fecha_hora_evento >= rango_inicio_temprano and fecha_hora_evento <= rango_fin_tarde:
            
            # Métrica de prioridad: la menor DISTANCIA ABSOLUTA al inicio programado.
            distancia_a_inicio = abs(fecha_hora_evento - inicio_posible_turno)

            # Priorización: Si es el primer turno encontrado O si es más cercano al inicio programado
            if mejor_turno_data is None or distancia_a_inicio < mejor_distancia_a_inicio:
                
                # Usamos <= para mantener la marcación más temprana si hay empates
                if distancia_a_inicio <= mejor_distancia_a_inicio:
                    mejor_turno_data = (nombre_turno, info_turno, inicio_posible_turno, fin_posible_turno, fecha_clave_asignada)
                    mejor_distancia_a_inicio = distancia_a_inicio 

    return mejor_turno_data if mejor_turno_data else (None, None, None, None, None)

# --- 4. Calculo de horas ---

def calcular_turnos(df: pd.DataFrame, lugares_puesto: list, lugares_porteria: list, tolerancia_llegada_tarde: int):
    """
    Agrupa por ID y FECHA_CLAVE_TURNO.
    Busca el turno priorizando las marcaciones de Puesto de Trabajo sobre Portería.
    """
    
    df_filtrado = df[(df['TIPO_MARCACION'].isin(['ent', 'sal']))].copy()
    df_filtrado.sort_values(by=['id_trabajador', 'FECHA_HORA'], inplace=True)

    if df_filtrado.empty: return pd.DataFrame()

    resultados = []

    for (id_trabajador, fecha_clave_turno), grupo in df_filtrado.groupby(['id_trabajador', 'FECHA_CLAVE_TURNO']):

        nombre = grupo['nombre'].iloc[0]
        entradas = grupo[grupo['TIPO_MARCACION'] == 'ent']
        
        entrada_real = pd.NaT
        porteria_entrada = 'N/A'
        salida_real = pd.NaT
        porteria_salida = 'N/A'
        turno_nombre, info_turno, inicio_turno, fin_turno, fecha_clave_final = (None, None, None, None, fecha_clave_turno)
        horas_trabajadas = 0.0
        horas_extra = 0.0
        llegada_tarde_flag = False
        estado_calculo = "Sin Marcaciones Válidas (E/S)"
        salida_fue_real = False 
        es_nocturno_flag = False 
        
        mejor_entrada_para_turno = pd.NaT
        mejor_turno_data = (None, None, None, None, None)

        # --- A. PRIORIDAD 1: Buscar Turno con Marcaciones de Puesto de Trabajo ---
        entradas_puesto = entradas[entradas['PORTERIA_NORMALIZADA'].isin(lugares_puesto)]
        
        if not entradas_puesto.empty:
            mejor_distancia_a_inicio = timedelta.max
            for index, row in entradas_puesto.iterrows():
                current_entry_time = row['FECHA_HORA']
                turno_data = obtener_turno_para_registro(current_entry_time, fecha_clave_turno)
                
                if turno_data[0] is not None:
                    inicio_programado = turno_data[2]
                    distancia = abs(current_entry_time - inicio_programado)
                    
                    if distancia < mejor_distancia_a_inicio:
                        mejor_distancia_a_inicio = distancia
                        mejor_entrada_para_turno = current_entry_time
                        mejor_turno_data = turno_data

        # --- B. PRIORIDAD 2: Buscar Turno con Marcaciones de Portería (Solo si no se encontró en Puesto) ---
        if mejor_turno_data[0] is None:
            entradas_porteria = entradas[entradas['PORTERIA_NORMALIZADA'].isin(lugares_porteria)]
            
            if not entradas_porteria.empty:
                mejor_distancia_a_inicio = timedelta.max
                for index, row in entradas_porteria.iterrows():
                    current_entry_time = row['FECHA_HORA']
                    turno_data = obtener_turno_para_registro(current_entry_time, fecha_clave_turno)
                    
                    if turno_data[0] is not None:
                        inicio_programado = turno_data[2]
                        distancia = abs(current_entry_time - inicio_programado)
                        
                        if distancia < mejor_distancia_a_inicio:
                            mejor_distancia_a_inicio = distancia
                            mejor_entrada_para_turno = current_entry_time
                            mejor_turno_data = turno_data
        
        # --- C. Asignación y Cálculo Final ---
        if pd.notna(mejor_entrada_para_turno):
            entrada_real = mejor_entrada_para_turno
            turno_nombre, info_turno, inicio_turno, fin_turno, fecha_clave_final = mejor_turno_data
            es_nocturno_flag = info_turno.get("nocturno", False)
            
            porteria_entrada = grupo[grupo['FECHA_HORA'] == entrada_real]['porteria'].iloc[0]
            
            # --- Inferencia de Salida ---
            max_salida_aceptable = fin_turno + timedelta(hours=MAX_EXCESO_SALIDA_HRS)
            
            valid_salidas = grupo[
                (grupo['TIPO_MARCACION'] == 'sal') &
                (grupo['FECHA_HORA'] > entrada_real) &
                (grupo['FECHA_HORA'] <= max_salida_aceptable)
            ]
            
            if valid_salidas.empty:
                salida_real = fin_turno
                porteria_salida = 'ASUMIDA (Falta Salida/Salida Inválida)'
                estado_calculo = "ASUMIDO (Falta Salida/Salida Inválida)"
                salida_fue_real = False
            else:
                salida_real = valid_salidas['FECHA_HORA'].max()
                porteria_salida = valid_salidas[valid_salidas['FECHA_HORA'] == salida_real]['porteria'].iloc[0]
                estado_calculo = "Calculado"
                salida_fue_real = True
                
            # --- Para Micro-jornadas ---
            if salida_fue_real:
                duracion_check = salida_real - entrada_real
                if duracion_check < timedelta(hours=MIN_DURACION_ACEPTABLE_REAL_SALIDA_HRS):
                    salida_real = fin_turno
                    porteria_salida = 'ASUMIDA (Micro-jornada detectada)'
                    estado_calculo = "ASUMIDO (Micro-jornada detectada)"
                    salida_fue_real = False

            # --- Reglas de Cálculo de Horas ---
            inicio_efectivo_calculo = inicio_turno
            llegada_tarde_flag = False
            
            # 1. Regla para LLEGADA TARDE
            if entrada_real > inicio_turno + timedelta(minutes=tolerancia_llegada_tarde):
                inicio_efectivo_calculo = entrada_real
                llegada_tarde_flag = True
                
            # 2. Regla para ENTRADA TEMPRANA
            elif entrada_real < inicio_turno:
                early_timedelta = inicio_turno - entrada_real
                
                if early_timedelta > timedelta(minutes=UMBRAL_PAGO_ENTRADA_TEMPRANA_MINUTOS):
                    inicio_efectivo_calculo = entrada_real
                else:
                    inicio_efectivo_calculo = inicio_turno
            
            duracion_efectiva_calculo = salida_real - inicio_efectivo_calculo

            if duracion_efectiva_calculo < timedelta(seconds=0):
                horas_trabajadas = 0.0
                horas_extra = 0.0
                estado_calculo = "Error: Duración efectiva negativa"
            else:
                horas_trabajadas = round(duracion_efectiva_calculo.total_seconds() / 3600, 2)
                
                horas_turno = info_turno["duracion_hrs"]
                horas_extra = max(0, round(horas_trabajadas - horas_turno, 2))

        else:
            estado_calculo = "Turno No Asignado (Ninguna marcación se alinea con un turno programado)"

        if pd.isna(entrada_real) and not grupo[grupo['TIPO_MARCACION'] == 'sal'].empty:
            continue
            
        # --- Añade los resultados a la lista (Se reporta todo) ---
        ent_str = entrada_real.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(entrada_real) else 'N/A'
        sal_str = salida_real.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(salida_real) else 'N/A'
        report_date = fecha_clave_final if fecha_clave_final else fecha_clave_turno
        inicio_str = inicio_turno.time().strftime("%H:%M:%S") if inicio_turno else 'N/A'
        fin_str = fin_turno.time().strftime("%H:%M:%S") if fin_turno else 'N/A'
        horas_turno_val = info_turno["duracion_hrs"] if info_turno else 0

        resultados.append({
            'NOMBRE': nombre,
            'ID_TRABAJADOR': id_trabajador,
            'FECHA': report_date,
            'Dia_Semana': report_date.strftime('%A'),
            'TURNO': turno_nombre if turno_nombre else 'N/A',
            'Inicio_Turno_Programado': inicio_str,
            'Fin_Turno_Programado': fin_str,
            'Duracion_Turno_Programado_Hrs': horas_turno_val,
            'ENTRADA_REAL': ent_str,
            'PORTERIA_ENTRADA': porteria_entrada,
            'SALIDA_REAL': sal_str,
            'PORTERIA_SALIDA': porteria_salida,
            'Horas_Trabajadas_Netas': horas_trabajadas,
            'Horas_Extra': horas_extra,
            'Horas': int(horas_extra),
            'Minutos': round((horas_extra - int(horas_extra)) * 60),
            'Llegada_Tarde_Mas_40_Min': llegada_tarde_flag,
            'Estado_Calculo': estado_calculo,
            'Es_Nocturno': es_nocturno_flag 
        })

    return pd.DataFrame(resultados)

# --- 5. Nueva Función de Filtrado Post-Cálculo (Filtro de Días Extremos) ---

def aplicar_filtro_primer_ultimo_dia(df_resultado):
    """
    Aplica el filtro para conservar el primer y último día solo si cumplen
    con la condición horaria de marcación de un turno nocturno (entrada ~22:40, salida ~5:40).
    Los días intermedios siempre se conservan.
    """
    if df_resultado.empty:
        return df_resultado

    df_filtrado = df_resultado.copy()
    rows_to_keep_indices = []
    
    df_filtrado['FECHA_DATE'] = pd.to_datetime(df_filtrado['FECHA']).dt.date
    df_filtrado['ENTRADA_DT'] = pd.to_datetime(df_filtrado['ENTRADA_REAL'], errors='coerce')
    df_filtrado['SALIDA_DT'] = pd.to_datetime(df_filtrado['SALIDA_REAL'], errors='coerce')


    # 1. Iterar por cada trabajador para aplicar la lógica individualmente
    for id_trabajador, df_worker_group in df_filtrado.groupby('ID_TRABAJADOR'):
        
        df_worker = df_worker_group.sort_values(by='FECHA_DATE').copy()
        unique_dates = df_worker['FECHA_DATE'].unique()
        
        if len(unique_dates) == 0:
            continue
            
        first_day = unique_dates[0]
        last_day = unique_dates[-1]

        for current_date in unique_dates:
            
            current_day_turnos = df_worker[df_worker['FECHA_DATE'] == current_date].copy()
            
            # --- Regla A: Días Intermedios (No son ni el primero ni el último) ---
            if current_date > first_day and current_date < last_day:
                rows_to_keep_indices.extend(current_day_turnos.index.tolist())
                continue
                
            
            # --- Regla B: Primer Día (Entrada Nocturna: 21:00 PM - 23:59 PM) ---
            if current_date == first_day:
                
                limite_min_entrada = datetime.combine(current_date, datetime.strptime("21:00:00", "%H:%M:%S").time())
                limite_max_entrada = datetime.combine(current_date, datetime.strptime("23:59:59", "%H:%M:%S").time())
                
                primer_dia_nocturno_valido = current_day_turnos[
                    (current_day_turnos['Es_Nocturno'] == True) &
                    (current_day_turnos['ENTRADA_DT'] >= limite_min_entrada) &
                    (current_day_turnos['ENTRADA_DT'] <= limite_max_entrada)
                ]

                if not primer_dia_nocturno_valido.empty:
                    rows_to_keep_indices.extend(current_day_turnos.index.tolist())
            
            
            # --- Regla C: Último Día (Salida Nocturna: 05:00 AM - 07:00 AM) ---
            if current_date == last_day and current_date != first_day:
                
                limite_min_salida = datetime.combine(current_date, datetime.strptime("05:00:00", "%H:%M:%S").time())
                limite_max_salida = datetime.combine(current_date, datetime.strptime("07:00:00", "%H:%M:%S").time())
                
                ultimo_dia_nocturno_valido = current_day_turnos[
                    (current_day_turnos['Es_Nocturno'] == True) &
                    (current_day_turnos['SALIDA_DT'] >= limite_min_salida) &
                    (current_day_turnos['SALIDA_DT'] <= limite_max_salida)
                ]

                if not ultimo_dia_nocturno_valido.empty:
                    rows_to_keep_indices.extend(current_day_turnos.index.tolist())


    # Filtrar el DataFrame original por los índices conservados y eliminar las columnas temporales
    df_final = df_resultado.loc[rows_to_keep_indices].copy()
    df_final.drop(columns=['Es_Nocturno', 'FECHA_DATE', 'ENTRADA_DT', 'SALIDA_DT'], inplace=True, errors='ignore')
    return df_final


# --- 6. Interfaz Streamlit ---

st.set_page_config(page_title="Calculadora de Horas Extra", layout="wide")
st.title("📊 Calculadora de Horas Extra - NOEL")
st.write("Sube tu archivo de Excel para calcular las horas extra del personal. **Nota Importante:** El primer y último día del reporte solo se incluyen si cumplen las condiciones de marcación del turno nocturno (Entrada ~22:40, Salida ~05:40).")
st.caption("La asignación de turno prioriza la **menor desviación horaria** al inicio programado del turno.")

archivo_excel = st.file_uploader("Sube un archivo Excel (.xlsx)", type=["xlsx"])

if archivo_excel is not None:
    try:
        # Intenta leer la hoja específica 
        try:
            df_raw = pd.read_excel(archivo_excel, sheet_name='data')
        except ValueError:
            df_raw = pd.read_excel(archivo_excel, sheet_name='BaseDatos Modificada')


        # 1. Definir la lista de nombres de columna que esperamos
        columnas_requeridas_lower = [
            'cc', 'codtrabajador', 'nombre', 'fecha', 'hora', 'porteria', 'puntomarcacion'
        ]
        
        # 2. Renombrar columnas a minúsculas
        col_map = {col: col.lower() for col in df_raw.columns}
        df_raw.rename(columns=col_map, inplace=True)

        # 3. Validar columnas
        if not all(col in df_raw.columns for col in columnas_requeridas_lower):
            st.error(f"⚠️ ERROR: Faltan columnas requeridas o tienen nombres incorrectos. Asegúrate de tener: **Cc, CodTrabajador, Nombre, Fecha, Hora, Porteria, PuntoMarcacion**.")
            st.stop()

        # 4. Preparar DataFrame
        df_raw = df_raw[columnas_requeridas_lower].copy()
        df_raw.rename(columns={'codtrabajador': 'id_trabajador'}, inplace=True)
        
        # --- FILTRADO POR CÓDIGO DE TRABAJADOR ---
        try:
            df_raw['id_trabajador'] = pd.to_numeric(df_raw['id_trabajador'], errors='coerce').astype('Int64')
        except:
            df_raw['id_trabajador'] = df_raw['id_trabajador'].astype(str)
            codigos_filtro = [str(c) for c in CODIGOS_TRABAJADORES_FILTRO]
        else:
            codigos_filtro = CODIGOS_TRABAJADORES_FILTRO

        df_raw = df_raw[df_raw['id_trabajador'].isin(codigos_filtro)].copy()
        
        if df_raw.empty:
            st.error("⚠️ ERROR: Después del filtrado por código de trabajador, no quedan registros para procesar.")
            st.stop()
        
        # Preprocesamiento de Fecha
        df_raw['fecha'] = pd.to_datetime(df_raw['fecha'], errors='coerce')  
        df_raw.dropna(subset=['fecha'], inplace=True)
        
        # --- Función para estandarizar el formato de la hora ---
        def standardize_time_format(time_val):
            if isinstance(time_val, float) and time_val <= 1.0: 
                total_seconds = int(time_val * 86400)
                hours, remainder = divmod(total_seconds, 3600)
                minutes, seconds = divmod(remainder, 60)
                return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
            time_str = str(time_val)
            parts = time_str.split(':')
            if len(parts) == 2:
                return f"{time_str}:00"
            elif len(parts) == 3:
                return time_str
            else:
                return '00:00:00'  # Retorno seguro

        # Combinar FECHA y HORA
        df_raw['hora'] = df_raw['hora'].apply(standardize_time_format)
        
        try:
            df_raw['FECHA_HORA'] = pd.to_datetime(df_raw['fecha'].dt.strftime('%Y-%m-%d') + ' ' + df_raw['hora'], errors='coerce')
            df_raw.dropna(subset=['FECHA_HORA'], inplace=True)
        except Exception as e:
            st.error(f"Error al combinar FECHA y HORA: {e}")
            st.stop() 

        # Normalización y Tipo de Marcación
        df_raw['PORTERIA_NORMALIZADA'] = df_raw['porteria'].astype(str).str.strip().str.lower()
        df_raw['TIPO_MARCACION'] = df_raw['puntomarcacion'].astype(str).str.strip().str.lower().replace({'entrada': 'ent', 'salida': 'sal'})

        # --- Función para asignar Fecha Clave de Turno ---
        def asignar_fecha_clave_turno_corregida(row):
            fecha_original = row['FECHA_HORA'].date()
            hora_marcacion = row['FECHA_HORA'].time()
            tipo_marcacion = row['TIPO_MARCACION']
            
            if tipo_marcacion == 'ent':
                return fecha_original
            
            if tipo_marcacion == 'sal' and hora_marcacion < HORA_CORTE_NOCTURNO:
                return fecha_original - timedelta(days=1)
            
            return fecha_original

        df_raw['FECHA_CLAVE_TURNO'] = df_raw.apply(asignar_fecha_clave_turno_corregida, axis=1)
        
        # Filtrado Final del dataset crudo
        df_raw_filtrado = df_raw[
            (df_raw['PORTERIA_NORMALIZADA'].isin(LUGARES_COMBINADOS_NORMALIZADOS)) & 
            (df_raw['TIPO_MARCACION'].isin(['ent', 'sal']))
        ].copy()

        st.success(f"✅ Archivo cargado y preprocesado con éxito. Se encontraron {len(df_raw_filtrado['FECHA_CLAVE_TURNO'].unique())} días de jornada para procesar de {len(df_raw_filtrado['id_trabajador'].unique())} trabajadores filtrados.")

        # --- Ejecutar el Cálculo ---
        df_resultado = calcular_turnos(
            df_raw_filtrado, 
            LUGARES_PUESTO_TRABAJO_NORMALIZADOS, 
            LUGARES_PORTERIA_NORMALIZADOS, 
            TOLERANCIA_LLEGADA_TARDE_MINUTOS
        )

        if not df_resultado.empty:
            
            # --- APLICAR EL NUEVO FILTRO DE PRIMER Y ÚLTIMO DÍA ---
            df_resultado_filtrado = aplicar_filtro_primer_ultimo_dia(df_resultado)
            
            if df_resultado_filtrado.empty:
                st.warning("No se encontraron jornadas válidas después de aplicar los filtros de primer/último día.")
                st.stop()
                
            # Post-procesamiento para el reporte
            df_resultado_filtrado['Estado_Llegada'] = df_resultado_filtrado['Llegada_Tarde_Mas_40_Min'].map({True: 'Tarde', False: 'A tiempo'})
            df_resultado_filtrado.sort_values(by=['NOMBRE', 'FECHA', 'ENTRADA_REAL'], inplace=True) 
            
            columnas_reporte = [
                'NOMBRE', 'ID_TRABAJADOR', 'FECHA', 'Dia_Semana', 'TURNO',
                'Inicio_Turno_Programado', 'Fin_Turno_Programado', 'Duracion_Turno_Programado_Hrs',
                'ENTRADA_REAL', 'PORTERIA_ENTRADA', 'SALIDA_REAL', 'PORTERIA_SALIDA',
                'Horas_Trabajadas_Netas', 'Horas_Extra', 'Horas', 'Minutos', 
                'Estado_Llegada', 'Estado_Calculo'
            ]

            st.subheader("Resultados de las Horas Extra")
            st.dataframe(df_resultado_filtrado[columnas_reporte], use_container_width=True)

            # --- Lógica de descarga en Excel con formato condicional ---
            buffer_excel = io.BytesIO()
            with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
                df_to_excel = df_resultado_filtrado[columnas_reporte].copy()
                df_to_excel.to_excel(writer, sheet_name='Reporte Horas Extra', index=False)

                workbook = writer.book
                worksheet = writer.sheets['Reporte Horas Extra']

                # Formatos de Excel
                orange_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}) 
                gray_format = workbook.add_format({'bg_color': '#D9D9D9'}) 
                yellow_format = workbook.add_format({'bg_color': '#FFF2CC', 'font_color': '#3C3C3C'}) 
                red_extra_format = workbook.add_format({'bg_color': '#F8E8E8', 'font_color': '#D83A56', 'bold': True})
                
                # Aplicación de formatos condicionales
                for row_num, row in df_resultado_filtrado.iterrows():
                    try:
                        excel_row = df_to_excel.index.get_loc(row_num) + 1 
                    except KeyError:
                        continue
                        
                    is_late = row['Llegada_Tarde_Mas_40_Min']
                    is_assumed = row['Estado_Calculo'].startswith("ASUMIDO")
                    is_missing_entry = row['Estado_Calculo'].startswith("Sin Marcaciones Válidas") or row['Estado_Calculo'].startswith("Turno No Asignado")
                    is_excessive_extra = row['Horas_Extra'] > UMBRAL_HORAS_EXTRA_RESALTAR

                    base_format = None
                    if is_missing_entry and not is_assumed:
                        base_format = gray_format
                    elif is_assumed:
                        base_format = yellow_format

                    for col_idx, col_name in enumerate(df_to_excel.columns):
                        value = row[col_name]
                        cell_format = base_format 
                        
                        if col_name == 'ENTRADA_REAL' and is_late:
                            cell_format = orange_format
                        
                        if is_excessive_extra and col_name in ['Horas_Extra', 'Horas', 'Minutos']:
                            cell_format = red_extra_format

                        worksheet.write(excel_row, col_idx, value if pd.notna(value) else 'N/A', cell_format)

                # Ajustar el ancho de las columnas
                for i, col in enumerate(df_to_excel.columns):
                    max_len = max(df_to_excel[col].astype(str).str.len().max(), len(col)) + 2
                    worksheet.set_column(i, i, max_len)

            buffer_excel.seek(0)

            st.download_button(
                label="Descargar Reporte de Horas Extra (Excel)",
                data=buffer_excel,
                file_name="Reporte_Marcacion_Horas_Extra_Filtrado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("No se encontraron jornadas válidas después de aplicar los filtros.")

    except KeyError as e:
        if "'data'" in str(e) or "'BaseDatos Modificada'" in str(e):
            st.error(f"⚠️ ERROR: El archivo Excel debe contener una hoja llamada **'data'** o **'BaseDatos Modificada'** y las columnas requeridas.")
        else:
            st.error(f"⚠️ ERROR: Faltan columnas requeridas o tienen nombres incorrectos: {e}")
    except Exception as e:
        st.error(f"Error crítico al procesar el archivo: {e}. Por favor, verifica el formato de los datos.")

st.markdown("---")
st.caption("Somos NOEL DE CORAZÓN ❤️ - Herramienta de Cálculo de Turnos y Horas Extra")





























