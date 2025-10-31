# -*- coding: utf-8 -*-
"""
Calculadora de Horas Extra.
"""

import pandas as pd
from datetime import datetime, timedelta
import streamlit as st
import io
import numpy as np

# --- CÓDIGOS DE TRABAJADORES PERMITIDOS (Agregado para el filtrado) ---
# Se filtra el DataFrame de entrada para incluir SOLAMENTE los registros con estos ID.
CODIGOS_TRABAJADORES_FILTRO = [
    81169, 82911, 81515, 81744, 82728, 83617, 81594, 81215, 79114, 80531, 
    71329, 82383, 79143, 80796, 80795, 79830, 80584, 81131, 79110, 80530, 
    82236, 82645, 80532, 71332, 82441, 79030, 81020, 82724, 82406, 81953, 
    81164, 81024, 81328, 81957, 80577, 14042, 82803, 80233, 83521, 82226, 
    71337381, 82631, 82725, 83309, 81947, 82385, 80765, 82642, 1128268115, 
    80526, 82979, 81240, 81873, 83320, 82617, 82243, 81948, 82954
]
# Se asegura que la lista de códigos sea de tipo entero para la comparación.

# --- 1. Definición de los Turnos ---

TURNOS = {
    "LV": { # Lunes a Viernes (0-4)
        "Turno 1 LV": {"inicio": "05:40:00", "fin": "13:40:00", "duracion_hrs": 8},
        "Turno 2 LV": {"inicio": "13:40:00", "fin": "21:40:00", "duracion_hrs": 8},
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

# --- 2. Configuración General ---

# Lista de porterías/lugares considerados como válidos para Entrada/Salida de jornada

LUGARES_TRABAJO_PRINCIPAL = [
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
    "NOEL_MDE_MR_MEZCLAS_ENT", "NOEL_MDE_OFIC_PRODUCCION_SAL", "NOEL_MDE_OFIC_PRODUCCION_ENT", "NOEL_MDE_MR_MEZCLAS_ENT",
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

LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS = [lugar.strip().lower() for lugar in LUGARES_TRABAJO_PRINCIPAL]

# Máximo de horas después del fin de turno programado que se acepta una salida como válida.
MAX_EXCESO_SALIDA_HRS = 3
# Hora de corte para definir si una SALIDA en la mañana pertenece al turno del día anterior (ej: 08:00 AM)
HORA_CORTE_NOCTURNO = datetime.strptime("08:00:00", "%H:%M:%S").time()

# --- CONSTANTES DE TOLERANCIA REVISADAS ---
# Tolerancia para considerar la llegada como 'tarde' para el cálculo de horas. 
TOLERANCIA_LLEGADA_TARDE_MINUTOS = 40

# Tolerancia MÁXIMA para considerar la llegada como 'temprana' para la asignación de turno.

TOLERANCIA_ENTRADA_TEMPRANA_MINUTOS = 180 

# --- HORAS EXTRA LLEGADA TEMPRANO ---
# Umbral de tiempo (en minutos) para determinar si la llegada temprana se paga desde la hora real.
# Si la antelación es > 30 minutos, se cuenta desde la entrada real. Si es <= 30 minutos, se cuenta desde el inicio programado.
UMBRAL_PAGO_ENTRADA_TEMPRANA_MINUTOS = 30 # 30 minutos

# --- EVITAR MICRO-JORNADAS ---
# Si la duración es menor a este umbral (ej: 1 hora) y se usó una SALIDA REAL, se ignora esa salida

MIN_DURACION_ACEPTABLE_REAL_SALIDA_HRS = 1

# ---HORA EXTRA MAS DE 30 MIN ---
# Umbral en horas para resaltar las Horas Extra (30 minutos / 60 minutos = 0.5)
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
                # Si es nocturno, el fin del turno ocurre al día siguiente
                fin_posible_turno = datetime.combine(fecha_clave + timedelta(days=1), hora_fin)
            else:
                fin_posible_turno = datetime.combine(fecha_clave, hora_fin)

            # (nombre, info, inicio_dt, fin_dt, fecha_clave_asignada)
            turnos_dia.append((nombre_turno, info_turno, inicio_posible_turno, fin_posible_turno, fecha_clave))
    return turnos_dia

def obtener_turno_para_registro(fecha_hora_evento: datetime, fecha_clave_turno_reporte: datetime.date):
    """
    Busca el turno programado más cercano a la marcación de entrada,
    verificando los turnos que inician en la FECHA_CLAVE_TURNO y,
    si es temprano en la mañana, también los nocturnos del día anterior.

    Se ha modificado la lógica para **priorizar la entrada más temprana**
    entre todos los turnos posibles que caen dentro de la ventana de aceptación.

    Retorna: (nombre, info, inicio_turno, fin_turno, fecha_clave_final)
    """
    mejor_turno_data = None
    # CAMBIO CLAVE 1: Usamos 'datetime.max' para asegurar que la primera entrada válida sea la mejor
    mejor_hora_entrada = datetime.max 

    # Candidatos a turno para el día de la FECHA CLAVE (Día X)
    turnos_candidatos = buscar_turnos_posibles(fecha_clave_turno_reporte)

    # Si la marcación es temprano en la mañana, añadir candidatos nocturnos del día anterior (Día X - 1)
    hora_evento = fecha_hora_evento.time()
    if hora_evento < HORA_CORTE_NOCTURNO:
        fecha_clave_anterior = fecha_clave_turno_reporte - timedelta(days=1)
        turnos_candidatos.extend(buscar_turnos_posibles(fecha_clave_anterior))

    for nombre_turno, info_turno, inicio_posible_turno, fin_posible_turno, fecha_clave_asignada in turnos_candidatos:

        # --- LÓGICA DE RESTRICCIÓN DE VENTANA DE ENTRADA (Mantenida) ---
        # 1. El límite más temprano que aceptamos la entrada (3 horas antes = 180 minutos)
        rango_inicio_temprano = inicio_posible_turno - timedelta(minutes=TOLERANCIA_ENTRADA_TEMPRANA_MINUTOS)
        
        # 2. El límite más tardío que aceptamos la entrada (45 minutos después del inicio programado: 40 + 5 min buffer)
        rango_fin_tarde = inicio_posible_turno + timedelta(minutes=TOLERANCIA_LLEGADA_TARDE_MINUTOS + 5)
        
        # Validar si el evento (la entrada) cae en esta ventana estricta alrededor del INICIO PROGRAMADO.
        if fecha_hora_evento >= rango_inicio_temprano and fecha_hora_evento <= rango_fin_tarde:
            
            current_entry_time = fecha_hora_evento

            # CAMBIO CLAVE 2 y 3: Priorizar la hora de entrada MÁS PEQUEÑA (más temprana)
            if mejor_turno_data is None or current_entry_time < mejor_hora_entrada:
                mejor_turno_data = (nombre_turno, info_turno, inicio_posible_turno, fin_posible_turno, fecha_clave_asignada)
                mejor_hora_entrada = current_entry_time # Almacenamos la hora de entrada más temprana encontrada

    # Si se encontró un turno, se devuelve la data del turno asociado a la entrada más temprana
    return mejor_turno_data if mejor_turno_data else (None, None, None, None, None)

# --- 4. Calculo de horas ---

def calcular_turnos(df: pd.DataFrame, lugares_normalizados: list, tolerancia_llegada_tarde: int):
    """
    Agrupa por ID y FECHA_CLAVE_TURNO.
    Prioriza la ENTRADA que mejor se alinea a un turno programado,
    usando la lógica robusta que puede reasignar la FECHA_CLAVE_TURNO.
    
    Nota: Se usan los nombres de columnas en minúscula que fueron normalizados y renombrados
    en la función de carga: 'id_trabajador', 'nombre', 'porteria', etc.
    """
    
    df_filtrado = df[(df['PORTERIA_NORMALIZADA'].isin(lugares_normalizados)) & (df['TIPO_MARCACION'].isin(['ent', 'sal']))].copy()
    # Usando 'id_trabajador' (renombrada) y 'FECHA_HORA'
    df_filtrado.sort_values(by=['id_trabajador', 'FECHA_HORA'], inplace=True)

    if df_filtrado.empty: return pd.DataFrame()

    resultados = []

    # Agrupa por ID de trabajador y por la fecha clave de la jornada (maneja turnos nocturnos)
    for (id_trabajador, fecha_clave_turno), grupo in df_filtrado.groupby(['id_trabajador', 'FECHA_CLAVE_TURNO']):

        nombre = grupo['nombre'].iloc[0]
        entradas = grupo[grupo['TIPO_MARCACION'] == 'ent']
        
        # Inicialización de variables para el cálculo
        entrada_real = pd.NaT
        porteria_entrada = 'N/A'
        salida_real = pd.NaT
        porteria_salida = 'N/A'
        turno_nombre, info_turno, inicio_turno, fin_turno, fecha_clave_final = (None, None, None, None, fecha_clave_turno)
        horas_trabajadas = 0.0
        horas_extra = 0.0
        llegada_tarde_flag = False
        estado_calculo = "Sin Marcaciones Válidas (E/S)"
        salida_fue_real = False # Flag para saber si se usó una marcación real de salida
        

        mejor_entrada_para_turno = pd.NaT
        mejor_turno_data = (None, None, None, None, None)
        # El valor de menor_diferencia_turno ya no se usa como criterio principal, pero se mantiene la lógica de iteración.
        # Se usa una hora lo más tardía posible para priorizar entradas tempranas.
        mejor_hora_entrada_global = datetime.max 

        # --- REVISIÓN CLAVE 1: Encontrar la mejor entrada que se alinee a un turno ---
        if not entradas.empty:
            for index, row in entradas.iterrows():
                current_entry_time = row['FECHA_HORA']
                
                # Intentar asignar un turno a esta marcación de entrada, permitiendo reasignación de fecha clave
                turno_data = obtener_turno_para_registro(current_entry_time, fecha_clave_turno)
                turno_nombre_temp, info_turno_temp, inicio_turno_temp, fin_turno_temp, fecha_clave_final_temp = turno_data
                
                if turno_nombre_temp is not None:
                    
                    # Si esta entrada es más temprana que la mejor entrada encontrada globalmente para el grupo, la actualizamos
                    if current_entry_time < mejor_hora_entrada_global:
                        mejor_hora_entrada_global = current_entry_time
                        mejor_entrada_para_turno = current_entry_time
                        mejor_turno_data = turno_data


            # Si se encontró un turno asociado a la mejor entrada
            if pd.notna(mejor_entrada_para_turno):
                entrada_real = mejor_entrada_para_turno
                turno_nombre, info_turno, inicio_turno, fin_turno, fecha_clave_final = mejor_turno_data
                
                # Obtener porteria de la entrada real (usando el nombre de columna en minúscula 'porteria')
                porteria_entrada = entradas[entradas['FECHA_HORA'] == entrada_real]['porteria'].iloc[0]
                
                # --- REVISIÓN CLAVE 2: Filtro y/o Inferencia de Salida ---
                
                # Calcula el límite máximo de salida aceptable
                max_salida_aceptable = fin_turno + timedelta(hours=MAX_EXCESO_SALIDA_HRS)
                
                # Filtra las salidas que ocurrieron DESPUÉS de la ENTRADA REAL seleccionada y DENTRO del límite aceptable
                valid_salidas = df_filtrado[
                    (df_filtrado['id_trabajador'] == id_trabajador) &
                    (df_filtrado['TIPO_MARCACION'] == 'sal') &
                    (df_filtrado['FECHA_HORA'] > entrada_real) &
                    (df_filtrado['FECHA_HORA'] <= max_salida_aceptable)
                ]
                
                if valid_salidas.empty:
                    # SI NO HAY SALIDA VÁLIDA: ASUMIR SALIDA A LA HORA PROGRAMADA DEL FIN DE TURNO
                    salida_real = fin_turno
                    porteria_salida = 'ASUMIDA (Falta Salida/Salida Inválida)'
                    estado_calculo = "ASUMIDO (Falta Salida/Salida Inválida)"
                    salida_fue_real = False
                else:
                    # Usar la última salida REAL válida
                    salida_real = valid_salidas['FECHA_HORA'].max()
                    # Usando el nombre de columna en minúscula 'porteria'
                    porteria_salida = valid_salidas[valid_salidas['FECHA_HORA'] == salida_real]['porteria'].iloc[0]
                    estado_calculo = "Calculado"
                    salida_fue_real = True
                    
                # --- PARA MICRO-JORNADAS ---
                # Si se usó una SALIDA REAL, pero la duración es muy corta (< 1 hora), 
                # forzamos la ASSUMPCIÓN al fin de turno para evitar el problema de jornadas de 2 minutos.
                if salida_fue_real:
                    duracion_check = salida_real - entrada_real
                    if duracion_check < timedelta(hours=MIN_DURACION_ACEPTABLE_REAL_SALIDA_HRS):
                        salida_real = fin_turno
                        porteria_salida = 'ASUMIDA (Micro-jornada detectada)'
                        estado_calculo = "ASUMIDO (Micro-jornada detectada)"
                        salida_fue_real = False

                # --- 3. REGLAS DE CÁLCULO DE HORAS ---

                # La duración total es el tiempo entre la entrada real y la salida (real o asumida)
                duracion_total = salida_real - entrada_real
                
                # Regla de cálculo por defecto: inicia en el turno programado
                inicio_efectivo_calculo = inicio_turno
                llegada_tarde_flag = False
                
                # 1. Regla para LLEGADA TARDE (Más de 40 minutos tarde) - Tiene prioridad
                if entrada_real > inicio_turno + timedelta(minutes=tolerancia_llegada_tarde):
                    # Si llega tarde más la tolerancia (40 min), el cálculo inicia en la entrada real
                    inicio_efectivo_calculo = entrada_real
                    llegada_tarde_flag = True
                    
                # 2. Regla para ENTRADA TEMPRANA (Cualquier entrada antes del inicio programado)
                elif entrada_real < inicio_turno:
                    
                    # Calcular el tiempo de antelación
                    early_timedelta = inicio_turno - entrada_real
                    
                    # Regla: Si la antelación es mayor a 30 minutos, se paga desde la hora de entrada real.
                    if early_timedelta > timedelta(minutes=UMBRAL_PAGO_ENTRADA_TEMPRANA_MINUTOS):
                        # Caso 1: Muy temprano (> 30 minutos antes) -> Contar desde la hora real de entrada
                        inicio_efectivo_calculo = entrada_real
                    
                    else:
                        # Caso 2: Temprano (<= 30 minutos antes) -> Contar desde el inicio programado
                        inicio_efectivo_calculo = inicio_turno
                
                # Si no cae en 1 o 2 (ej: llega a tiempo o ligeramente tarde [<= 40 min]),
                # el cálculo se mantiene en el valor por defecto: inicio_efectivo_calculo = inicio_turno.
                
                duracion_efectiva_calculo = salida_real - inicio_efectivo_calculo

                if duracion_efectiva_calculo < timedelta(seconds=0):
                    horas_trabajadas = 0.0
                    horas_extra = 0.0
                    estado_calculo = "Error: Duración efectiva negativa"
                else:
                    horas_trabajadas = round(duracion_efectiva_calculo.total_seconds() / 3600, 2)
                    
                    horas_turno = info_turno["duracion_hrs"]
                    
                    # Para jornadas asumidas, aún se aplica el cálculo de horas extra si la duración supera el turno.
                    # Mantenemos esta regla para jornadas que no fueron ni Calculadas (por micro-jornada) ni Asumidas.
                    if duracion_total < timedelta(hours=4) and estado_calculo == "Jornada Corta (< 4h de Ent-Sal)":
                        # Solo se aplica si el estado no fue modificado por la detección de micro-jornada
                        horas_extra = 0.0
                    else:
                        horas_extra = max(0, round(horas_trabajadas - horas_turno, 2))
                        if estado_calculo == "Calculado" and not salida_fue_real:
                            # Si se forzó la asunción por micro-jornada, el estado ya está bien.
                            pass
                        elif estado_calculo == "Calculado":
                            # Mantener "Calculado" si se usaron entradas/salidas reales.
                            estado_calculo = "Calculado" 
                        elif estado_calculo == "ASUMIDO (Micro-jornada detectada)":
                            # Si se forzó la asunción por micro-jornada, usar este estado
                            pass

            else:
                estado_calculo = "Turno No Asignado (Entradas existen, pero ninguna se alinea con un turno programado)"

        elif pd.isna(entrada_real) and not grupo[grupo['TIPO_MARCACION'] == 'sal'].empty:
            # Caso de "Primer día" donde solo hay una salida de madrugada (FECHA_CLAVE_TURNO = Día anterior).
            # Esta es la jornada que el usuario quiere omitir, ya que no hay datos de entrada previos.
            # Se omite para limpiar el reporte.
            continue
            
        # --- Añade los resultados a la lista (Se reporta todo) ---
        ent_str = entrada_real.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(entrada_real) else 'N/A'
        sal_str = salida_real.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(salida_real) else 'N/A'
        # Usamos la fecha clave final reasignada para el reporte
        report_date = fecha_clave_final if fecha_clave_final else fecha_clave_turno
        inicio_str = inicio_turno.strftime("%H:%M:%S") if inicio_turno else 'N/A'
        fin_str = fin_turno.strftime("%H:%M:%S") if fin_turno else 'N/A'
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
            'Estado_Calculo': estado_calculo
        })

    return pd.DataFrame(resultados)

# --- 5. Interfaz Streamlit ---

st.set_page_config(page_title="Calculadora de Horas Extra", layout="wide")
st.title("📊 Calculadora de Horas Extra - NOEL")
st.write("Sube tu archivo de Excel para calcular las horas extra del personal. El sistema ahora **prioriza la Entrada más temprana** que se alinee a un turno programado.")

archivo_excel = st.file_uploader("Sube un archivo Excel (.xlsx)", type=["xlsx"])

if archivo_excel is not None:
    try:
        # Intenta leer la hoja específica 
        df_raw = pd.read_excel(archivo_excel, sheet_name='data')

        # 1. Definir la lista de nombres de columna que esperamos DESPUÉS de convertirlos a minúsculas
        columnas_requeridas_lower = [
            'cc', 'codtrabajador', 'nombre', 'fecha', 'hora', 'porteria', 'puntomarcacion'
        ]
        
        # 2. Crear un mapeo de nombres de columna actuales a sus versiones en minúscula.
        # ESTE PASO GARANTIZA LA ROBUSTEZ A MAYÚSCULAS/MINÚSCULAS
        col_map = {col: col.lower() for col in df_raw.columns}
        df_raw.rename(columns=col_map, inplace=True)

        # 3. Validar la existencia de todas las columnas requeridas normalizadas.
        if not all(col in df_raw.columns for col in columnas_requeridas_lower):
            st.error(f"⚠️ ERROR: Faltan columnas requeridas o tienen nombres incorrectos. Asegúrate de tener: **Cc, CodTrabajador, Nombre, Fecha, Hora, Porteria, PuntoMarcacion** (en cualquier formato de mayúsculas/minúsculas).")
            st.stop()

        # 4. Seleccionar las columnas normalizadas y renombrar 'codtrabajador' a 'id_trabajador'.
        df_raw = df_raw[columnas_requeridas_lower].copy()
        df_raw.rename(columns={'codtrabajador': 'id_trabajador'}, inplace=True)
        
        # --- NUEVO FILTRADO POR CÓDIGO DE TRABAJADOR ---
        # Asegura que la columna 'id_trabajador' sea de tipo entero para la comparación
        try:
            df_raw['id_trabajador'] = pd.to_numeric(df_raw['id_trabajador'], errors='coerce').astype('Int64')
        except:
             # Si falla la conversión directa a Int64 (por ejemplo, si hay muchos NaNs), se intenta con string
            st.warning("No se pudo convertir 'id_trabajador' a entero de forma segura. Se intentará con string.")
            df_raw['id_trabajador'] = df_raw['id_trabajador'].astype(str)
            codigos_filtro = [str(c) for c in CODIGOS_TRABAJADORES_FILTRO]
        else:
            codigos_filtro = CODIGOS_TRABAJADORES_FILTRO


        df_raw = df_raw[df_raw['id_trabajador'].isin(codigos_filtro)].copy()
        
        if df_raw.empty:
            st.error("⚠️ ERROR: Después del filtrado por código de trabajador, no quedan registros para procesar. Verifica que los códigos sean correctos.")
            st.stop()
        # --- FIN DEL FILTRADO ---
        
        # Preprocesamiento inicial de columnas (usando 'fecha')
        df_raw['fecha'] = pd.to_datetime(df_raw['fecha'], errors='coerce')  
        df_raw.dropna(subset=['fecha'], inplace=True)
            
        # --- Función para estandarizar el formato de la hora (manejo de floats y strings) ---
        def standardize_time_format(time_val):
            # Caso: la hora es un float (formato de Excel)
            if isinstance(time_val, float) and time_val <= 1.0: 
                total_seconds = int(time_val * 86400)
                hours, remainder = divmod(total_seconds, 3600)
                minutes, seconds = divmod(remainder, 60)
                return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                
            # Caso: la hora es un string (o fue convertida a string)
            time_str = str(time_val)
            parts = time_str.split(':')
            if len(parts) == 2:
                return f"{time_str}:00"
            elif len(parts) == 3:
                return time_str
            else:
                return '00:00:00' 

        # Aplica la estandarización y luego combina FECHA y HORA
        df_raw['hora'] = df_raw['hora'].apply(standardize_time_format)
            
        try:
            # Usando 'fecha' y 'hora' normalizadas
            df_raw['FECHA_HORA'] = pd.to_datetime(df_raw['fecha'].dt.strftime('%Y-%m-%d') + ' ' + df_raw['hora'], errors='coerce')
            df_raw.dropna(subset=['FECHA_HORA'], inplace=True)
        except Exception as e:
            st.error(f"Error al combinar FECHA y HORA. Revisa el formato de la columna HORA: {e}")
            st.stop() 

        # Normalización de las otras columnas de marcación (usando 'porteria' y 'puntomarcacion')
        df_raw['PORTERIA_NORMALIZADA'] = df_raw['porteria'].astype(str).str.strip().str.lower()
        # Mapeo de PuntoMarcacion a 'ent' o 'sal' (usando 'puntomarcacion')
        df_raw['TIPO_MARCACION'] = df_raw['puntomarcacion'].astype(str).str.strip().str.lower().replace({'entrada': 'ent', 'salida': 'sal'})

        # --- Función para asignar Fecha Clave de Turno (Lógica Nocturna) ---
        def asignar_fecha_clave_turno_corregida(row):
            fecha_original = row['FECHA_HORA'].date()
            hora_marcacion = row['FECHA_HORA'].time()
            tipo_marcacion = row['TIPO_MARCACION']
            
            # Regla de oro: Las ENTRADAS anclan la jornada al día en que ocurrieron.
            if tipo_marcacion == 'ent':
                return fecha_original
            
            # Regla nocturna: Las SALIDAS antes del corte se asocian al turno del día anterior.
            if tipo_marcacion == 'sal' and hora_marcacion < HORA_CORTE_NOCTURNO:
                return fecha_original - timedelta(days=1)
            
            # Otras salidas (después de 8 AM) pertenecen al día en que fueron marcadas.
            return fecha_original

        df_raw['FECHA_CLAVE_TURNO'] = df_raw.apply(asignar_fecha_clave_turno_corregida, axis=1)

        st.success(f"✅ Archivo cargado y preprocesado con éxito. Se encontraron {len(df_raw['FECHA_CLAVE_TURNO'].unique())} días de jornada para procesar de {len(df_raw['id_trabajador'].unique())} trabajadores filtrados.")

        # --- Ejecutar el Cálculo ---
        df_resultado = calcular_turnos(df_raw.copy(), LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS, TOLERANCIA_LLEGADA_TARDE_MINUTOS)

        if not df_resultado.empty:
            # Post-procesamiento para el reporte
            df_resultado['Estado_Llegada'] = df_resultado['Llegada_Tarde_Mas_40_Min'].map({True: 'Tarde', False: 'A tiempo'})
            df_resultado.sort_values(by=['NOMBRE', 'FECHA', 'ENTRADA_REAL'], inplace=True) 
            
            # Columnas a mostrar en la tabla final
            columnas_reporte = [
                'NOMBRE', 'ID_TRABAJADOR', 'FECHA', 'Dia_Semana', 'TURNO',
                'Inicio_Turno_Programado', 'Fin_Turno_Programado', 'Duracion_Turno_Programado_Hrs',
                'ENTRADA_REAL', 'PORTERIA_ENTRADA', 'SALIDA_REAL', 'PORTERIA_SALIDA',
                'Horas_Trabajadas_Netas', 'Horas_Extra', 'Horas', 'Minutos', 
                'Estado_Llegada', 'Estado_Calculo'
            ]

            st.subheader("Resultados de las Horas Extra")
            st.dataframe(df_resultado[columnas_reporte], use_container_width=True)

            # --- Lógica de descarga en Excel con formato condicional ---
            buffer_excel = io.BytesIO()
            with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
                df_to_excel = df_resultado[columnas_reporte].copy()
                df_to_excel.to_excel(writer, sheet_name='Reporte Horas Extra', index=False)

                workbook = writer.book
                worksheet = writer.sheets['Reporte Horas Extra']

                # Formatos
                orange_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}) # Tarde (> 40 min)
                gray_format = workbook.add_format({'bg_color': '#D9D9D9'}) # No calculado/Faltante
                yellow_format = workbook.add_format({'bg_color': '#FFF2CC', 'font_color': '#3C3C3C'}) # Asumido
                # Formato para Horas Extra > 30 minutos (Rojo Fuerte)
                red_extra_format = workbook.add_format({'bg_color': '#F8E8E8', 'font_color': '#D83A56', 'bold': True})
                
                # Aplica formatos condicionales basados en el dataframe original
                for row_num, row in df_resultado.iterrows():
                    excel_row = row_num + 1
                    
                    is_calculated = row['Estado_Calculo'] in ["Calculado", "ASUMIDO (Falta Salida/Salida Inválida)"]
                    is_late = row['Llegada_Tarde_Mas_40_Min']
                    is_assumed = row['Estado_Calculo'].startswith("ASUMIDO")
                    is_missing_entry = row['Estado_Calculo'].startswith("Sin Marcaciones Válidas") or row['Estado_Calculo'].startswith("Turno No Asignado")
                    
                    # Verifica si las horas extra son mayores al umbral de 30 minutos (0.5 horas)
                    is_excessive_extra = row['Horas_Extra'] > UMBRAL_HORAS_EXTRA_RESALTAR

                    # PASO 1: Determinar el formato base de la fila (Baja prioridad)
                    base_format = None
                    if is_missing_entry or (not is_calculated and not is_assumed):
                        base_format = gray_format
                    elif is_assumed:
                        # Formato ASUMIDO (Amarillo claro)
                        base_format = yellow_format

                    for col_idx, col_name in enumerate(df_to_excel.columns):
                        value = row[col_name]
                        cell_format = base_format # Iniciar con el formato base de la fila
                        
                        # PASO 2: Aplicar Overrides de Alta Prioridad
                        
                        # Override A: Llegada Tarde (Naranja/Rojo)
                        if col_name == 'ENTRADA_REAL' and is_late:
                            cell_format = orange_format
                        
                        # Override B: Horas Extra > 30 minutos (Rojo Fuerte) - SATISFACE SOLICITUD DEL USUARIO
                        if is_excessive_extra and col_name in ['Horas_Extra', 'Horas', 'Minutos']:
                            cell_format = red_extra_format

                        # Escribir el valor en la celda
                        worksheet.write(excel_row, col_idx, value if pd.notna(value) else 'N/A', cell_format)

                # Ajustar el ancho de las columnas
                for i, col in enumerate(df_to_excel.columns):
                    max_len = max(df_to_excel[col].astype(str).str.len().max(), len(col)) + 2
                    worksheet.set_column(i, i, max_len)

            buffer_excel.seek(0)

            st.download_button(
                label="Descargar Reporte de Horas Extra (Excel)",
                data=buffer_excel,
                file_name="Reporte_Marcacion_Horas_Extra.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("No se encontraron jornadas válidas después de aplicar los filtros.")

    except KeyError as e:
        if 'BaseDatos Modificada' in str(e):
            st.error(f"⚠️ ERROR: El archivo Excel debe contener una hoja llamada **'BaseDatos Modificada'** y las columnas requeridas.")
        else:
            st.error(f"⚠️ ERROR: Faltan columnas requeridas o tienen nombres incorrectos: {e}")
    except Exception as e:
        st.error(f"Error crítico al procesar el archivo: {e}. Por favor, verifica el formato de los datos.")

st.markdown("---")
st.caption("Somos NOEL DE CORAZÓN ❤️ - Herramienta de Cálculo de Turnos y Horas Extra")

