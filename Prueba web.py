# -*- coding: utf-8 -*-
"""
Calculadora de Horas Extra con Prioridad de Ubicación (Puesto > Portería).
"""

import pandas as pd
from datetime import datetime, timedelta
import streamlit as st
import io
import numpy as np

# --- CÓDIGOS DE TRABAJADORES PERMITIDOS ---
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

# Lista de puestos de trabajo (máquinas, oficinas, etc.) - PRIORIDAD ALTA
LUGARES_PUESTOS = [
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

# Lista de porterías (tornos peatonales y vehiculares) - PRIORIDAD BAJA (Fallback)
LUGARES_PORTERIAS = [
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

LUGARES_PUESTOS_NORMALIZADOS = [lugar.strip().lower() for lugar in LUGARES_PUESTOS]
LUGARES_PORTERIAS_NORMALIZADOS = [lugar.strip().lower() for lugar in LUGARES_PORTERIAS]

# Máximo de horas después del fin de turno programado que se acepta una salida como válida.
MAX_EXCESO_SALIDA_HRS = 3
# Hora de corte para definir si una SALIDA en la mañana pertenece al turno del día anterior (ej: 08:00 AM)
HORA_CORTE_NOCTURNO = datetime.strptime("08:00:00", "%H:%M:%S").time()

# --- CONSTANTES DE TOLERANCIA REVISADAS ---
# Tolerancia para considerar la llegada como 'tarde' para el cálculo de horas.
TOLERANCIA_LLEGADA_TARDE_MINUTOS = 40

# Tolerancia MÁXIMA para considerar la llegada como 'temprana' para la asignación de turno (6 horas)
TOLERANCIA_ENTRADA_TEMPRANA_MINUTOS = 360

# Máxima tardanza permitida para que una entrada CUENTE para la ASIGNACIÓN de un turno (3 horas)
TOLERANCIA_ASIGNACION_TARDE_MINUTOS = 180

# --- HORAS EXTRA LLEGADA TEMPRANO ---
# Umbral de tiempo (en minutos) para determinar si la llegada temprana se paga desde la hora real.
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
    Busca el turno programado más cercano (menor diferencia absoluta) a la marcación de entrada,
    dentro de las ventanas de tolerancia.
    Retorna: (nombre, info, inicio_turno, fin_turno, fecha_clave_final)
    """
    mejor_turno_data = None
    menor_diferencia = timedelta(days=999)

    # Candidatos a turno para el día de la FECHA CLAVE (Día X)
    turnos_candidatos = buscar_turnos_posibles(fecha_clave_turno_reporte)

    # Si la marcación es temprano en la mañana, añadir candidatos nocturnos del día anterior (Día X - 1)
    hora_evento = fecha_hora_evento.time()
    if hora_evento < HORA_CORTE_NOCTURNO:
        fecha_clave_anterior = fecha_clave_turno_reporte - timedelta(days=1)
        turnos_candidatos.extend(buscar_turnos_posibles(fecha_clave_anterior))

    for nombre_turno, info_turno, inicio_posible_turno, fin_posible_turno, fecha_clave_asignada in turnos_candidatos:

        # --- LÓGICA DE RESTRICCIÓN DE VENTANA DE ENTRADA ---
        # 1. Límite más temprano (6 horas antes)
        rango_inicio_temprano = inicio_posible_turno - timedelta(minutes=TOLERANCIA_ENTRADA_TEMPRANA_MINUTOS)

        # 2. Límite más tardío (3 horas después)
        rango_fin_tarde = inicio_posible_turno + timedelta(minutes=TOLERANCIA_ASIGNACION_TARDE_MINUTOS)

        # Validar si el evento (la entrada) cae en esta ventana estricta alrededor del INICIO PROGRAMADO.
        if fecha_hora_evento >= rango_inicio_temprano and fecha_hora_evento <= rango_fin_tarde:

            # La diferencia se calcula entre la entrada real y el inicio PROGRAMADO del turno
            diferencia = abs(fecha_hora_evento - inicio_posible_turno)

            # Buscamos la menor diferencia absoluta para encontrar el turno más probable
            if mejor_turno_data is None or diferencia < menor_diferencia:
                mejor_turno_data = (nombre_turno, info_turno, inicio_posible_turno, fin_posible_turno, fecha_clave_asignada)
                menor_diferencia = diferencia

    return mejor_turno_data if mejor_turno_data else (None, None, None, None, None)

# --- 4. Calculo de horas ---

def _encontrar_mejor_jornada(grupo_completo: pd.DataFrame, lugares_normalizados: list, fecha_clave_turno: datetime.date):
    """
    Busca la mejor marcación de entrada dentro de los lugares_normalizados
    que se alinee con un turno programado (el que esté más cerca).
    Retorna: (mejor_entrada_dt, turno_data)
    """
    # 1. Filtrar solo entradas que estén en la lista de lugares provista (Puestos o Porterías)
    entradas_candidatas = grupo_completo[
        (grupo_completo['TIPO_MARCACION'] == 'ent') &
        (grupo_completo['PORTERIA_NORMALIZADA'].isin(lugares_normalizados))
    ].copy()

    mejor_entrada = pd.NaT
    mejor_turno_data = (None, None, None, None, None)
    menor_diferencia = timedelta(days=999)

    if not entradas_candidatas.empty:
        for index, row in entradas_candidatas.iterrows():
            current_entry_time = row['FECHA_HORA']

            # 2. Intentar asignar un turno a esta marcación (busca el más cercano dentro de la ventana de +/- 6h)
            turno_data = obtener_turno_para_registro(current_entry_time, fecha_clave_turno)
            turno_nombre_temp, info_turno_temp, inicio_turno_temp, _, _ = turno_data

            if turno_nombre_temp is not None:
                # 3. Calcula la diferencia absoluta con el inicio programado del turno
                diferencia = abs(current_entry_time - inicio_turno_temp)

                # 4. Si es la mejor diferencia hasta ahora, guardarla
                if pd.isna(mejor_entrada) or diferencia < menor_diferencia:
                    menor_diferencia = diferencia
                    mejor_entrada = current_entry_time
                    mejor_turno_data = turno_data

    return mejor_entrada, mejor_turno_data


def calcular_turnos(df: pd.DataFrame, puestos_normalizados: list, porterias_normalizadas: list, tolerancia_llegada_tarde: int):
    """
    Agrupa por ID y luego por FECHA_CLAVE_TURNO.
    Implementa la lógica de prioridad de lugares (Puestos > Porterías) y el filtro de límites.
    """
    
    # 1. Filtro inicial solo por tipo de marcación (ent/sal) y ordenar
    df_base = df[df['TIPO_MARCACION'].isin(['ent', 'sal'])].copy()
    df_base.sort_values(by=['id_trabajador', 'FECHA_HORA'], inplace=True)

    if df_base.empty: return pd.DataFrame()

    resultados = []

    # Agrupa por ID de trabajador
    for id_trabajador, grupo_trabajador in df_base.groupby('id_trabajador'):

        nombre = grupo_trabajador['nombre'].iloc[0]

        # 1. IDENTIFICAR RANGO ACTIVO DE JORNADAS (FECHA_CLAVE_TURNO con al menos una ENTRADA)
        fechas_con_entrada = grupo_trabajador[grupo_trabajador['TIPO_MARCACION'] == 'ent']['FECHA_CLAVE_TURNO'].unique()

        if fechas_con_entrada.size == 0:
            continue # No hay entradas para este trabajador

        # Estas son las fechas clave que delimitan el bloque de trabajo real del empleado en la data cargada
        min_fecha_activa = fechas_con_entrada.min()
        max_fecha_activa = fechas_con_entrada.max()

        # 2. Iterar sobre las jornadas agrupadas por FECHA_CLAVE_TURNO
        for fecha_clave_turno, grupo_completo in grupo_trabajador.groupby('FECHA_CLAVE_TURNO'):

            # Reiniciar variables para la jornada actual
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
            porterias_validas = [] # Lista de lugares que finalmente se usó para la ENTRADA/SALIDA

            # --- 3. LÓGICA DE PRIORIDAD: PASADA 1 (PUESTOS DE TRABAJO) ---
            entrada_real, mejor_turno_data = _encontrar_mejor_jornada(
                grupo_completo, puestos_normalizados, fecha_clave_turno
            )

            if pd.notna(entrada_real):
                # Se encontró una jornada válida con Puestos
                porterias_validas = puestos_normalizados
            else:
                # --- 4. FALLBACK: PASADA 2 (PORTERIAS) ---
                entrada_real, mejor_turno_data = _encontrar_mejor_jornada(
                    grupo_completo, porterias_normalizadas, fecha_clave_turno
                )

                if pd.notna(entrada_real):
                    # Se encontró una jornada válida con Porterías
                    porterias_validas = porterias_normalizadas
                else:
                    # Caso: No se encontró entrada válida en Puestos ni en Porterías
                    estado_calculo = "Turno No Asignado (Entrada no alinea con turno)"
                    pass

            # --- CONTINUAR CÁLCULO SI SE ENCONTRÓ UNA ENTRADA VÁLIDA ---
            if pd.notna(entrada_real):
                turno_nombre, info_turno, inicio_turno, fin_turno, fecha_clave_final = mejor_turno_data

                # Obtener porteria de la entrada real
                porteria_entrada = grupo_completo[grupo_completo['FECHA_HORA'] == entrada_real]['porteria'].iloc[0]

                # --- REVISIÓN CLAVE 5: Filtro y/o Inferencia de Salida ---

                max_salida_aceptable = fin_turno + timedelta(hours=MAX_EXCESO_SALIDA_HRS)

                # Filtrar salidas que están DENTRO de los LUGARES VALIDADOS (Puestos o Porterías)
                # OJO: La salida debe pertenecer al mismo tipo de lugar que la entrada (implícito en valid_salidas)
                valid_salidas = grupo_completo[
                    (grupo_completo['TIPO_MARCACION'] == 'sal') &
                    (grupo_completo['FECHA_HORA'] > entrada_real) &
                    (grupo_completo['FECHA_HORA'] <= max_salida_aceptable) &
                    (grupo_completo['PORTERIA_NORMALIZADA'].isin(porterias_validas)) # FILTRO CLAVE
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
                # forzamos la ASSUMPCIÓN al fin de turno.
                if salida_fue_real:
                    duracion_check = salida_real - entrada_real
                    if duracion_check < timedelta(hours=MIN_DURACION_ACEPTABLE_REAL_SALIDA_HRS):
                        salida_real = fin_turno
                        porteria_salida = 'ASUMIDA (Micro-jornada detectada)'
                        estado_calculo = "ASUMIDO (Micro-jornada detectada)"
                        salida_fue_real = False

                # --- 6. REGLAS DE CÁLCULO DE HORAS ---

                inicio_efectivo_calculo = inicio_turno
                llegada_tarde_flag = False

                # 1. Regla para LLEGADA TARDE (Más de 40 minutos tarde) - Tiene prioridad
                if entrada_real > inicio_turno + timedelta(minutes=tolerancia_llegada_tarde):
                    inicio_efectivo_calculo = entrada_real
                    llegada_tarde_flag = True

                # 2. Regla para ENTRADA TEMPRANA (Cualquier entrada antes del inicio programado)
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

                    # Ajuste de estado después del cálculo
                    if estado_calculo.startswith("ASUMIDO"):
                         pass # Mantiene el estado de asunción
                    else:
                        estado_calculo = "Calculado"

            # --- FILTROS POST-CÁLCULO PARA INCONSISTENCIAS Y LÍMITES (LÓGICA DE VACÍOS) ---

            is_boundary_date = (fecha_clave_turno == min_fecha_activa) or (fecha_clave_turno == max_fecha_activa)

            # FILTRO 1: Descartar jornadas ASUMIDAS en los días de límite
            if is_boundary_date and estado_calculo == "ASUMIDO (Falta Salida/Salida Inválida)":
                continue # Omitir este cálculo

            # FILTRO 2: Descartar jornadas SAL-only
            if pd.isna(entrada_real) and not grupo_completo[grupo_completo['TIPO_MARCACION'] == 'sal'].empty and estado_calculo == "Sin Marcaciones Válidas (E/S)":
                 continue # Omitir este cálculo

            # --- Añade los resultados a la lista (Se reporta todo lo que no se descartó) ---
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
st.write("Sube tu archivo de Excel para calcular las horas extra. El sistema ahora **filtra por los IDs de trabajador permitidos**, prioriza las marcaciones en **Puestos de Trabajo** y descarta jornadas incompletas en los límites del reporte.")

archivo_excel = st.file_uploader("Sube un archivo Excel (.xlsx)", type=["xlsx"])

if archivo_excel is not None:
    try:
        # Lee la primera hoja si no se especifica
        # OJO: Se cambió para leer la hoja 'data' explícitamente, si existe. Si no existe, leerá la primera.
        try:
             df_raw = pd.read_excel(archivo_excel, sheet_name='data')
        except ValueError:
             df_raw = pd.read_excel(archivo_excel)


        # 1. Definir la lista de nombres de columna que esperamos DESPUÉS de convertirlos a minúsculas
        columnas_requeridas_lower = [
            'cc', 'codtrabajador', 'nombre', 'fecha', 'hora', 'porteria', 'puntomarcacion'
        ]

        # 2. Crear un mapeo de nombres de columna actuales a sus versiones en minúscula.
        col_map = {col: col.lower() for col in df_raw.columns}
        df_raw.rename(columns=col_map, inplace=True)

        # 3. Validar la existencia de todas las columnas requeridas normalizadas.
        if not all(col in df_raw.columns for col in columnas_requeridas_lower):
            st.error(f"⚠️ ERROR: Faltan columnas requeridas o tienen nombres incorrectos. Asegúrate de tener: **Cc, CodTrabajador, Nombre, Fecha, Hora, Porteria, PuntoMarcacion** (en cualquier formato de mayúsculas/minúsculas).")
            st.stop()

        # 4. Seleccionar las columnas normalizadas y renombrar 'codtrabajador' a 'id_trabajador'.
        df_raw = df_raw[columnas_requeridas_lower].copy()
        df_raw.rename(columns={'codtrabajador': 'id_trabajador'}, inplace=True)

        # --- FILTRADO POR CÓDIGO DE TRABAJADOR (Máxima Robustez STR) ---
        
        # 1. Preparar el filtro como strings para robustez (cubre IDs numéricos grandes y pequeños)
        codigos_filtro_str = [str(c) for c in CODIGOS_TRABAJADORES_FILTRO]
        
        # 2. Asegurar que la columna del DataFrame también sea string, eliminando espacios
        df_raw['id_trabajador'] = df_raw['id_trabajador'].astype(str).str.strip()

        # 3. Aplicar el filtro
        df_raw = df_raw[df_raw['id_trabajador'].isin(codigos_filtro_str)].copy()
        
        if df_raw.empty:
            st.error("⚠️ ERROR: Después del filtrado por código de trabajador, no quedan registros para procesar. Verifica que los códigos en el archivo coincidan con la lista permitida.")
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
            try:
                time_str = str(time_val)
                parts = time_str.split(':')
                if len(parts) == 2:
                    return f"{time_str}:00"
                elif len(parts) == 3:
                    return time_str
                # Si es un datetime.time object
                elif isinstance(time_val, datetime.time):
                    return time_val.strftime("%H:%M:%S")
                else:
                    return '00:00:00'
            except Exception:
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
        df_resultado = calcular_turnos(
            df_raw.copy(),
            LUGARES_PUESTOS_NORMALIZADOS, # Prioridad 1
            LUGARES_PORTERIAS_NORMALIZADOS, # Prioridad 2 (Fallback)
            TOLERANCIA_LLEGADA_TARDE_MINUTOS
        )

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
                yellow_format = workbook.add_format({'bg_color': '#FFF2CC'}) # Asumido
                # Formato para Horas Extra > 30 minutos (Rojo Fuerte)
                red_extra_format = workbook.add_format({'bg_color': '#F8E8E8', 'font_color': '#D83A56', 'bold': True})

                # Aplica formatos condicionales basados en el dataframe original
                for row_num, row in df_resultado.iterrows():
                    excel_row = row_num + 1

                    is_late = row['Llegada_Tarde_Mas_40_Min']
                    is_assumed = row['Estado_Calculo'].startswith("ASUMIDO")
                    is_unassigned = row['Estado_Calculo'].startswith("Turno No Asignado") or row['Estado_Calculo'].startswith("Sin Marcaciones Válidas")

                    # Verifica si las horas extra son mayores al umbral de 30 minutos (0.5 horas)
                    is_excessive_extra = row['Horas_Extra'] > UMBRAL_HORAS_EXTRA_RESALTAR

                    # PASO 1: Determinar el formato base de la fila (Baja prioridad)
                    base_format = None
                    if is_unassigned:
                        base_format = gray_format
                    elif is_assumed:
                        base_format = yellow_format

                    for col_idx, col_name in enumerate(df_to_excel.columns):
                        value = row[col_name]
                        cell_format = base_format # Iniciar con el formato base de la fila

                        # PASO 2: Aplicar Overrides de Alta Prioridad

                        # Override A: Llegada Tarde (Naranja/Rojo)
                        if col_name == 'ENTRADA_REAL' and is_late:
                            cell_format = orange_format

                        # Override B: Horas Extra > 30 minutos (Rojo Fuerte)
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

    except Exception as e:
        st.error(f"Error crítico al procesar el archivo: {e}. Por favor, verifica el formato de los datos.")

st.markdown("---")
st.caption("Somos NOEL DE CORAZÓN ❤️ - Herramienta de Cálculo de Turnos y Horas Extra")















