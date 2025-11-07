# -*- coding: utf-8 -*-
"""
Calculadora de Horas Extra - Modificada con Prioridad de Marcación.
Prioridad: Puestos de Trabajo (PT) > Porterías (P)
"""

import pandas as pd
from datetime import datetime, timedelta
import streamlit as st
import io
import numpy as np

# --- CÓDIGOS DE TRABAJADORES PERMITIDOS (ACTUALIZADO) ---
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

# Lista de lugares considerados como PUESTOS DE TRABAJO (PT) - PRIORIDAD 1 (Ordenada Alfabéticamente)
LUGARES_PUESTOS_TRABAJO = [
    "NOEL_MDE_CONTROL_BUHLER_ENT", "NOEL_MDE_CONTROL_BUHLER_SAL", "NOEL_MDE_ESENCIAS_1_ENT",
    "NOEL_MDE_ESENCIAS_1_SAL", "NOEL_MDE_ESENCIAS_2_SAL", "NOEL_MDE_ING_MENORES_1_ENT",
    "NOEL_MDE_ING_MENORES_1_SAL", "NOEL_MDE_ING_MENORES_2_ENT", "NOEL_MDE_ING_MENORES_2_SAL",
    "NOEL_MDE_ING_MEN_ALERGENOS_ENT", "NOEL_MDE_ING_MEN_ALERGENOS_SAL", "NOEL_MDE_ING_MEN_CREMAS_ENT",
    "NOEL_MDE_ING_MEN_CREMAS_SAL", "NOEL_MDE_MOLINETE_BODEGA_EXT_SAL", "NOEL_MDE_MR_ASPIRACION_ENT",
    "NOEL_MDE_MR_HORNO_11_ENT", "NOEL_MDE_MR_HORNO_11_SAL", "NOEL_MDE_MR_HORNO_18_ENT",
    "NOEL_MDE_MR_HORNO_18_SAL", "NOEL_MDE_MR_HORNO_1-3_ENT", "NOEL_MDE_MR_HORNO_1-3_SAL",
    "NOEL_MDE_MR_HORNO_2-12_ENT", "NOEL_MDE_MR_HORNO_2-12_SAL", "NOEL_MDE_MR_HORNO_2-4-5_SAL",
    "NOEL_MDE_MR_HORNO_4-5_ENT", "NOEL_MDE_MR_HORNO_4-5_SAL", "NOEL_MDE_MR_HORNO_6-8-9_ENT",
    "NOEL_MDE_MR_HORNO_6-8-9_SAL", "NOEL_MDE_MR_HORNO_6-8-9_SAL_2", "NOEL_MDE_MR_HORNO_7-10_ENT",
    "NOEL_MDE_MR_HORNO_7-10_SAL", "NOEL_MDE_MR_HORNOS_ENT", "NOEL_MDE_MR_HORNOS_SAL",
    "NOEL_MDE_MR_MEZCLAS_ENT", "NOEL_MDE_MR_MEZCLAS_SAL", "NOEL_MDE_MR_SERVICIOS_2_ENT",
    "NOEL_MDE_MR_SERVICIOS_2_SAL", "NOEL_MDE_MR_TUNEL_VIENTO_1_ENT", "NOEL_MDE_MR_TUNEL_VIENTO_2_ENT",
    "NOEL_MDE_MR_WAFER_RCH_CREMAS_ENT", "NOEL_MDE_MR_WAFER_RCH_CREMAS_SAL", "NOEL_MDE_OFIC_PRODUCCION_ENT",
    "NOEL_MDE_OFIC_PRODUCCION_SAL", "NOEL_MDE_PRINCIPAL_ENT", "NOEL_MDE_PRINCIPAL_SAL",
    "NOEL_MDE_RECURSOS_HUMANOS_ENT", "NOEL_MDE_RECURSOS_HUMANOS_SAL", "NOEL_MDE_TORNIQUETE_PATIO_ENT",
    "NOEL_MDE_TORNIQUETE_PATIO_SAL", "NOEL_MDE_TORNIQUETE_SORTER_ENT", "NOEL_MDE_TORNIQUETE_SORTER_SAL"
]

# Lista de Porterías (P) - PRIORIDAD 2 (Ordenada Alfabéticamente)
LUGARES_PORTERIAS = [
    "NOEL_MDE_PORT_1_PEATONAL_1_ENT", "NOEL_MDE_PORT_2_PEATONAL_1_ENT", "NOEL_MDE_PORT_2_PEATONAL_1_SAL",
    "NOEL_MDE_PORT_2_PEATONAL_2_ENT", "NOEL_MDE_PORT_2_PEATONAL_2_SAL", "NOEL_MDE_PORT_2_PEATONAL_3_ENT",
    "NOEL_MDE_PORT_2_PEATONAL_3_SAL", "NOEL_MDE_TORN_PORTERIA_3_ENT", "NOEL_MDE_TORN_PORTERIA_3_SAL",
    "NOEL_MDE_VEHICULAR_PORT_1_ENT", "NOEL_MDE_VEHICULAR_PORT_1_SAL"
]

LUGARES_PUESTOS_TRABAJO_NORMALIZADOS = [lugar.strip().lower() for lugar in LUGARES_PUESTOS_TRABAJO]
LUGARES_PORTERIAS_NORMALIZADAS = [lugar.strip().lower() for lugar in LUGARES_PORTERIAS]

# Máximo de horas después del fin de turno programado que se acepta una salida como válida.
MAX_EXCESO_SALIDA_HRS = 3
# Hora de corte para definir si una SALIDA en la mañana pertenece al turno del día anterior (ej: 08:00 AM)
HORA_CORTE_NOCTURNO = datetime.strptime("08:00:00", "%H:%M:%S").time()

# --- CONSTANTES DE TOLERANCIA REVISADAS ---
TOLERANCIA_LLEGADA_TARDE_MINUTOS = 40
TOLERANCIA_ENTRADA_TEMPRANA_MINUTOS = 360
TOLERANCIA_ASIGNACION_TARDE_MINUTOS = 180
UMBRAL_PAGO_ENTRADA_TEMPRANA_MINUTOS = 30
MIN_DURACION_ACEPTABLE_REAL_SALIDA_HRS = 1
UMBRAL_HORAS_EXTRA_RESALTAR = 30 / 60

# --- 3. Obtener turno basado en fecha y hora (Lógica sin cambios) ---

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
    """Busca el turno programado más cercano a la marcación de entrada."""
    mejor_turno_data = None
    mejor_hora_entrada = datetime.max

    turnos_candidatos = buscar_turnos_posibles(fecha_clave_turno_reporte)

    hora_evento = fecha_hora_evento.time()
    if hora_evento < HORA_CORTE_NOCTURNO:
        fecha_clave_anterior = fecha_clave_turno_reporte - timedelta(days=1)
        turnos_candidatos.extend(buscar_turnos_posibles(fecha_clave_anterior))

    for nombre_turno, info_turno, inicio_posible_turno, fin_posible_turno, fecha_clave_asignada in turnos_candidatos:

        # Límite más temprano (6 horas antes)
        rango_inicio_temprano = inicio_posible_turno - timedelta(minutes=TOLERANCIA_ENTRADA_TEMPRANA_MINUTOS)
        
        # Límite más tardío (3 horas después)
        rango_fin_tarde = inicio_posible_turno + timedelta(minutes=TOLERANCIA_ASIGNACION_TARDE_MINUTOS + 5)
        
        if fecha_hora_evento >= rango_inicio_temprano and fecha_hora_evento <= rango_fin_tarde:
            current_entry_time = fecha_hora_evento

            if mejor_turno_data is None or current_entry_time < mejor_hora_entrada:
                mejor_turno_data = (nombre_turno, info_turno, inicio_posible_turno, fin_posible_turno, fecha_clave_asignada)
                mejor_hora_entrada = current_entry_time

    return mejor_turno_data if mejor_turno_data else (None, None, None, None, None)

# --- 4. Lógica de Cálculo de Horas por Grupo (Extraída para la Prioridad) ---

def _ejecutar_calculo_para_grupo(grupo_marcaciones: pd.DataFrame, fecha_clave_turno: datetime.date, tolerancia_llegada_tarde: int):
    """
    Función auxiliar que ejecuta la lógica de asignación de turno y cálculo de horas 
    para un grupo de marcaciones filtrado (que ya son solo PT o solo P).
    Retorna un DataFrame de 1 fila con el resultado si hay turno asignado, o vacío.
    """
    
    if grupo_marcaciones.empty:
        return pd.DataFrame()
        
    id_trabajador = grupo_marcaciones['id_trabajador'].iloc[0]
    nombre = grupo_marcaciones['nombre'].iloc[0]
    entradas = grupo_marcaciones[grupo_marcaciones['TIPO_MARCACION'] == 'ent']
    
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
    salida_fue_real = False
    
    mejor_entrada_para_turno = pd.NaT
    mejor_turno_data = (None, None, None, None, None)
    mejor_hora_entrada_global = datetime.max

    # --- REVISIÓN CLAVE 1: Encontrar la mejor entrada (la más temprana) que se alinee a un turno ---
    if not entradas.empty:
        for index, row in entradas.iterrows():
            current_entry_time = row['FECHA_HORA']
            turno_data = obtener_turno_para_registro(current_entry_time, fecha_clave_turno)
            turno_nombre_temp, _, _, _, _ = turno_data
            
            if turno_nombre_temp is not None:
                if current_entry_time < mejor_hora_entrada_global:
                    mejor_hora_entrada_global = current_entry_time
                    mejor_entrada_para_turno = current_entry_time
                    mejor_turno_data = turno_data

        # Si se encontró un turno asociado a la mejor entrada
        if pd.notna(mejor_entrada_para_turno):
            entrada_real = mejor_entrada_para_turno
            turno_nombre, info_turno, inicio_turno, fin_turno, fecha_clave_final = mejor_turno_data
            
            porteria_entrada = entradas[entradas['FECHA_HORA'] == entrada_real]['porteria'].iloc[0]
            
            # --- REVISIÓN CLAVE 2: Filtro y/o Inferencia de Salida ---
            
            max_salida_aceptable = fin_turno + timedelta(hours=MAX_EXCESO_SALIDA_HRS)
            
            valid_salidas = grupo_marcaciones[
                (grupo_marcaciones['TIPO_MARCACION'] == 'sal') &
                (grupo_marcaciones['FECHA_HORA'] > entrada_real) &
                (grupo_marcaciones['FECHA_HORA'] <= max_salida_aceptable)
            ]
            
            if valid_salidas.empty:
                # ASUMIR SALIDA A LA HORA PROGRAMADA DEL FIN DE TURNO
                salida_real = fin_turno
                porteria_salida = 'ASUMIDA (Falta Salida/Salida Inválida)'
                estado_calculo = "ASUMIDO (Falta Salida/Salida Inválida)"
                salida_fue_real = False
            else:
                # Usar la última salida REAL válida
                salida_real = valid_salidas['FECHA_HORA'].max()
                porteria_salida = valid_salidas[valid_salidas['FECHA_HORA'] == salida_real]['porteria'].iloc[0]
                estado_calculo = "Calculado"
                salida_fue_real = True
                
            # --- PARA MICRO-JORNADAS ---
            if salida_fue_real:
                duracion_check = salida_real - entrada_real
                if duracion_check < timedelta(hours=MIN_DURACION_ACEPTABLE_REAL_SALIDA_HRS):
                    salida_real = fin_turno
                    porteria_salida = 'ASUMIDA (Micro-jornada detectada)'
                    estado_calculo = "ASUMIDO (Micro-jornada detectada)"
                    salida_fue_real = False

            # --- 3. REGLAS DE CÁLCULO DE HORAS ---

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
            estado_calculo = "Turno No Asignado (Entradas existen, pero ninguna se alinea con un turno programado)"
            return pd.DataFrame() # No se asignó turno -> Retorna vacío para probar la siguiente prioridad

    else:
        # No hay entradas en este subgrupo filtrado
        return pd.DataFrame()

    # --- Retorno del resultado (1 fila) ---
    report_date = fecha_clave_final if fecha_clave_final else fecha_clave_turno
    inicio_str = inicio_turno.strftime("%H:%M:%S") if inicio_turno else 'N/A'
    fin_str = fin_turno.strftime("%H:%M:%S") if fin_turno else 'N/A'
    horas_turno_val = info_turno["duracion_hrs"] if info_turno else 0
    ent_str = entrada_real.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(entrada_real) else 'N/A'
    sal_str = salida_real.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(salida_real) else 'N/A'

    resultado = {
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
    }
    
    return pd.DataFrame([resultado])


def calcular_turnos_con_prioridad(df: pd.DataFrame, lugares_pt_normalizados: list, lugares_p_normalizadas: list, tolerancia_llegada_tarde: int):
    """
    Función Maestra: Intenta calcular con PT. Si falla, intenta con P.
    """
    
    resultados_finales = []
    
    # Agrupa por ID de trabajador y por la fecha clave de la jornada (maneja turnos nocturnos)
    grupos = df.groupby(['id_trabajador', 'FECHA_CLAVE_TURNO'])
    
    for (id_trabajador, fecha_clave_turno), grupo in grupos:

        nombre = grupo['nombre'].iloc[0]
        
        # 1. INTENTO CON PUESTOS DE TRABAJO (PRIORIDAD 1)
        df_pt = grupo[grupo['PORTERIA_NORMALIZADA'].isin(lugares_pt_normalizados)].copy()
        df_resultado_pt = _ejecutar_calculo_para_grupo(df_pt, fecha_clave_turno, tolerancia_llegada_tarde)
        
        if not df_resultado_pt.empty:
            df_resultado_pt['Fuente_Marcacion'] = 'Puesto de Trabajo (PT)'
            resultados_finales.append(df_resultado_pt.iloc[0].to_dict())
            continue

        # 2. INTENTO CON PORTERÍAS (PRIORIDAD 2 - SOLO SI PT NO TUVO DATOS/NO ASIGNÓ TURNO)
        
        df_p = grupo[grupo['PORTERIA_NORMALIZADA'].isin(lugares_p_normalizadas)].copy()
        df_resultado_p = _ejecutar_calculo_para_grupo(df_p, fecha_clave_turno, tolerancia_llegada_tarde)
        
        if not df_resultado_p.empty:
            df_resultado_p['Fuente_Marcacion'] = 'Portería (P)'
            resultados_finales.append(df_resultado_p.iloc[0].to_dict())
            continue

        # 3. NO SE ENCONTRARON DATOS VÁLIDOS EN PT NI EN P
        
        report_date = fecha_clave_turno
        
        resultados_finales.append({
            'NOMBRE': nombre,
            'ID_TRABAJADOR': id_trabajador,
            'FECHA': report_date,
            'Dia_Semana': report_date.strftime('%A'),
            'TURNO': 'N/A',
            'Inicio_Turno_Programado': 'N/A',
            'Fin_Turno_Programado': 'N/A',
            'Duracion_Turno_Programado_Hrs': 0,
            'ENTRADA_REAL': 'N/A',
            'PORTERIA_ENTRADA': 'N/A',
            'SALIDA_REAL': 'N/A',
            'PORTERIA_SALIDA': 'N/A',
            'Horas_Trabajadas_Netas': 0.0,
            'Horas_Extra': 0.0,
            'Horas': 0,
            'Minutos': 0,
            'Llegada_Tarde_Mas_40_Min': False,
            'Estado_Calculo': "No hay marcaciones válidas en PT ni en P para asignar un turno",
            'Fuente_Marcacion': 'Ninguna'
        })
        
    if not resultados_finales:
        return pd.DataFrame()
        
    return pd.DataFrame(resultados_finales)

# --- 5. Interfaz Streamlit ---

st.set_page_config(page_title="Calculadora de Horas Extra", layout="wide")
st.title("📊 Calculadora de Horas Extra - NOEL")
st.write("Sube tu archivo de Excel para calcular las horas extra del personal. El sistema ahora **prioriza los Puestos de Trabajo (PT)**. Si no hay marcación válida en PT, usa las **Porterías (P)**.")

archivo_excel = st.file_uploader("Sube un archivo Excel (.xlsx)", type=["xlsx"])

if archivo_excel is not None:
    try:
        # Intenta leer la hoja específica
        # NOTA: En el código original se intentaba 'data', pero tu error KeyEror sugiere 'BaseDatos Modificada'
        # He ajustado a 'data' como en tu código, asumiendo que el error se corrigió o el usuario debe renombrar.
        df_raw = pd.read_excel(archivo_excel, sheet_name='data') 

        # 1. Definir la lista de nombres de columna que esperamos DESPUÉS de convertirlos a minúsculas
        columnas_requeridas_lower = [
            'cc', 'codtrabajador', 'nombre', 'fecha', 'hora', 'porteria', 'puntomarcacion'
        ]
        
        # 2. Mapeo y Normalización de columnas
        col_map = {col: col.lower() for col in df_raw.columns}
        df_raw.rename(columns=col_map, inplace=True)

        # 3. Validar la existencia de todas las columnas requeridas normalizadas.
        if not all(col in df_raw.columns for col in columnas_requeridas_lower):
            st.error(f"⚠️ ERROR: Faltan columnas requeridas o tienen nombres incorrectos. Asegúrate de tener: **Cc, CodTrabajador, Nombre, Fecha, Hora, Porteria, PuntoMarcacion** (en cualquier formato de mayúsculas/minúsculas).")
            st.stop()

        # 4. Seleccionar las columnas normalizadas y renombrar 'codtrabajador' a 'id_trabajador'.
        df_raw = df_raw[columnas_requeridas_lower].copy()
        df_raw.rename(columns={'codtrabajador': 'id_trabajador'}, inplace=True)
        
        # --- FILTRADO POR CÓDIGO DE TRABAJADOR ---
        try:
            df_raw['id_trabajador'] = pd.to_numeric(df_raw['id_trabajador'], errors='coerce').astype('Int64')
            codigos_filtro = CODIGOS_TRABAJADORES_FILTRO
        except:
            df_raw['id_trabajador'] = df_raw['id_trabajador'].astype(str)
            codigos_filtro = [str(c) for c in CODIGOS_TRABAJADORES_FILTRO]
        
        df_raw = df_raw[df_raw['id_trabajador'].isin(codigos_filtro)].copy()
        
        if df_raw.empty:
            st.error("⚠️ ERROR: Después del filtrado por código de trabajador, no quedan registros para procesar. Verifica que los códigos sean correctos.")
            st.stop()
        
        # Preprocesamiento inicial de columnas (usando 'fecha')
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
            if len(parts) == 2: return f"{time_str}:00"
            elif len(parts) == 3: return time_str
            else: return '00:00:00'

        # Aplica la estandarización y luego combina FECHA y HORA
        df_raw['hora'] = df_raw['hora'].apply(standardize_time_format)
            
        try:
            df_raw['FECHA_HORA'] = pd.to_datetime(df_raw['fecha'].dt.strftime('%Y-%m-%d') + ' ' + df_raw['hora'], errors='coerce')
            df_raw.dropna(subset=['FECHA_HORA'], inplace=True)
        except Exception as e:
            st.error(f"Error al combinar FECHA y HORA. Revisa el formato de la columna HORA: {e}")
            st.stop()

        # Normalización de las otras columnas de marcación
        df_raw['PORTERIA_NORMALIZADA'] = df_raw['porteria'].astype(str).str.strip().str.lower()
        df_raw['TIPO_MARCACION'] = df_raw['puntomarcacion'].astype(str).str.strip().str.lower().replace({'entrada': 'ent', 'salida': 'sal'})

        # --- Función para asignar Fecha Clave de Turno (Lógica Nocturna) ---
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

        st.success(f"✅ Archivo cargado y preprocesado con éxito. Se encontraron {len(df_raw['FECHA_CLAVE_TURNO'].unique())} días de jornada para procesar de {len(df_raw['id_trabajador'].unique())} trabajadores filtrados.")

        # --- Ejecutar el Cálculo con Prioridad ---
        df_resultado = calcular_turnos_con_prioridad(
            df_raw.copy(), 
            LUGARES_PUESTOS_TRABAJO_NORMALIZADOS,
            LUGARES_PORTERIAS_NORMALIZADAS,
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
                'Estado_Llegada', 'Estado_Calculo', 'Fuente_Marcacion' # Nueva columna
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
                orange_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                gray_format = workbook.add_format({'bg_color': '#D9D9D9'})
                yellow_format = workbook.add_format({'bg_color': '#FFF2CC', 'font_color': '#3C3C3C'})
                red_extra_format = workbook.add_format({'bg_color': '#F8E8E8', 'font_color': '#D83A56', 'bold': True})
                green_pt_format = workbook.add_format({'bg_color': '#E0FFE0'}) # Nuevo formato para PT

                # Aplica formatos condicionales basados en el dataframe original
                for row_num, row in df_resultado.iterrows():
                    excel_row = row_num + 1
                    
                    is_calculated = row['Estado_Calculo'] in ["Calculado", "ASUMIDO (Falta Salida/Salida Inválida)"]
                    is_late = row['Llegada_Tarde_Mas_40_Min']
                    is_assumed = row['Estado_Calculo'].startswith("ASUMIDO")
                    is_missing_entry = row['Estado_Calculo'].startswith("No hay marcaciones") or row['Estado_Calculo'].startswith("Turno No Asignado")
                    is_excessive_extra = row['Horas_Extra'] > UMBRAL_HORAS_EXTRA_RESALTAR
                    
                    # PASO 1: Determinar el formato base de la fila
                    base_format = None
                    if is_missing_entry and not is_assumed:
                        base_format = gray_format
                    elif is_assumed:
                        base_format = yellow_format
                    elif row['Fuente_Marcacion'] == 'Puesto de Trabajo (PT)':
                         base_format = green_pt_format

                    for col_idx, col_name in enumerate(df_to_excel.columns):
                        value = row[col_name]
                        cell_format = base_format
                        
                        # PASO 2: Aplicar Overrides
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
                file_name="Reporte_Marcacion_Horas_Extra.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("No se encontraron jornadas válidas después de aplicar los filtros.")

    except KeyError as e:
        st.error(f"⚠️ ERROR: Faltan columnas requeridas o tienen nombres incorrectos: {e}. Asegúrate que la hoja se llame 'data' y contenga las columnas requeridas.")
    except Exception as e:
        st.error(f"Error crítico al procesar el archivo: {e}. Por favor, verifica el formato de los datos.")

st.markdown("---")
st.caption("Somos NOEL DE CORAZÓN ❤️ - Herramienta de Cálculo de Turnos y Horas Extra")
