# -*- coding: utf-8 -*-
"""
Calculadora de Horas Extra.
Versión Final: Implementa Min/Max, Prioriza la Entrada del Turno y Asume Salida si falta la marcación.
"""

import pandas as pd
from datetime import datetime, timedelta
import streamlit as st
import io
import numpy as np

# --- 1. Definición de los Turnos ---
# NOTA: Los horarios de los turnos definen el rango de búsqueda para el turno más cercano a la entrada real.
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
    "NOEL_MDE_MR_MEZCLAS_ENT", "NOEL_MDE_OFIC_PRODUCCION_SAL", "NOEL_MDE_OFIC_PRODUCCION_ENT",
    # Nuevas porterías peatonales añadidas
    "NOEL_MDE_PORT_2_PEATONAL_1_ENT",
    "NOEL_MDE_PORT_2_PEATONAL_1_SAL",
    "NOEL_MDE_PORT_2_PEATONAL_2_ENT",
    "NOEL_MDE_PORT_2_PEATONAL_2_SAL",
    "NOEL_MDE_PORT_2_PEATONAL_3_SAL",
    "NOEL_MDE_PORT_2_PEATONAL_3_ENT",
    "NOEL_MDE_PORT_1_PEATONAL_1_ENT"
]

LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS = [lugar.strip().lower() for lugar in LUGARES_TRABAJO_PRINCIPAL]

# Tolerancia en minutos para buscar el turno programado MÁS CERCANO (50 min).
TOLERANCIA_INFERENCIA_MINUTOS = 50 
# Máximo de horas para aceptar una ENTRADA muy temprana (6 horas antes del inicio programado).
MAX_ANTICIPACION_ENTRADA_HRS = 6 
# Máximo de horas después del fin de turno programado que se acepta una salida como válida.
MAX_EXCESO_SALIDA_HRS = 3 
# Hora de corte para definir si una SALIDA matutina pertenece al turno del día anterior (ej: 05:40 AM)
HORA_CORTE_NOCTURNO = datetime.strptime("08:00:00", "%H:%M:%S").time() 
# Tolerancia para considerar la llegada como 'tarde' para el cálculo de horas.
TOLERANCIA_LLEGADA_TARDE_MINUTOS = 40
# Tolerancia que usaba la regla anterior, no usada actualmente.
TOLERANCIA_ENTRADA_TEMPRANA_MINUTOS = 30 

# --- 3. Obtener turno basado en fecha y hora ---

def obtener_turno_para_registro(fecha_hora_evento: datetime, fecha_clave_turno_reporte: datetime.date, tolerancia_minutos: int):
    """ 
    Busca el turno programado más cercano a la marcación de entrada (primera y más temprana) 
    usando la FECHA_CLAVE_TURNO. 
    """
    dia_semana_clave = fecha_clave_turno_reporte.weekday()

    if dia_semana_clave < 5: tipo_dia = "LV"
    elif dia_semana_clave == 5: tipo_dia = "SAB"
    else: tipo_dia = "DOM"

    if tipo_dia not in TURNOS: return (None, None, None, None)

    mejor_turno = None
    menor_diferencia = timedelta(days=999)

    for nombre_turno, info_turno in TURNOS[tipo_dia].items():
        hora_inicio = datetime.strptime(info_turno["inicio"], "%H:%M:%S").time()
        hora_fin = datetime.strptime(info_turno["fin"], "%H:%M:%S").time()
        es_nocturno = info_turno.get("nocturno", False)

        # La hora de inicio programada se combina con la FECHA CLAVE
        inicio_posible_turno = datetime.combine(fecha_clave_turno_reporte, hora_inicio)

        if es_nocturno:
            # Si es nocturno, el fin del turno ocurre al día siguiente
            fin_posible_turno = datetime.combine(fecha_clave_turno_reporte + timedelta(days=1), hora_fin)
        else:
            fin_posible_turno = datetime.combine(fecha_clave_turno_reporte, hora_fin)

        # Rango de tolerancia amplio para la entrada (para capturar entradas muy tempranas).
        max_anticipacion = timedelta(hours=MAX_ANTICIPACION_ENTRADA_HRS)
        rango_inicio_aceptacion = inicio_posible_turno - max_anticipacion
        
        # Validar si el evento (la entrada) cae en el rango amplio definido.
        if fecha_hora_evento >= rango_inicio_aceptacion and fecha_hora_evento <= fin_posible_turno + timedelta(hours=MAX_EXCESO_SALIDA_HRS):
            
            # La diferencia se calcula entre la entrada real y el inicio PROGRAMADO del turno
            diferencia = abs(fecha_hora_evento - inicio_posible_turno)

            if mejor_turno is None or diferencia < menor_diferencia:
                mejor_turno = (nombre_turno, info_turno, inicio_posible_turno, fin_posible_turno)
                menor_diferencia = diferencia

    return mejor_turno if mejor_turno else (None, None, None, None)

# --- 4. Calculo de horas (Selección de Min/Max y Priorización de Turno) ---

def calcular_turnos(df: pd.DataFrame, lugares_normalizados: list, tolerancia_minutos: int, tolerancia_llegada_tarde: int):
    """
    Agrupa por ID y FECHA_CLAVE_TURNO. 
    Prioriza la ENTRADA que mejor se alinea a un turno programado.
    Toma la SALIDA MÁS TARDÍA válida posterior a esa entrada o ASUME el Fin de Turno si falta.
    """
    df_filtrado = df[(df['PORTERIA_NORMALIZADA'].isin(lugares_normalizados)) & (df['TIPO_MARCACION'].isin(['ent', 'sal']))].copy()
    df_filtrado.sort_values(by=['ID_TRABAJADOR', 'FECHA_HORA'], inplace=True)

    if df_filtrado.empty: return pd.DataFrame()

    resultados = []

    # Agrupa por ID de trabajador y por la fecha clave de la jornada (maneja turnos nocturnos)
    for (id_trabajador, fecha_clave_turno), grupo in df_filtrado.groupby(['ID_TRABAJADOR', 'FECHA_CLAVE_TURNO']):

        nombre = grupo['NOMBRE'].iloc[0]
        # Aseguramos que las entradas y salidas dentro del grupo se ordenen por hora
        entradas = grupo[grupo['TIPO_MARCACION'] == 'ent'].sort_values(by='FECHA_HORA')
        salidas = grupo[grupo['TIPO_MARCACION'] == 'sal'].sort_values(by='FECHA_HORA')

        entrada_real = pd.NaT
        porteria_entrada = 'N/A'
        salida_real = pd.NaT
        porteria_salida = 'N/A'
        turno_nombre, info_turno, inicio_turno, fin_turno = (None, None, None, None)
        horas_trabajadas = 0.0
        horas_extra = 0.0
        llegada_tarde_flag = False
        estado_calculo = "Sin Marcaciones Válidas (E/S)"
        
        mejor_entrada_para_turno = pd.NaT
        mejor_turno_data = (None, None, None, None)
        menor_diferencia_segundos = 999999999 # Usamos segundos para una comparación segura

        # --- REVISIÓN CLAVE 1: Encontrar la mejor entrada que se alinee a un turno ---
        if not entradas.empty:
            for index, row in entradas.iterrows():
                current_entry_time = row['FECHA_HORA']
                
                # Intentar asignar un turno a esta marcación de entrada. 
                turno_nombre_temp, info_turno_temp, inicio_turno_temp, fin_turno_temp = obtener_turno_para_registro(current_entry_time, fecha_clave_turno, tolerancia_minutos)
                
                if turno_nombre_temp is not None:
                    # Calcula la diferencia absoluta con el inicio programado (para encontrar el mejor ajuste)
                    diferencia = abs(current_entry_time - inicio_turno_temp)
                    diferencia_segundos = diferencia.total_seconds()
                    
                    if diferencia_segundos < menor_diferencia_segundos:
                        menor_diferencia_segundos = diferencia_segundos
                        mejor_entrada_para_turno = current_entry_time
                        mejor_turno_data = (turno_nombre_temp, info_turno_temp, inicio_turno_temp, fin_turno_temp)
                    # Tie-breaker: Si la diferencia es la misma, preferir la entrada MÁS TEMPRANA.
                    elif diferencia_segundos == menor_diferencia_segundos and current_entry_time < mejor_entrada_para_turno:
                        mejor_entrada_para_turno = current_entry_time
                        mejor_turno_data = (turno_nombre_temp, info_turno_temp, inicio_turno_temp, fin_turno_temp)


            # Si se encontró un turno asociado a la mejor entrada
            if pd.notna(mejor_entrada_para_turno):
                turno_nombre, info_turno, inicio_turno, fin_turno = mejor_turno_data
                
                # *** NUEVA REGLA: Prevención de Asignación de Turno Diurno a Marcación Matutina Accidental ***
                # Si el mejor turno encontrado NO es un Turno 3 (nocturno) Y la marcación de entrada es ANTES de las 05:40 AM,
                # esta marcación es probablemente accesoria de un turno nocturno cuyo inicio no está en el archivo.
                # Rechazamos la asignación para evitar falsos Turno 1.
                if not turno_nombre.startswith("Turno 3") and mejor_entrada_para_turno.time() < datetime.strptime("05:40:00", "%H:%M:%S").time():
                    estado_calculo = "Entrada Matutina No Alineada a Turno Nocturno (Ignorado)"
                    # Reseteamos los datos para caer en el estado de "Turno No Asignado"
                    mejor_entrada_para_turno = pd.NaT 
                    mejor_turno_data = (None, None, None, None) 
                
            # Continuar solo si el turno no fue invalidado
            if pd.notna(mejor_entrada_para_turno):
                entrada_real = mejor_entrada_para_turno
                turno_nombre, info_turno, inicio_turno, fin_turno = mejor_turno_data
                
                # Obtener porteria de la entrada real
                porteria_entrada = entradas[entradas['FECHA_HORA'] == entrada_real]['PORTERIA'].iloc[0]
                
                # --- REVISIÓN CLAVE 2: Filtro y/o Inferencia de Salida ---
                
                # Calcula el límite máximo de salida aceptable
                max_salida_aceptable = fin_turno + timedelta(hours=MAX_EXCESO_SALIDA_HRS)
                
                # Filtra las salidas que ocurrieron DESPUÉS de la ENTRADA REAL seleccionada y DENTRO del límite aceptable
                valid_salidas = salidas[
                    (salidas['FECHA_HORA'] > entrada_real) & 
                    (salidas['FECHA_HORA'] <= max_salida_aceptable)
                ]
                
                if valid_salidas.empty:
                    # SI NO HAY SALIDA VÁLIDA: ASUMIR SALIDA A LA HORA PROGRAMADA DEL FIN DE TURNO
                    salida_real = fin_turno  
                    porteria_salida = 'ASUMIDA (FIN TURNO)'
                    estado_calculo = "ASUMIDO (Falta Salida/Salida Inválida)"
                else:
                    # Usar la última salida REAL válida
                    salida_real = valid_salidas['FECHA_HORA'].max()
                    porteria_salida = valid_salidas[valid_salidas['FECHA_HORA'] == salida_real]['PORTERIA'].iloc[0]
                    estado_calculo = "Calculado"
                    
                # --- 3. REGLAS DE CÁLCULO DE HORAS ---

                duracion_total = salida_real - entrada_real
                
                # Regla de cálculo por defecto: inicia en el turno programado
                inicio_efectivo_calculo = inicio_turno
                llegada_tarde_flag = False
                
                # 1. Regla para LLEGADA TARDE (Más de 40 minutos tarde) - Tiene prioridad
                if entrada_real > inicio_turno + timedelta(minutes=tolerancia_llegada_tarde):
                    # Si llega tarde más la tolerancia (40 min), el cálculo inicia en la entrada real
                    inicio_efectivo_calculo = entrada_real
                    llegada_tarde_flag = True
                    
                # 2. Regla para ENTRADA TEMPRANA (Cualquier entrada antes del inicio programado, si no es llegada tarde)
                elif entrada_real < inicio_turno:
                    # Se cuenta desde la hora de entrada real. Esto calcula la hora extra de entrada.
                    inicio_efectivo_calculo = entrada_real
                
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
                    # Mantenemos el estado de cálculo original (Calculado, Asumido o Jornada Corta)
                    if duracion_total < timedelta(hours=4) and estado_calculo not in ["ASUMIDO (Falta Salida/Salida Inválida)", "Calculado"]:
                        estado_calculo = "Jornada Corta (< 4h de Ent-Sal)"
                        horas_extra = 0.0
                    else:
                        horas_extra = max(0, round(horas_trabajadas - horas_turno, 2))
                        if estado_calculo == "Calculado":
                            estado_calculo = "Calculado" 


            else:
                # Cae aquí si la entrada fue invalidada por la nueva regla o no se encontró ninguna entrada coincidente
                estado_calculo = "Turno No Asignado (Entradas existen, pero ninguna se alinea con un turno programado)"
                # Se mantiene el estado de "Entrada Matutina No Alineada..." si fue el caso.
                if "Entrada Matutina" in estado_calculo:
                    pass
                else:
                    estado_calculo = "Turno No Asignado (Entradas existen, pero ninguna se alinea con un turno programado)"

        elif pd.isna(entrada_real) and not salidas.empty:
            estado_calculo = "Falta Entrada (Salida marcada)"
            
        # --- Añade los resultados a la lista (Se reporta todo) ---
        ent_str = entrada_real.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(entrada_real) else 'N/A'
        sal_str = salida_real.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(salida_real) else 'N/A'
        inicio_str = inicio_turno.strftime("%H:%M:%S") if inicio_turno else 'N/A'
        fin_str = fin_turno.strftime("%H:%M:%S") if fin_turno else 'N/A'
        horas_turno_val = info_turno["duracion_hrs"] if info_turno else 0

        resultados.append({
            'NOMBRE': nombre,
            'ID_TRABAJADOR': id_trabajador,
            'FECHA': fecha_clave_turno,
            'Dia_Semana': fecha_clave_turno.strftime('%A'),
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
st.write("Sube tu archivo de Excel para calcular las horas extra del personal. El sistema toma la **Entrada más cercana al turno programado** y la **Salida más tardía válida** posterior a esa entrada.")

archivo_excel = st.file_uploader("Sube un archivo Excel (.xlsx)", type=["xlsx"])

if archivo_excel is not None:
    try:
        # Intenta leer la hoja específica 'BaseDatos Modificada'
        df_raw = pd.read_excel(archivo_excel, sheet_name='BaseDatos Modificada')

        columnas = ['COD_TRABAJADOR', 'NOMBRE', 'FECHA', 'HORA', 'PORTERIA', 'PuntoMarcacion']
        if not all(col in df_raw.columns for col in columnas):
            st.error(f"⚠️ ERROR: Faltan columnas requeridas o tienen nombres incorrectos. Asegúrate de tener: **COD_TRABAJADOR**, **NOMBRE**, **FECHA**, **HORA**, **PORTERIA**, **PuntoMarcacion**.")
        else:
            # Preprocesamiento inicial de columnas
            df_raw['FECHA'] = pd.to_datetime(df_raw['FECHA'], errors='coerce')  
            df_raw.dropna(subset=['FECHA'], inplace=True)
            
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
            df_raw['HORA'] = df_raw['HORA'].apply(standardize_time_format)
            
            try:
                df_raw['FECHA_HORA'] = pd.to_datetime(df_raw['FECHA'].dt.strftime('%Y-%m-%d') + ' ' + df_raw['HORA'], errors='coerce')
                df_raw.dropna(subset=['FECHA_HORA'], inplace=True)
            except Exception as e:
                st.error(f"Error al combinar FECHA y HORA. Revisa el formato de la columna HORA: {e}")
                st.stop() 

            df_raw['PORTERIA_NORMALIZADA'] = df_raw['PORTERIA'].astype(str).str.strip().str.lower()
            # Mapeo de PuntoMarcacion a 'ent' o 'sal'
            df_raw['TIPO_MARCACION'] = df_raw['PuntoMarcacion'].astype(str).str.strip().str.lower().replace({'entrada': 'ent', 'salida': 'sal'})
            df_raw.rename(columns={'COD_TRABAJADOR': 'ID_TRABAJADOR'}, inplace=True)

            # --- Función para asignar Fecha Clave de Turno (Lógica Nocturna) ---
            def asignar_fecha_clave_turno_corregida(row):
                fecha_original = row['FECHA_HORA'].date()
                hora_marcacion = row['FECHA_HORA'].time()
                tipo_marcacion = row['TIPO_MARCACION']
                
                # Regla de oro: Las ENTRADAS anclan la jornada al día en que ocurrieron.
                if tipo_marcacion == 'ent':
                    return fecha_original
                
                # Regla nocturna: Las SALIDAS antes del corte se asocian al turno del día anterior.
                # Esto es crucial para agrupar Entrada (Día 1 Noche) y Salida (Día 2 Madrugada).
                if tipo_marcacion == 'sal' and hora_marcacion < HORA_CORTE_NOCTURNO:
                    return fecha_original - timedelta(days=1)
                
                # Otras salidas (después de 8 AM) pertenecen al día en que fueron marcadas.
                return fecha_original

            df_raw['FECHA_CLAVE_TURNO'] = df_raw.apply(asignar_fecha_clave_turno_corregida, axis=1)

            st.success(f"✅ Archivo cargado y preprocesado con éxito. Se encontraron {len(df_raw['FECHA_CLAVE_TURNO'].unique())} días de jornada para procesar.")

            # --- Ejecutar el Cálculo ---
            df_resultado = calcular_turnos(df_raw.copy(), LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS, TOLERANCIA_INFERENCIA_MINUTOS, TOLERANCIA_LLEGADA_TARDE_MINUTOS)

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
                    orange_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}) # Tarde
                    gray_format = workbook.add_format({'bg_color': '#D9D9D9'}) # No calculado
                    yellow_format = workbook.add_format({'bg_color': '#FFF2CC', 'font_color': '#3C3C3C'}) # Asumido
                    
                    # Aplica formatos condicionales basados en el dataframe original
                    for row_num, row in df_resultado.iterrows():
                        excel_row = row_num + 1
                        
                        is_calculated = row['Estado_Calculo'] in ["Calculado", "ASUMIDO (Falta Salida/Salida Inválida)"]
                        is_late = row['Llegada_Tarde_Mas_40_Min']
                        is_assumed = row['Estado_Calculo'].startswith("ASUMIDO") or row['Estado_Calculo'].startswith("Entrada Matutina")

                        for col_idx, col_name in enumerate(df_to_excel.columns):
                            value = row[col_name]
                            cell_format = None
                            
                            # Prioridad 1: No calculado / Ignorado (gris)
                            if row['Estado_Calculo'] in ["Turno No Asignado (Entradas existen, pero ninguna se alinea con un turno programado)", "Entrada Matutina No Alineada a Turno Nocturno (Ignorado)", "Error: Duración efectiva negativa"]:
                                cell_format = gray_format
                            # Prioridad 2: Asumido (amarillo claro)
                            elif is_assumed:
                                cell_format = yellow_format
                            # Prioridad 3: Llegada Tarde (naranja/rojo)
                            elif col_name == 'ENTRADA_REAL' and is_late:
                                cell_format = orange_format

                            # Escribir el valor en la celda
                            worksheet.write(excel_row, col_idx, value if pd.notna(value) else 'N/A', cell_format)

                buffer_excel.seek(0)

                st.download_button(
                    label="Descargar Reporte de Horas Extra (Excel)",
                    data=buffer_excel,
                    file_name="Reporte_Marcacion_Horas_Extra.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("No se encontraron jornadas válidas después de aplicar los filtros.")

    except KeyError:
        st.error(f"⚠️ ERROR: El archivo Excel debe contener una hoja llamada **'BaseDatos Modificada'** y las columnas requeridas.")
    except Exception as e:
        st.error(f"Error crítico al procesar el archivo: {e}. Por favor, verifica el formato de los datos.")

st.markdown("---")
st.caption("Somos NOEL DE CORAZÓN ❤️ - Herramienta de Cálculo de Turnos y Horas Extra")


