# -*- coding: utf-8 -*-
"""
Calculadora de Horas Extra y Reporte Diario de Marcaciones con Corrección de Descarga.
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
HORA_CORTE_NOCTURNO = datetime.strptime("08:00:00", "%H:%M:%S").time()

# Constante para la tolerancia de llegada tarde
TOLERANCIA_LLEGADA_TARDE_MINUTOS = 40

# --- 3. Obtener turno basado en fecha y hora (Sin Cambios) ---

def obtener_turno_para_registro(fecha_hora_evento: datetime, fecha_clave_turno_reporte: datetime.date, tolerancia_minutos: int):
    # ... (Cuerpo de la función 'obtener_turno_para_registro' sin cambios) ...
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
    menor_diferencia = timedelta(days=999) 

    for nombre_turno, info_turno in TURNOS[tipo_dia].items():
        hora_inicio = datetime.strptime(info_turno["inicio"], "%H:%M:%S").time()
        hora_fin = datetime.strptime(info_turno["fin"], "%H:%M:%S").time()

        candidatos_inicio = [datetime.combine(fecha_clave_turno_reporte, hora_inicio)]

        for inicio_posible_turno in candidatos_inicio:

            fin_posible_turno = inicio_posible_turno.replace(hour=hora_fin.hour, minute=hora_fin.minute, second=hora_fin.second)

            if hora_inicio > hora_fin:
                fin_posible_turno += timedelta(days=1) 

            rango_inicio = inicio_posible_turno - timedelta(minutes=tolerancia_minutos)
            rango_fin = fin_posible_turno + timedelta(minutes=tolerancia_minutos)

            if not (rango_inicio <= fecha_hora_evento <= rango_fin):
                continue

            diferencia = abs(fecha_hora_evento - inicio_posible_turno)

            if mejor_turno is None or diferencia < menor_diferencia:
                mejor_turno = (nombre_turno, info_turno, inicio_posible_turno, fin_posible_turno)
                menor_diferencia = diferencia

    return mejor_turno if mejor_turno else (None, None, None, None)

# --- 4. Calculo de horas (Lógica del reporte diario completo) ---

def calcular_turnos(df: pd.DataFrame, lugares_normalizados: list, tolerancia_minutos: int, tolerancia_llegada_tarde: int):

    """
    Función modificada para reportar CADA DÍA que cumple la regla de jornada mínima
    (aunque las Horas_Extra sean cero).
    """

    df = df[(df['PORTERIA_NORMALIZADA'].isin(lugares_normalizados)) & (df['TIPO_MARCACION'].isin(['ent', 'sal']))]

    df.sort_values(by=['ID_TRABAJADOR', 'FECHA_HORA'], inplace=True)

    if df.empty:
        return pd.DataFrame() 

    resultados = [] 

    for (id_trabajador, fecha_clave_turno), grupo in df.groupby(['ID_TRABAJADOR', 'FECHA_CLAVE_TURNO']):

        nombre = grupo['NOMBRE'].iloc[0] 
        entradas = grupo[grupo['TIPO_MARCACION'] == 'ent'] 
        salidas = grupo[grupo['TIPO_MARCACION'] == 'sal'] 

        # Regla 1: Si no hay entradas o salidas, se ignora el grupo
        if entradas.empty or salidas.empty:
            continue

        entrada_real = entradas['FECHA_HORA'].min()
        salida_real = salidas['FECHA_HORA'].max()

        porteria_entrada = entradas[entradas['FECHA_HORA'] == entrada_real]['PORTERIA'].iloc[0] if not entradas.empty else None
        porteria_salida = salidas[salidas['FECHA_HORA'] == salida_real]['PORTERIA'].iloc[0] if not salidas.empty else None

        # Regla 2: Si la duración total es menor a 4 horas, se ignora (criterio de jornada mínima)
        if salida_real <= entrada_real or (salida_real - entrada_real) < timedelta(hours=4):
            continue

        # Regla 3: Intenta asignar un turno programado a la jornada
        turno_nombre, info_turno, inicio_turno, fin_turno = obtener_turno_para_registro(entrada_real, fecha_clave_turno, tolerancia_minutos)
        if turno_nombre is None:
            continue # Si no se puede asignar un turno, se ignora el grupo

        # Regla 4: Valida que la salida real no exceda un límite razonable
        if salida_real > fin_turno + timedelta(hours=MAX_EXCESO_SALIDA_HRS):
            continue

        # --- Lógica de cálculo de horas ---
        
        # 1. Ajuste por llegada tarde
        inicio_efectivo_calculo = inicio_turno
        llegada_tarde_flag = False

        if entrada_real > inicio_turno:
            diferencia_entrada = entrada_real - inicio_turno
            if diferencia_entrada > timedelta(minutes=tolerancia_llegada_tarde):
                inicio_efectivo_calculo = entrada_real
                llegada_tarde_flag = True


        # 2. Cálculo de horas trabajadas y extra
        duracion_efectiva_calculo = salida_real - inicio_efectivo_calculo
        horas_trabajadas = round(duracion_efectiva_calculo.total_seconds() / 3600, 2) 

        horas_turno = info_turno["duracion_hrs"] 

        horas_extra = max(0, round(horas_trabajadas - horas_turno, 2))


        # Añade los resultados a la lista (ESTO INCLUYE DÍAS CON HORAS_EXTRA = 0)
        resultados.append({
            'NOMBRE': nombre,
            'ID_TRABAJADOR': id_trabajador,
            'FECHA': fecha_clave_turno, 
            'Dia_Semana': fecha_clave_turno.strftime('%A'), 
            'TURNO': turno_nombre,
            'Inicio_Turno_Programado': inicio_turno.strftime("%H:%M:%S"),
            'Fin_Turno_Programado': fin_turno.strftime("%H:%M:%S"),
            'Duracion_Turno_Programado_Hrs': horas_turno,
            'ENTRADA_REAL': entrada_real.strftime("%Y-%m-%d %H:%M:%S"), 
            'PORTERIA_ENTRADA': porteria_entrada,
            'SALIDA_REAL': salida_real.strftime("%Y-%m-%d %H:%M:%S"),
            'PORTERIA_SALIDA': porteria_salida,
            'Horas_Trabajadas': horas_trabajadas, 
            'Horas_Extra': horas_extra,
            'Horas': int(horas_extra),
            'Minutos': round((horas_extra - int(horas_extra)) * 60),
            # Esta columna se usará para el display y el formato
            'Estado_Llegada': 'Tarde' if llegada_tarde_flag else 'A tiempo',
            # Esta columna auxiliar se usará solo para el formato condicional de Excel
            'Llegada_Tarde_Aux': llegada_tarde_flag 
        })

    return pd.DataFrame(resultados) 

# --- Interfaz Streamlit (con corrección de descarga) ---
st.set_page_config(page_title="Calculadora de Horas Extra", layout="wide")
st.title("📊 Calculadora de Horas Trabajadas y Extra por Día")
st.write("Sube tu archivo de Excel para obtener el reporte de cada día marcado.")

archivo_excel = st.file_uploader("Sube un archivo Excel (.xlsx)", type=["xlsx"])

if archivo_excel is not None:
    try:
        df_raw = pd.read_excel(archivo_excel, sheet_name='BaseDatos Modificada')

        columnas = ['COD_TRABAJADOR', 'APUNTADOR', 'DESC_APUNTADOR', 'NOMBRE', 'FECHA', 'HORA', 'PORTERIA', 'PuntoMarcacion']
        if not all(col in df_raw.columns for col in columnas):
            st.error(f"ERROR: Faltan columnas requeridas: {', '.join(columnas)}")
        else:
            # Preprocesamiento inicial de columnas (igual que el código original)
            df_raw['FECHA'] = pd.to_datetime(df_raw['FECHA'])
            df_raw['HORA'] = df_raw['HORA'].astype(str)

            def standardize_time_format(time_str):
                parts = time_str.split(':')
                if len(parts) == 2: return f"{time_str}:00"
                elif len(parts) == 3: return time_str
                else: return time_str 

            df_raw['HORA'] = df_raw['HORA'].apply(standardize_time_format)
            
            df_raw['FECHA_HORA'] = pd.to_datetime(df_raw['FECHA'].dt.strftime('%Y-%m-%d') + ' ' + df_raw['HORA'])
            df_raw['PORTERIA_NORMALIZADA'] = df_raw['PORTERIA'].astype(str).str.strip().str.lower()
            df_raw['TIPO_MARCACION'] = df_raw['PuntoMarcacion'].astype(str).str.strip().str.lower().replace({'entrada': 'ent', 'salida': 'sal'})
            df_raw.rename(columns={'COD_TRABAJADOR': 'ID_TRABAJADOR'}, inplace=True)

            # LÓGICA: Asignar Fecha Clave de Turno (igual que el código original)
            def asignar_fecha_clave_turno(row):
                fecha_original = row['FECHA_HORA'].date()
                hora_marcacion = row['FECHA_HORA'].time()
                tipo_marcacion = row['TIPO_MARCACION'] 

                if tipo_marcacion == 'sal' and hora_marcacion < HORA_CORTE_NOCTURNO:
                    return fecha_original - timedelta(days=1)
                else:
                    return fecha_original

            df_raw['FECHA_CLAVE_TURNO'] = df_raw.apply(asignar_fecha_clave_turno, axis=1)

            st.success("Archivo cargado y preprocesado con éxito.")

            df_resultado = calcular_turnos(df_raw.copy(), LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS, TOLERANCIA_INFERENCIA_MINUTOS, TOLERANCIA_LLEGADA_TARDE_MINUTOS)

            if not df_resultado.empty:
                
                # Columnas a mostrar en el reporte (sin la auxiliar 'Llegada_Tarde_Aux')
                columnas_a_mostrar = [
                    'NOMBRE', 'ID_TRABAJADOR', 'FECHA', 'Dia_Semana', 'TURNO', 
                    'Inicio_Turno_Programado', 'Fin_Turno_Programado', 'Duracion_Turno_Programado_Hrs', 
                    'ENTRADA_REAL', 'SALIDA_REAL', 'PORTERIA_ENTRADA', 'PORTERIA_SALIDA',
                    'Horas_Trabajadas', 'Horas_Extra', 'Horas', 'Minutos', 
                    'Estado_Llegada'
                ]
                df_reporte_final = df_resultado[columnas_a_mostrar].copy()

                st.subheader("Reporte de Horas Trabajadas y Extra por Día")
                st.dataframe(df_reporte_final)

                # --- BLOQUE DE DESCARGA DE EXCEL CON CORRECCIÓN ---
                buffer_excel = io.BytesIO()
                
                # Abrimos el escritor de Excel
                with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
                    # Exporta el DataFrame sin la columna auxiliar al Excel
                    df_reporte_final.to_excel(writer, sheet_name='Reporte Diario', index=False)

                    workbook = writer.book
                    worksheet = writer.sheets['Reporte Diario']
                    orange_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

                    # Obtener el índice de la columna ENTRADA_REAL en el reporte final
                    try:
                        entrada_real_col_idx = df_reporte_final.columns.get_loc('ENTRADA_REAL')
                    except KeyError:
                        entrada_real_col_idx = -1

                    if entrada_real_col_idx != -1:
                        # Aplica formato condicional usando la columna auxiliar original
                        for row_num, is_late in enumerate(df_resultado['Llegada_Tarde_Aux']): 
                            if is_late:
                                worksheet.write(row_num + 1, entrada_real_col_idx, df_resultado.iloc[row_num]['ENTRADA_REAL'], orange_format)
                            else:
                                worksheet.write(row_num + 1, entrada_real_col_idx, df_resultado.iloc[row_num]['ENTRADA_REAL'])

                # *** CORRECCIÓN DE DESCARGA: Mover el cursor del buffer al inicio ***
                buffer_excel.seek(0)

                # Botón de descarga para el usuario
                st.download_button(
                    label="Descargar Reporte Diario Completo (Excel)",
                    data=buffer_excel,
                    file_name="Reporte_Diario_Horas_Trabajadas_y_Extra.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("No se pudieron asignar turnos o hubo inconsistencias en los registros que cumplieran los criterios mínimos de jornada (ej. duración de la jornada menor a 4 horas).")

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}. Asegúrate de que la hoja se llama 'BaseDatos Modificada' y que tiene todas las columnas requeridas.")

st.markdown("---")
st.caption("Somos NOEL DE CORAZÓN ❤️")

