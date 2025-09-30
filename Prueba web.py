# -*- coding: utf-8 -*-
"""
Calculadora de Horas Extra.
Corregido: Manejo de turnos nocturnos y SyntaxError en Streamlit.
"""

import pandas as pd
from datetime import datetime, timedelta
import streamlit as st
import io
import numpy as np

# --- 1. Definición de los Turnos ---
# Se añade 'nocturno': True para facilitar la identificación del cruce de medianoche.
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
]

LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS = [lugar.strip().lower() for lugar in LUGARES_TRABAJO_PRINCIPAL]

TOLERANCIA_INFERENCIA_MINUTOS = 50
MAX_EXCESO_SALIDA_HRS = 3
HORA_CORTE_NOCTURNO = datetime.strptime("08:00:00", "%H:%M:%S").time() # 8 AM
TOLERANCIA_LLEGADA_TARDE_MINUTOS = 40

# --- 3. Obtener turno basado en fecha y hora ---

def obtener_turno_para_registro(fecha_hora_evento: datetime, fecha_clave_turno_reporte: datetime.date, tolerancia_minutos: int):

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

        inicio_posible_turno = datetime.combine(fecha_clave_turno_reporte, hora_inicio)

        # Ajusta el fin si es turno nocturno (cruce de medianoche)
        if es_nocturno:
            fin_posible_turno = datetime.combine(fecha_clave_turno_reporte + timedelta(days=1), hora_fin)
        else:
            fin_posible_turno = datetime.combine(fecha_clave_turno_reporte, hora_fin)

        rango_inicio = inicio_posible_turno - timedelta(minutes=tolerancia_minutos)
        rango_fin = fin_posible_turno + timedelta(minutes=tolerancia_minutos)

        # La entrada real debe estar cerca del inicio programado (con tolerancia)
        if not (rango_inicio <= fecha_hora_evento <= rango_fin):
            continue

        diferencia = abs(fecha_hora_evento - inicio_posible_turno)

        if mejor_turno is None or diferencia < menor_diferencia:
            mejor_turno = (nombre_turno, info_turno, inicio_posible_turno, fin_posible_turno)
            menor_diferencia = diferencia

    return mejor_turno if mejor_turno else (None, None, None, None)

# --- 4. Calculo de horas ---

def calcular_turnos(df: pd.DataFrame, lugares_normalizados: list, tolerancia_minutos: int, tolerancia_llegada_tarde: int):

    # Filtra por lugares principales y tipos 'ent'/'sal'
    df_filtrado = df[(df['PORTERIA_NORMALIZADA'].isin(lugares_normalizados)) & (df['TIPO_MARCACION'].isin(['ent', 'sal']))].copy()
    df_filtrado.sort_values(by=['ID_TRABAJADOR', 'FECHA_HORA'], inplace=True)

    if df_filtrado.empty: return pd.DataFrame()

    resultados = []

    # Agrupa por ID de trabajador y por fecha clave turno
    for (id_trabajador, fecha_clave_turno), grupo in df_filtrado.groupby(['ID_TRABAJADOR', 'FECHA_CLAVE_TURNO']):

        nombre = grupo['NOMBRE'].iloc[0]
        entradas = grupo[grupo['TIPO_MARCACION'] == 'ent']
        salidas = grupo[grupo['TIPO_MARCACION'] == 'sal']

        # Obtiene la primera entrada y la última salida
        entrada_real = entradas['FECHA_HORA'].min() if not entradas.empty else pd.NaT
        salida_real = salidas['FECHA_HORA'].max() if not salidas.empty else pd.NaT

        porteria_entrada = entradas[entradas['FECHA_HORA'] == entrada_real]['PORTERIA'].iloc[0] if not entradas.empty and pd.notna(entrada_real) else 'Sin Entrada'
        porteria_salida = salidas[salidas['FECHA_HORA'] == salida_real]['PORTERIA'].iloc[0] if not salidas.empty and pd.notna(salida_real) else 'Sin Salida'

        turno_nombre, info_turno, inicio_turno, fin_turno = (None, None, None, None)
        horas_trabajadas = 0.0
        horas_extra = 0.0
        llegada_tarde_flag = False
        estado_calculo = "No Calculado"

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
                        # --- Lógica de cálculo de horas ---
                        inicio_efectivo_calculo = inicio_turno
                        
                        if entrada_real > inicio_turno:
                            diferencia_entrada = entrada_real - inicio_turno
                            if diferencia_entrada > timedelta(minutes=tolerancia_llegada_tarde):
                                # Si llega tarde más la tolerancia, el cálculo inicia en la entrada real
                                inicio_efectivo_calculo = entrada_real 
                                llegada_tarde_flag = True
                        
                        duracion_efectiva_calculo = salida_real - inicio_efectivo_calculo
                        horas_trabajadas = round(duracion_efectiva_calculo.total_seconds() / 3600, 2)
                        
                        horas_turno = info_turno["duracion_hrs"]
                        horas_extra = max(0, round(horas_trabajadas - horas_turno, 2))
                        estado_calculo = "Calculado"

        elif pd.notna(entrada_real) and pd.isna(salida_real):
             estado_calculo = "Falta Salida"
        elif pd.isna(entrada_real) and pd.notna(salida_real):
             estado_calculo = "Falta Entrada"
        else:
             estado_calculo = "Sin Marcaciones Válidas"
             
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
            'Horas_Trabajadas': horas_trabajadas,
            'Horas_Extra': horas_extra,
            'Horas': int(horas_extra),
            'Minutos': round((horas_extra - int(horas_extra)) * 60),
            'Llegada_Tarde_Mas_40_Min': llegada_tarde_flag,
            'Estado_Calculo': estado_calculo
        })

    return pd.DataFrame(resultados)

# --- 5. Interfaz Streamlit (Corregida la Indentación) ---

st.set_page_config(page_title="Calculadora de Horas Extra", layout="wide")
st.title("📊 Calculadora de Horas Extra")
st.write("Sube tu archivo de Excel para calcular las horas extra del personal.")

archivo_excel = st.file_uploader("Sube un archivo Excel (.xlsx)", type=["xlsx"])

if archivo_excel is not None:
    try:
        # Intenta leer la hoja 'BaseDatos Modificada'
        df_raw = pd.read_excel(archivo_excel, sheet_name='BaseDatos Modificada')

        columnas = ['COD_TRABAJADOR', 'NOMBRE', 'FECHA', 'HORA', 'PORTERIA', 'PuntoMarcacion']
        if not all(col in df_raw.columns for col in columnas):
            st.error(f"ERROR: Faltan columnas requeridas o tienen nombres incorrectos. Asegúrate de tener: **COD_TRABAJADOR**, **NOMBRE**, **FECHA**, **HORA**, **PORTERIA**, **PuntoMarcacion**.")
        else:
            # Preprocesamiento inicial de columnas
            df_raw['FECHA'] = pd.to_datetime(df_raw['FECHA'], errors='coerce') 
            df_raw.dropna(subset=['FECHA'], inplace=True)
            df_raw['HORA'] = df_raw['HORA'].astype(str)

            # --- Función para estandarizar el formato de la hora (maneja floats y strings) ---
            def standardize_time_format(time_val):
                if isinstance(time_val, float) and time_val <= 1.0: # Asume formato de Excel (fracción de día)
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
                    return '00:00:00' # Valor por defecto seguro para horas no detectadas

            df_raw['HORA'] = df_raw['HORA'].apply(standardize_time_format)
            
            # Combina FECHA y HORA
            try:
                df_raw['FECHA_HORA'] = pd.to_datetime(df_raw['FECHA'].dt.strftime('%Y-%m-%d') + ' ' + df_raw['HORA'], errors='coerce')
                df_raw.dropna(subset=['FECHA_HORA'], inplace=True)
            except Exception as e:
                 st.error(f"Error al combinar FECHA y HORA. Revisa el formato de la columna HORA: {e}")
                 return

            df_raw['PORTERIA_NORMALIZADA'] = df_raw['PORTERIA'].astype(str).str.strip().str.lower()
            df_raw['TIPO_MARCACION'] = df_raw['PuntoMarcacion'].astype(str).str.strip().str.lower().replace({'entrada': 'ent', 'salida': 'sal'})
            df_raw.rename(columns={'COD_TRABAJADOR': 'ID_TRABAJADOR'}, inplace=True)

            # --- Función para asignar Fecha Clave de Turno (Lógica Nocturna) ---
            
            def asignar_fecha_clave_turno_corregida(row):
                fecha_original = row['FECHA_HORA'].date()
                hora_marcacion = row['FECHA_HORA'].time()
                
                # Si la marcación es antes de 8 AM, se asocia al turno del día anterior
                if hora_marcacion < HORA_CORTE_NOCTURNO:
                    return fecha_original - timedelta(days=1)
                
                # De lo contrario, se asocia al turno de ese mismo día.
                else:
                    return fecha_original

            df_raw['FECHA_CLAVE_TURNO'] = df_raw.apply(asignar_fecha_clave_turno_corregida, axis=1)

            st.success(f"Archivo cargado y preprocesado con éxito. Se encontraron {len(df_raw['FECHA_CLAVE_TURNO'].unique())} días de jornada para procesar.")

            # --- Llamada a la función principal de cálculo ---
            df_resultado = calcular_turnos(df_raw.copy(), LUGARES_TRABAJO_PRINCIPAL_NORMALIZADOS, TOLERANCIA_INFERENCIA_MINUTOS, TOLERANCIA_LLEGADA_TARDE_MINUTOS)

            if not df_resultado.empty:
                # Post-procesamiento para el reporte
                df_resultado['Llegada_Tarde'] = df_resultado['Llegada_Tarde_Mas_40_Min']
                df_resultado.rename(columns={'Llegada_Tarde_Mas_40_Min': 'Estado_Llegada'}, inplace=True)
                df_resultado['Estado_Llegada'] = df_resultado['Estado_Llegada'].map({True: 'Tarde', False: 'A tiempo'})
                df_resultado.sort_values(by=['NOMBRE', 'FECHA'], inplace=True)

                st.subheader("Resultados de las horas extra")
                st.dataframe(df_resultado.drop(columns=['Llegada_Tarde']))

                # --- Lógica de descarga en Excel con formato condicional ---
                buffer_excel = io.BytesIO()
                with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
                    df_to_excel = df_resultado.drop(columns=['Llegada_Tarde']).copy()
                    df_to_excel.to_excel(writer, sheet_name='Reporte Horas Extra', index=False)

                    workbook = writer.book
                    worksheet = writer.sheets['Reporte Horas Extra']

                    orange_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                    gray_format = workbook.add_format({'bg_color': '#D9D9D9'})
                    
                    try: entrada_real_col_idx = df_to_excel.columns.get_loc('ENTRADA_REAL')
                    except KeyError: entrada_real_col_idx = -1

                    for row_num, row in df_resultado.iterrows():
                        excel_row = row_num + 1
                        is_calculated = row['Estado_Calculo'] == "Calculado"
                        
                        for col_idx, col_name in enumerate(df_to_excel.columns):
                            value = row[col_name]
                            cell_format = None
                            
                            if not is_calculated:
                                cell_format = gray_format
                            elif col_name == 'ENTRADA_REAL' and row['Llegada_Tarde']:
                                cell_format = orange_format

                            if pd.isna(value) or value in ['N/A', 'Sin Entrada', 'Sin Salida']:
                                value_to_write = str(value)
                            elif col_name in ['FECHA'] and isinstance(value, (datetime, pd.Timestamp)):
                                value_to_write = value.strftime("%Y-%m-%d")
                            elif col_name in ['ENTRADA_REAL', 'SALIDA_REAL'] and isinstance(value, (str)):
                                # Escribe la cadena de texto de la fecha/hora
                                value_to_write = value
                            else:
                                value_to_write = value

                            worksheet.write(excel_row, col_idx, value_to_write, cell_format)

                buffer_excel.seek(0)

                st.download_button(
                    label="Descargar Reporte de Horas extra (Excel)",
                    data=buffer_excel,
                    file_name="Marcación_horas_extra.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("No se encontraron jornadas válidas después de aplicar los filtros.")

    except Exception as e:
        # Se asegura que la excepción se maneje dentro del bloque.
        st.error(f"Error crítico al procesar el archivo: {e}. Por favor, verifica el nombre de la hoja ('BaseDatos Modificada') y los formatos de las columnas, especialmente **FECHA** y **HORA**.")

st.markdown("---")
st.caption("Somos NOEL DE CORAZÓN ❤️")

