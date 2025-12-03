
import pandas as pd
from datetime import datetime
import streamlit as st
from io import BytesIO

# Tabla completa de recargos según ley laboral
PORCENTAJES = {
    'Hora extra diurna': '25%',
    'Hora extra nocturna': '75%',
    'Hora extra diurna en domingo o festivo': '105%',
    'Hora extra nocturna en domingo o festivo': '155%',
    'Hora ordinaria en domingo o festivo': '80%',
    'Recargo nocturno': '35%'
}

def convertir_hora(hora_str):
    hora_str = hora_str.strip().lower()
    hora_str = hora_str.replace(' ', '')
    hora_str = hora_str.replace('p.m', 'pm').replace('a.m', 'am')
    # Si no tiene minutos, agregamos ":00"
    if ':' not in hora_str:
        hora_str = hora_str[:-2] + ':00' + hora_str[-2:]
    return datetime.strptime(hora_str, '%I:%M%p')

def procesar_excel(df):
    df.columns = [col.strip().upper() for col in df.columns]
    df['FECHA'] = pd.to_datetime(df['FECHA'])

    conceptos = []
    for i in range(len(df)):
        ini = convertir_hora(df['INICIAL'].iloc[i])
        fin = convertir_hora(df['FINAL'].iloc[i])
        duracion = (fin - ini).seconds / 3600
        dia_semana = df['FECHA'].iloc[i].day_name()
        es_domingo = (dia_semana == 'Sunday')

        # Caso especial: hora ordinaria en domingo o festivo
        if es_domingo:
            concepto_base = 'Hora ordinaria en domingo o festivo'
            concepto = f"{concepto_base} ({PORCENTAJES[concepto_base]})"
            conceptos.append((df['NOMBRE'].iloc[i], concepto, duracion))
            continue

        # Determinar si es diurna o nocturna
        if ini.hour >= 6 and fin.hour <= 21:
            concepto_base = 'Hora extra diurna'
            concepto = f"{concepto_base} ({PORCENTAJES[concepto_base]})"
            conceptos.append((df['NOMBRE'].iloc[i], concepto, duracion))
        elif ini.hour >= 21 or fin.hour < 6:
            concepto_base = 'Hora extra nocturna'
            concepto = f"{concepto_base} ({PORCENTAJES[concepto_base]})"
            conceptos.append((df['NOMBRE'].iloc[i], concepto, duracion))
        else:
            # Mixto: parte diurna y parte nocturna
            corte_nocturno = ini.replace(hour=21, minute=0)
            horas_diurna = 0
            horas_nocturna = 0
            if fin > corte_nocturno:
                horas_diurna = (corte_nocturno - ini).seconds / 3600
                horas_nocturna = (fin - corte_nocturno).seconds / 3600
            else:
                horas_diurna = duracion
            if horas_diurna > 0:
                concepto_base = 'Hora extra diurna'
                concepto = f"{concepto_base} ({PORCENTAJES[concepto_base]})"
                conceptos.append((df['NOMBRE'].iloc[i], concepto, horas_diurna))
            if horas_nocturna > 0:
                concepto_base = 'Hora extra nocturna'
                concepto = f"{concepto_base} ({PORCENTAJES[concepto_base]})"
                conceptos.append((df['NOMBRE'].iloc[i], concepto, horas_nocturna))

        # Recargo nocturno adicional si hay horas nocturnas
        if ini.hour >= 21 or fin.hour < 6:
            concepto_base = 'Recargo nocturno'
            concepto = f"{concepto_base} ({PORCENTAJES[concepto_base]})"
            conceptos.append((df['NOMBRE'].iloc[i], concepto, duracion))

    df_conceptos = pd.DataFrame(conceptos, columns=['NOMBRE', 'CONCEPTO', 'HORAS'])
    resumen = df_conceptos.groupby(['NOMBRE', 'CONCEPTO'])['HORAS'].sum().reset_index()
    return resumen

# Interfaz Streamlit
st.title("📝 Horas Extras Universidad Autónoma del Caribe")
st.write("Sube tu archivo Excel y genera el resumen con conceptos y porcentajes.")

archivo = st.file_uploader("Selecciona tu archivo Excel", type=["xlsx"])

if archivo:
    df = pd.read_excel(archivo, sheet_name='Hoja1', engine='openpyxl')
    resumen = procesar_excel(df)

    st.success("✅ Archivo procesado correctamente.")
    st.write("### Resumen de horas por concepto:")
    st.dataframe(resumen)

    # Convertir a Excel en memoria
    buffer = BytesIO()
    resumen.to_excel(buffer, index=False)
    buffer.seek(0)

    st.download_button(
        label="📥 Descargar resumen en Excel",
        data=buffer,
        file_name="resumen_todos_conceptos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

