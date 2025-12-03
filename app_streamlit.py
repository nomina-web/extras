
import pandas as pd
from datetime import datetime, timedelta, date
from functools import lru_cache
import streamlit as st
from io import BytesIO

# --- Tabla de recargos según ley laboral ---
PORCENTAJES = {
    'Hora extra diurna': '25%',
    'Hora extra nocturna': '75%',
    'Hora extra diurna en domingo o festivo': '105%',
    'Hora extra nocturna en domingo o festivo': '155%',
    'Hora ordinaria en domingo o festivo': '80%',
    'Recargo nocturno festivo': '115%'
}

HORAS_JORNADA = 8  # Jornada ordinaria aplicable solo en festivos/domingo por día real

# --- Utilidades ---
def convertir_hora(hora_str: str) -> datetime:
    s = str(hora_str).strip().lower().replace(' ', '')
    s = s.replace('.', '')
    s = s.replace('p.m', 'pm').replace('a.m', 'am')
    if ':' not in s and (s.endswith('am') or s.endswith('pm')):
        s = s[:-2] + ':00' + s[-2:]
    try:
        return datetime.strptime(s, '%I:%M%p')
    except ValueError:
        if ':' not in s:
            s = f'{s}:00'
        return datetime.strptime(s, '%H:%M')

def combinar_fecha_hora(fecha, hora_dt):
    return datetime.combine(pd.to_datetime(fecha).date(), hora_dt.time())

def dividir_por_dia(fecha, ini_time_dt, fin_time_dt):
    ini = combinar_fecha_hora(fecha, ini_time_dt)
    fin = combinar_fecha_hora(fecha, fin_time_dt)
    if fin <= ini:
        fin += timedelta(days=1)
    bloques = []
    actual = ini
    while actual.date() < fin.date():
        corte = datetime.combine(actual.date(), datetime.strptime('23:59', '%H:%M').time()) + timedelta(minutes=1)
        bloques.append((actual.date(), actual, corte))
        actual = corte
    bloques.append((fin.date(), actual, fin))
    return bloques

def segmentar_por_franja(ini_dt, fin_dt):
    if fin_dt <= ini_dt:
        fin_dt += timedelta(days=1)
    cortes = []
    for base in [ini_dt.date(), (ini_dt + timedelta(days=1)).date()]:
        cortes.append(datetime.combine(base, datetime.strptime('06:00', '%H:%M').time()))
        cortes.append(datetime.combine(base, datetime.strptime('21:00', '%H:%M').time()))
    puntos = [ini_dt, fin_dt] + [c for c in cortes if ini_dt < c < fin_dt]
    puntos.sort()
    segmentos = []
    for s, e in zip(puntos[:-1], puntos[1:]):
        dur = (e - s).total_seconds() / 3600.0
        mid = s + (e - s) / 2
        tipo = 'diurna' if 6 <= mid.hour < 21 else 'nocturna'
        segmentos.append({'dur': dur, 'tipo': tipo, 'dia': s.date(), 'start': s})
    return segmentos

# --- Festivos ---
def next_monday(d: date) -> date:
    return d + timedelta(days=(7 - d.weekday()) % 7)

def easter_sunday(year: int) -> date:
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
   ()
    fest.update({
        date(year, 1, 1), date(year, 5, 1), date(year, 7, 20),
        date(year, 8, 7), date(year, 12, 8), date(year, 12, 25)
    })
    easter = easter_sunday(year)
    fest.update({easter - timedelta(days=3), easter - timedelta(days=2)})
    fest.update({
        next_monday(date(year, 1, 6)), next_monday(date(year, 3, 19)),
        next_monday(date(year, 6, 29)), next_monday(date(year, 8, 15)),
        next_monday(date(year, 10, 12)), next_monday(date(year, 11, 1)),
        next_monday(date(year, 11, 11))
    })
    fest.add(next_monday(easter + timedelta(days=43)))  # Ascensión
    fest.add(next_monday(easter + timedelta(days=60)))  # Corpus Christi
    fest.add(next_monday(easter + timedelta(days=68)))  # Sagrado Corazón
    return fest

def construir_calendario_festivos(col_fechas: pd.Series) -> set[date]:
    anos = sorted(pd.to_datetime(col_fechas).dt.year.unique().tolist())
    calendario = set()
    for y in anos:
        calendario |= festivos_colombia(y)
    return calendario

# --- Procesamiento principal ---
def procesar_excel(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [col.strip().upper() for col in df.columns]
    df['FECHA'] = pd.to_datetime(df['FECHA'])
    df['INI_DT'] = df['INICIAL'].apply(convertir_hora)
    df['FIN_DT'] = df['FINAL'].apply(convertir_hora)

    festivos_set = construir_calendario_festivos(df['FECHA'])
    conceptos = []

    def add_concepto(nombre, concepto_base, horas):
        if horas > 0:
            conceptos.append((nombre, concepto_base, horas))

    for (nombre, fecha), grupo in df.groupby(['NOMBRE', 'FECHA']):
        grupo = grupo.sort_values(by='INI_DT')
        segmentos_turno = []
        for _, row in grupo.iterrows():
            bloques = dividir_por_dia(fecha, row['INI_DT'], row['FIN_DT'])
            for _, ini_dt, fin_dt in bloques:
                segmentos_turno.extend(segmentar_por_franja(ini_dt, fin_dt))

        segmentos_turno.sort(key=lambda seg: seg['start'])

        horas_ordinarias_restantes_por_dia = {}

        for seg in segmentos_turno:
            dur = seg['dur']
            tipo = seg['tipo']
            dia_real = seg['dia']

            es_domingo = (dia_real.weekday() == 6)
            es_festivo = (dia_real in festivos_set)
            es_festivo_o_domingo = es_domingo or es_festivo

            if es_festivo_o_domingo:
                if dia_real not in horas_ordinarias_restantes_por_dia:
                    horas_ordinarias_restantes_por_dia[dia_real] = HORAS_JORNADA

                restante = horas_ordinarias_restantes_por_dia[dia_real]
                ordinaria = min(restante, dur)
                extra = max(0.0, dur - ordinaria)

                if ordinaria > 0:
                    if tipo == 'diurna':
                        add_concepto(nombre, 'Hora ordinaria en domingo o festivo', ordinaria)
                    else:
                        add_concepto(nombre, 'Recargo nocturno festivo', ordinaria)

                if extra > 0:
                    if tipo == 'diurna':
                        add_concepto(nombre, 'Hora extra diurna en domingo o festivo', extra)
                    else:
                        add_concepto(nombre, 'Hora extra nocturna en domingo o festivo', extra)

                horas_ordinarias_restantes_por_dia[dia_real] = max(0.0, restante - ordinaria)

            else:
                if tipo == 'diurna':
                    add_concepto(nombre, 'Hora extra diurna', dur)
                else:
                    add_concepto(nombre, 'Hora extra nocturna', dur)

    df_conceptos = pd.DataFrame(conceptos, columns=['NOMBRE', 'CONCEPTO_BASE', 'HORAS'])
    if df_conceptos.empty:
        return pd.DataFrame(columns=['NOMBRE', 'CONCEPTO', 'HORAS'])

    resumen = df_conceptos.groupby(['NOMBRE', 'CONCEPTO_BASE'], as_index=False)['HORAS'].sum()
    resumen['CONCEPTO'] = resumen['CONCEPTO_BASE'].apply(lambda c: f"{c} ({PORCENTAJES[c]})")
    return resumen[['NOMBRE', 'CONCEPTO', 'HORAS']]

# --- Interfaz Streamlit ---
st.title("📝 Horas Extras Universidad Autónoma del Caribe")
st.write("Sube tu archivo Excel y genera el resumen con conceptos y porcentajes.")

archivo = st.file_uploader("Selecciona tu archivo Excel", type=["xlsx"])

if archivo:
    df = pd.read_excel(archivo, sheet_name='Hoja1', engine='openpyxl')
    resumen = procesar_excel(df)

    st.success("✅ Archivo procesado correctamente.")
    st.write("### Resumen de horas por concepto:")
    st.dataframe(resumen)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        resumen.to_excel(writer, index=False, sheet_name='Resumen')
    buffer.seek(0)

    st.download_button(
        label="📥 Descargar resumen en Excel",
        data=buffer,
        file_name="resumen_todos_conceptos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
