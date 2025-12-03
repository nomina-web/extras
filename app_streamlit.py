
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
    'Recargo nocturno': '35%'
}

HORAS_JORNADA = 8  # Jornada diaria estándar

# --- Funciones auxiliares ---
def convertir_hora(hora_str):
    """Convierte '08am', '6:30 pm', '10 p.m' a datetime (solo hora)."""
    hora_str = hora_str.strip().lower().replace(' ', '')
    hora_str = hora_str.replace('p.m', 'pm').replace('a.m', 'am')
    if ':' not in hora_str:
        hora_str = hora_str[:-2] + ':00' + hora_str[-2:]
    return datetime.strptime(hora_str, '%I:%M%p')

def combinar_fecha_hora(fecha, hora_dt):
    return datetime.combine(pd.to_datetime(fecha).date(), hora_dt.time())

def segmentar_por_franja(fecha, ini_time_dt, fin_time_dt):
    """
    Divide un intervalo en segmentos diurnos (06:00–21:00) y nocturnos (21:00–06:00).
    """
    ini = combinar_fecha_hora(fecha, ini_time_dt)
    fin = combinar_fecha_hora(fecha, fin_time_dt)
    if fin <= ini:
        fin = fin + timedelta(days=1)

    cortes = []
    for base in [ini.date(), (ini + timedelta(days=1)).date()]:
        dia_06 = datetime.combine(base, datetime.strptime('06:00', '%H:%M').time())
        dia_21 = datetime.combine(base, datetime.strptime('21:00', '%H:%M').time())
        cortes.extend([dia_06, dia_21])

    puntos = [ini, fin] + [c for c in cortes if ini < c < fin]
    puntos = sorted(puntos)

    segmentos = []
    for s, e in zip(puntos[:-1], puntos[1:]):
        dur = (e - s).total_seconds() / 3600.0
        mid = s + (e - s) / 2
        tipo = 'diurna' if 6 <= mid.hour < 21 else 'nocturna'
        segmentos.append((dur, tipo))
    return segmentos

# --- Computus: Pascua (Meeus/Jones/Butcher) ---
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
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(year, month, day)

def next_monday(d: date) -> date:
    return d + timedelta(days=(0 - d.weekday()) % 7)

@lru_cache(maxsize=None)
def festivos_colombia(year: int) -> set[date]:
    fest = set()
    # Inamovibles
    fest.update({
        date(year, 1, 1), date(year, 5, 1), date(year, 7, 20),
        date(year, 8, 7), date(year, 12, 8), date(year, 12, 25)
    })
    # Pascua y Semana Santa
    easter = easter_sunday(year)
    fest.update({easter - timedelta(days=3), easter - timedelta(days=2)})
    # Trasladables (Ley Emiliani)
    fest.update({
        next_monday(date(year, 1, 6)), next_monday(date(year, 3, 19)),
        next_monday(date(year, 6, 29)), next_monday(date(year, 8, 15)),
        next_monday(date(year, 10, 12)), next_monday(date(year, 11, 1)),
        next_monday(date(year, 11, 11))
    })
    # Móviles ligados a Pascua
    fest.add(easter + timedelta(days=43))  # Ascensión
    fest.add(easter + timedelta(days=64))  # Corpus Christi
    fest.add(easter + timedelta(days=71))  # Sagrado Corazón
    return fest

def construir_calendario_festivos(col_fechas: pd.Series) -> set[date]:
    anos = sorted(pd.to_datetime(col_fechas).dt.year.unique().tolist())
    calendario = set()
    for y in anos:
        calendario |= festivos_colombia(y)
    return calendario

# --- Procesamiento principal ---
def procesar_excel(df):
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
        es_domingo = (fecha.weekday() == 6)
        es_festivo = (fecha.date() in festivos_set)
        es_festivo_o_domingo = es_domingo or es_festivo

        horas_restantes_ordinarias = HORAS_JORNADA
        grupo = grupo.sort_values(by='INI_DT')

        for _, row in grupo.iterrows():
            segmentos = segmentar_por_franja(fecha, row['INI_DT'], row['FIN_DT'])
            for dur, tipo in segmentos:
                ordinaria = min(horas_restantes_ordinarias, dur)
                extra = max(0.0, dur - ordinaria)

                if es_festivo_o_domingo:
                    if tipo == 'diurna':
                        if ordinaria > 0:
                            add_concepto(nombre, 'Hora ordinaria en domingo o festivo', ordinaria)
                        if extra > 0:
                            add_concepto(nombre, 'Hora extra diurna en domingo o festivo', extra)
                    else:
                        if ordinaria > 0:
                            add_concepto(nombre, 'Hora ordinaria en domingo o festivo', ordinaria)
                            add_concepto(nombre, 'Recargo nocturno', ordinaria)
                        if extra > 0:
                            add_concepto(nombre, 'Hora extra nocturna en domingo o festivo', extra)
                else:
                    if tipo == 'diurna':
                        if extra > 0:
                            add_concepto(nombre, 'Hora extra diurna', extra)
                    else:
                        if ordinaria > 0:
                            add_concepto(nombre, 'Recargo nocturno', ordinaria)
                        if extra > 0:
                            add_concepto(nombre, 'Hora extra nocturna', extra)

                horas_restantes_ordinarias -= ordinaria
                if horas_restantes_ordinarias < 0:
                    horas_restantes_ordinarias = 0.0

    df_conceptos = pd.DataFrame(conceptos, columns=['NOMBRE', 'CONCEPTO_BASE', 'HORAS'])
    resumen = df_conceptos.groupby(['NOMBRE', 'CONCEPTO_BASE'], as_index=False)['HORAS'].sum()
    resumen['CONCEPTO'] = resumen['CONCEPTO_BASE'].apply(lambda c: f"{c} ({PORCENTAJES[c]})")
    resumen = resumen[['NOMBRE', 'CONCEPTO', 'HORAS']]
    return resumen

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
    resumen.to_excel(buffer, index=False, engine='openpyxl')
    buffer.seek(0)

    st.download_button(
        label="📥 Descargar resumen en Excel",
        data=buffer,
        file_name="resumen_todos_conceptos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
