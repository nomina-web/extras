
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

HORAS_JORNADA = 8  # Umbral para domingos/festivos

# --- Utilidades de hora ---
def convertir_hora(hora_str: str) -> datetime:
    """
    Convierte cadenas tipo '08am', '6:30 pm', '10 p.m' a datetime (solo hora).
    """
    s = hora_str.strip().lower().replace(' ', '')
    s = s.replace('p.m', 'pm').replace('a.m', 'am')
    if ':' not in s:
        s = s[:-2] + ':00' + s[-2:]  # '08am' -> '08:00am'
    return datetime.strptime(s, '%I:%M%p')

def combinar_fecha_hora(fecha, hora_dt):
    return datetime.combine(pd.to_datetime(fecha).date(), hora_dt.time())

def segmentar_por_franja(fecha, ini_time_dt, fin_time_dt):
    """
    Divide un intervalo en segmentos diurnos (06:00–21:00) y nocturnos (21:00–06:00).
    Soporta cruces por 21:00, 06:00 y medianoche.
    """
    ini = combinar_fecha_hora(fecha, ini_time_dt)
    fin = combinar_fecha_hora(fecha, fin_time_dt)
    if fin <= ini:
        # Si fin es "menor", asumimos que cruza medianoche
        fin = fin + timedelta(days=1)

    # Construye posibles cortes relevantes (06:00 y 21:00 para el día y el siguiente)
    cortes = []
    for base in [ini.date(), (ini + timedelta(days=1)).date()]:
        cortes.append(datetime.combine(base, datetime.strptime('06:00', '%H:%M').time()))
        cortes.append(datetime.combine(base, datetime.strptime('21:00', '%H:%M').time()))

    puntos = [ini, fin] + [c for c in cortes if ini < c < fin]
    puntos.sort()

    segmentos = []
    for s, e in zip(puntos[:-1], puntos[1:]):
        dur = (e - s).total_seconds() / 3600.0
        mid = s + (e - s) / 2
        tipo = 'diurna' if 6 <= mid.hour < 21 else 'nocturna'
        segmentos.append((dur, tipo))
    return segmentos

# --- Computus: Pascua (Meeus/Jones/Butcher) ---
def easter_sunday(year: int) -> date:
    """
    Calcula el Domingo de Pascua (calendario gregoriano) para 'year'.
    Algoritmo Meeus/Jones/Butcher.
    """
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
    """Devuelve el lunes de observancia (incluye el mismo día si ya es lunes)."""
    return d + timedelta(days=(0 - d.weekday()) % 7)

@lru_cache(maxsize=None)
def festivos_colombia(year: int) -> set[date]:
    """
    Set de festivos nacionales observados en Colombia para 'year'.
    Incluye:
      - Inamovibles (Año Nuevo, Trabajo, Independencia, Boyacá, Inmaculada, Navidad)
      - Semana Santa (Jueves/Viernes Santo)
      - Trasladables por Ley Emiliani (se observan lunes)
      - Móviles ligados a Pascua (observados en lunes)
    """
    fest = set()

    # Inamovibles
    fest.update({
        date(year, 1, 1),   # Año Nuevo
        date(year, 5, 1),   # Día del Trabajo
        date(year, 7, 20),  # Independencia
        date(year, 8, 7),   # Batalla de Boyacá
        date(year, 12, 8),  # Inmaculada Concepción
        date(year, 12, 25), # Navidad
    })

    # Semana Santa (no se trasladan)
    easter = easter_sunday(year)
    jueves_santo = easter - timedelta(days=3)
    viernes_santo = easter - timedelta(days=2)
    fest.update({jueves_santo, viernes_santo})

    # Trasladables por Ley Emiliani: observados en lunes
    fest.update({
        next_monday(date(year, 1, 6)),   # Reyes
        next_monday(date(year, 3, 19)),  # San José
        next_monday(date(year, 6, 29)),  # San Pedro y San Pablo
        next_monday(date(year, 8, 15)),  # Asunción
        next_monday(date(year, 10, 12)), # Diversidad (Colón)
        next_monday(date(year, 11, 1)),  # Todos los Santos
        next_monday(date(year, 11, 11)), # Independencia de Cartagena
    })

    # Festivos móviles ligados a Pascua (observados en lunes)
    fest.add(easter + timedelta(days=43))  # Ascensión (Easter+39 -> lunes observado a +43)
    fest.add(easter + timedelta(days=64))  # Corpus Christi (Easter+60 -> lunes observado a +64)
    fest.add(easter + timedelta(days=71))  # Sagrado Corazón (Easter+68 -> lunes observado a +71)

    return fest

def construir_calendario_festivos(col_fechas: pd.Series) -> set[date]:
    """
    Construye el calendario de festivos para todos los años presentes en el DF.
    """
    anos = sorted(pd.to_datetime(col_fechas).dt.year.unique().tolist())
    calendario = set()
    for y in anos:
        calendario |= festivos_colombia(y)
    return calendario

# --- Procesamiento principal ---
def procesar_excel(df: pd.DataFrame) -> pd.DataFrame:
    """
    Procesa el DataFrame con columnas:
      - NOMBRE
      - FECHA
      - INICIAL
      - FINAL
    Aplica reglas:
      Día normal: todas las horas son extra (25% diurna, 75% nocturna)
      Domingo/festivo: primeras 8 horas ordinarias dominicales (80%), resto extras dominicales (105%/155%)
    """
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
        es_domingo = (fecha.weekday() == 6)  # Sunday=6
        es_festivo = (fecha.date() in festivos_set)
        es_festivo_o_domingo = es_domingo or es_festivo

        # Umbral de 8 horas SOLO aplica en domingo/festivo
        horas_restantes_ordinarias = HORAS_JORNADA if es_festivo_o_domingo else 0

        grupo = grupo.sort_values(by='INI_DT')

        for _, row in grupo.iterrows():
            segmentos = segmentar_por_franja(fecha, row['INI_DT'], row['FIN_DT'])
            for dur, tipo in segmentos:
                if es_festivo_o_domingo:
                    ordinaria = min(horas_restantes_ordinarias, dur)
                    extra = max(0.0, dur - ordinaria)

                    if tipo == 'diurna':
                        if ordinaria > 0:
                            add_concepto(nombre, 'Hora ordinaria en domingo o festivo', ordinaria)
                        if extra > 0:
                            add_concepto(nombre, 'Hora extra diurna en domingo o festivo', extra)
                    else:  # nocturna
                        if ordinaria > 0:
                            add_concepto(nombre, 'Hora ordinaria en domingo o festivo', ordinaria)
                            add_concepto(nombre, 'Recargo nocturno', ordinaria)
                        if extra > 0:
                            add_concepto(nombre, 'Hora extra nocturna en domingo o festivo', extra)

                    horas_restantes_ordinarias -= ordinaria
                    if horas_restantes_ordinarias < 0:
                        horas_restantes_ordinarias = 0.0
                else:
                    # Día normal: todas las horas son extras
                    if tipo == 'diurna':
                        add_concepto(nombre, 'Hora extra diurna', dur)
                    else:
                        add_concepto(nombre, 'Hora extra nocturna', dur)

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
