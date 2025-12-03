
import re
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

HORAS_JORNADA = 8  # Jornada ordinaria aplicable solo en festivo/domingo por día real

# --- Utilidades ---
def convertir_hora(hora_str: str) -> datetime:
    """
    Convierte una cadena de hora en objeto datetime, manejando:
    - Formatos con AM/PM (con o sin puntos, espacios)
    - Formatos 12h y 24h
    - Casos sin minutos (ej: 8AM -> 08:00AM)
    - Corrige casos como '16AM' o '16 PM' (lo interpreta en 24h)
    Lanza ValueError con mensaje claro si no puede convertir.
    """
    if pd.isna(hora_str):
        raise ValueError("Valor de hora vacío o nulo")

    s = str(hora_str).strip().lower()
    # Eliminar espacios y puntos
    s = re.sub(r'[ .]', '', s)
    # Homologar sufijos am/pm
    s = s.replace('a.m', 'am').replace('p.m', 'pm')

    # Solo número + am/pm sin minutos -> agregar :00 (p.e. '8am' -> '8:00am')
    if re.match(r'^\d{1,2}(am|pm)$', s):
        s = s[:-2] + ':00' + s[-2:]

    # Casos como '16am' o '16pm' (inválidos en 12h): tratar como 24h
    if re.match(r'^\d{2}(am|pm)$', s):
        num = int(s[:-2])
        if num > 12:
            s = f'{num}:00'

    # Solo número (sin minutos ni am/pm) -> agregar :00
    if re.match(r'^\d{1,2}$', s):
        s = s + ':00'

    # Intentos de parseo en orden
    for fmt in ['%I:%M%p', '%H:%M']:
        try:
            return datetime.strptime(s.upper(), fmt)
        except ValueError:
            continue

    raise ValueError(f"Formato de hora inválido: '{hora_str}'")

def combinar_fecha_hora(fecha, hora_dt):
    return datetime.combine(pd.to_datetime(fecha).date(), hora_dt.time())

def dividir_por_dia(fecha, ini_time_dt, fin_time_dt):
    ini = combinar_fecha_hora(fecha, ini_time_dt)
    fin = combinar_fecha_hora(fecha, fin_time_dt)
    # Si fin ≤ ini, asumimos que cruza la medianoche (siguiente día)
    if fin <= ini:
        fin += timedelta(days=1)

    bloques = []
    actual = ini
    while actual.date() < fin.date():
        # Corte al fin del día
        corte = datetime.combine(actual.date(), datetime.strptime('23:59', '%H:%M').time()) + timedelta(minutes=1)
        bloques.append((actual.date(), actual, corte))
        actual = corte
    # Último bloque hasta fin
    bloques.append((fin.date(), actual, fin))
    return bloques

def segmentar_por_franja(ini_dt, fin_dt):
    # Garantiza que fin > ini (si cruza medianoche, se corrige)
    if fin_dt <= ini_dt:
        fin_dt += timedelta(days=1)

    # Cortes a las 06:00 y 21:00 en el día base y el siguiente
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

# --- Festivos (Colombia) ---
def next_monday(d: date) -> date:
    return d + timedelta(days=(7 - d.weekday()) % 7)

def easter_sunday(year: int) -> date:
    # Computus (algoritmo clásico)
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

@lru_cache(maxsize=None)
def festivos_colombia(year: int) -> set[date]:
    fest = set()
    # Fijos
    fest.update({
        date(year, 1, 1),   # Año Nuevo
        date(year, 5, 1),   # Día del Trabajo
        date(year, 7, 20),  # Independencia
        date(year, 8, 7),   # Batalla de Boyacá
        date(year, 12, 8),  # Inmaculada Concepción
        date(year, 12, 25), # Navidad
    })
    easter = easter_sunday(year)
    # Jueves/Viernes Santo
    fest.update({easter - timedelta(days=3), easter - timedelta(days=2)})
    # Movibles (Ley Emiliani)
    fest.update({
        next_monday(date(year, 1, 6)),   # Epifanía
        next_monday(date(year, 3, 19)),  # San José
        next_monday(date(year, 6, 29)),  # San Pedro y San Pablo
        next_monday(date(year, 8, 15)),  # Asunción
        next_monday(date(year, 10, 12)), # Día de la Raza
        next_monday(date(year, 11, 1)),  # Todos los Santos
        next_monday(date(year, 11, 11)), # Independencia de Cartagena
    })
    # Móviles alrededor de Pascua
    fest.add(next_monday(easter + timedelta(days=43))) # Ascensión
    fest.add(next_monday(easter + timedelta(days=60))) # Corpus Christi
    fest.add(next_monday(easter + timedelta(days=68))) # Sagrado Corazón
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

    # Parseo de fechas día/mes/año (según tu hoja) y detección de inválidas
    df['FECHA'] = pd.to_datetime(df['FECHA'], dayfirst=True, errors='coerce')
    if df['FECHA'].isna().any():
        filas_invalidas = df[df['FECHA'].isna()]
        raise ValueError(
            f"Se encontraron fechas inválidas en {len(filas_invalidas)} fila(s). "
            "Verifica el formato (dd/mm/aaaa)."
        )

    # Conversión de horas
    df['INI_DT'] = df['INICIAL'].apply(convertir_hora)
    df['FIN_DT'] = df['FINAL'].apply(convertir_hora)

    festivos_set = construir_calendario_festivos(df['FECHA'])

    conceptos = []
    def add_concepto(nombre, concepto_base, horas):
        if horas > 0:
            conceptos.append((nombre, concepto_base, horas))

    # Agrupar por persona y fecha
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

# --- Helpers de validación para la interfaz ---
def encontrar_invalidos(serie: pd.Series, etiqueta_col: str) -> pd.DataFrame:
    """
    Devuelve un DataFrame con las filas cuyo valor no se puede convertir a hora.
    Incluye el índice y el valor original.
    """
    errores = []
    for idx, val in serie.items():
        try:
            convertir_hora(val)
        except Exception as e:
            errores.append({'FILA_DF': idx, 'COLUMNA': etiqueta_col, 'VALOR': val, 'ERROR': str(e)})
    return pd.DataFrame(errores)

# --- Interfaz Streamlit ---
st.title("📝 Horas Extras Universidad Autónoma del Caribe")
st.write("Sube tu archivo Excel y genera el resumen con conceptos y porcentajes.")

archivo = st.file_uploader("Selecciona tu archivo Excel", type=["xlsx"])

if archivo:
    try:
        # Leer hoja 'Hoja1'
        df = pd.read_excel(archivo, sheet_name='Hoja1', engine='openpyxl')
        df.columns = [c.strip().upper() for c in df.columns]

        # Validación de columnas requeridas
        requeridas = {'FECHA', 'NOMBRE', 'INICIAL', 'FINAL'}
        faltantes = requeridas.difference(set(df.columns))
        if faltantes:
            st.error(f"Faltan columnas requeridas en el Excel: {', '.join(sorted(faltantes))}")
            st.stop()

        # Validar fechas antes de procesar (sin romper la app)
        fechas = pd.to_datetime(df['FECHA'], dayfirst=True, errors='coerce')
        invalidas_fechas = df[fechas.isna()]
        if not invalidas_fechas.empty:
            st.error(f"Hay {len(invalidas_fechas)} fila(s) con FECHA inválida. Usa formato dd/mm/aaaa.")
            st.dataframe(invalidas_fechas)
            st.stop()

        # Validar horas INICIAL/FINAL antes de aplicar
        inv_ini = encontrar_invalidos(df['INICIAL'], 'INICIAL')
        inv_fin = encontrar_invalidos(df['FINAL'], 'FINAL')
        inv_total = pd.concat([inv_ini, inv_fin], ignore_index=True)

        if not inv_total.empty:
            st.error("Se encontraron horas inválidas. Corrige estos valores y vuelve a subir el archivo:")
            st.dataframe(inv_total)
            st.stop()

        # Procesar normalmente (ya validado)
        resumen = procesar_excel(df)

        st.success("✅ Archivo procesado correctamente.")
        st.write("### Resumen de horas por concepto:")
        st.dataframe(resumen)

        # Descarga en Excel
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

    except Exception as e:
        # Mostrar el error con más claridad en la app
        st.error(f"❌ Ocurrió un error al procesar el archivo: {e}")
