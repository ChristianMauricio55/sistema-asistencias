import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta
import zipfile
import io

st.title("Sistema de Procesamiento de Asistencias")

# -------------------------------------------------
# NORMALIZAR HORA 🔥
# -------------------------------------------------

def normalizar_hora(texto):
    texto = str(texto).strip().replace(".", ":")

    if re.match(r"^\d{3,4}$", texto):  # 800 → 08:00
        texto = texto.zfill(4)
        texto = texto[:2] + ":" + texto[2:]

    return texto

# -------------------------------------------------
# FORMATEAR TIEMPO
# -------------------------------------------------

def formatear_tiempo(td):
    total_segundos = int(td.total_seconds())
    negativo = total_segundos < 0
    total_segundos = abs(total_segundos)

    horas = total_segundos // 3600
    minutos = (total_segundos % 3600) // 60
    segundos = total_segundos % 60

    tiempo = f"{horas:02}:{minutos:02}:{segundos:02}"
    return "-" + tiempo if negativo else tiempo

def texto_a_timedelta(texto):
    texto = str(texto).strip()
    if texto == "":
        return timedelta()

    negativo = texto.startswith("-")
    texto = texto.replace("-", "")

    try:
        partes = texto.split(":")
        h = int(partes[0])
        m = int(partes[1])
        s = int(partes[2]) if len(partes) > 2 else 0

        td = timedelta(hours=h, minutes=m, seconds=s)
        return -td if negativo else td
    except:
        return timedelta()

# -------------------------------------------------
# LEER EXCEL ROBUSTO 🔥
# -------------------------------------------------

def leer_excel_robusto(archivo):
    try:
        xls = pd.ExcelFile(archivo, engine="xlrd")
    except:
        xls = pd.ExcelFile(archivo)

    for sheet in xls.sheet_names:
        df = xls.parse(sheet, header=None)
        texto = df.astype(str).stack().str.cat(sep=" ").lower()

        if "asistencia" in texto or "entrada" in texto:
            return df

    return xls.parse(xls.sheet_names[0], header=None)

# -------------------------------------------------
# LIMPIAR BLOQUES SIN HORAS
# -------------------------------------------------

def limpiar_bloques_sin_horas(df):
    pattern = re.compile(r"\d{1,2}[:\.]?\d{2}")
    dias_validos = ["lunes","martes","miercoles","jueves","viernes","sabado","domingo"]

    resultado = []
    i = 0

    while i < len(df):
        fila = df.iloc[i]
        fila_texto = " ".join(str(x).lower() for x in fila)

        if any(dia in fila_texto for dia in dias_validos):
            resultado.append(list(fila))
            i += 1
            continue

        nums = sum(1 for x in fila if str(x).isdigit() and 1 <= int(x) <= 31)
        if nums >= 3:
            resultado.append(list(fila))
            i += 1
            continue

        if any(re.search(r"nombre|id|empleado|colaborador", str(x).lower()) for x in fila):
            bloque = [list(fila)]
            i += 1
            tiene_horas = False

            while i < len(df):
                fila_bloque = df.iloc[i]
                texto_bloque = " ".join(str(x).lower() for x in fila_bloque)

                if any(re.search(r"nombre|id|empleado|colaborador", str(x).lower()) for x in fila_bloque):
                    break

                if any(dia in texto_bloque for dia in dias_validos):
                    break

                if any(pattern.search(normalizar_hora(str(x))) for x in fila_bloque):
                    tiene_horas = True

                bloque.append(list(fila_bloque))
                i += 1

            if tiene_horas:
                resultado.extend(bloque)

            continue

        resultado.append(list(fila))
        i += 1

    return pd.DataFrame(resultado)

# -------------------------------------------------
# CALCULAR HORAS
# -------------------------------------------------

def calcular_horas(df, es_cedis=False):
    pattern = re.compile(r"-?\d{1,2}:\d{2}")

    resultado = []
    i = 0
    fila_dias_actual = None
    total_cols = len(df.columns) + 1

    dias_validos = ["lunes","martes","miercoles","jueves","viernes","sabado","domingo"]

    while i < len(df):
        fila = df.iloc[i]

        if any(str(x).strip().lower() in dias_validos for x in fila):
            fila_dias_actual = list(fila)
            resultado.append(list(fila) + [""] * (total_cols - len(fila)))
            i += 1
            continue

        if any(re.search(r"nombre|id|empleado|colaborador", str(x).lower()) for x in fila):
            resultado.append(list(fila) + [""] * (total_cols - len(fila)))
            i += 1
            continue

        bloque = []

        while i < len(df):
            fila_actual = df.iloc[i]

            tiene_hora = any(pattern.search(normalizar_hora(str(x))) for x in fila_actual)

            if tiene_hora:
                bloque.append(fila_actual)
                resultado.append(list(fila_actual) + [""] * (total_cols - len(fila_actual)))
                i += 1
            else:
                break

        if len(bloque) >= 2:
            horas_por_columna = [[] for _ in range(total_cols)]

            for fila_bloque in bloque:
                for col, val in enumerate(fila_bloque):
                    val = normalizar_hora(val)
                    match = pattern.search(val)

                    if match:
                        try:
                            hora = datetime.strptime(match.group(), "%H:%M")
                            horas_por_columna[col].append(hora)
                        except:
                            pass

            fila_trabajadas = [""] * total_cols
            fila_jornada = [""] * total_cols
            fila_diferencia = [""] * total_cols

            for col in range(len(df.columns)):
                horas = horas_por_columna[col]
                if len(horas) < 2:
                    continue

                es_domingo = fila_dias_actual and col < len(fila_dias_actual) and str(fila_dias_actual[col]).lower() == "domingo"
                jornada_base = timedelta(hours=9 if not es_domingo else (9 if es_cedis else 8))

                trabajadas = horas[-1] - horas[0]
                diferencia = jornada_base - trabajadas

                fila_trabajadas[col] = formatear_tiempo(trabajadas)
                fila_jornada[col] = formatear_tiempo(jornada_base)
                fila_diferencia[col] = formatear_tiempo(diferencia)

            suma = sum((texto_a_timedelta(x) for x in fila_diferencia if x), timedelta())
            fila_diferencia[-1] = formatear_tiempo(suma)

            resultado.extend([fila_trabajadas, fila_jornada, fila_diferencia])

        if i < len(df):
            resultado.append(list(df.iloc[i]) + [""] * (total_cols - len(df.iloc[i])))
            i += 1

    return pd.DataFrame(resultado)

# -------------------------------------------------
# PROCESAR EXCEL
# -------------------------------------------------

def procesar_excel(archivo):
    df = leer_excel_robusto(archivo)

    # 🔥 LIMPIEZA FUERTE
    df = df.dropna(how="all")
    df = df.fillna(method="ffill")
    df = df.fillna("")
    df = df.reset_index(drop=True)

    texto = df.astype(str).stack().str.cat(sep=" ")
    es_cedis = "CEDIS" in texto.upper()

    pattern = re.compile(r"\d{1,2}[:\.]?\d{2}")

    resultado = []

    for _, row in df.iterrows():
        columnas = []

        for celda in row:
            texto = normalizar_hora(celda)
            horas = pattern.findall(texto)
            columnas.append(horas if horas else [texto])

        max_len = max(len(c) for c in columnas)

        for i in range(max_len):
            resultado.append([c[i] if i < len(c) else "" for c in columnas])

    df_final = pd.DataFrame(resultado)

    df_final = limpiar_bloques_sin_horas(df_final)

    return calcular_horas(df_final, es_cedis)

# -------------------------------------------------
# UI
# -------------------------------------------------

archivos = st.file_uploader(
    "Subir archivos Excel",
    type=["xls","xlsx"],
    accept_multiple_files=True
)

if archivos:
    if st.button("Procesar archivos"):

        tablas = []

        for archivo in archivos:
            df_final = procesar_excel(archivo)

            separador = pd.DataFrame([[f"===== {archivo.name} ====="]])
            tablas.append(separador)
            tablas.append(df_final)
            tablas.append(pd.DataFrame([[""]]))

        df_unido = pd.concat(tablas, ignore_index=True)

        excel_buffer = io.BytesIO()
        df_unido.to_excel(excel_buffer, index=False, header=False)
        excel_buffer.seek(0)

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            zipf.writestr("reporte_unificado.xlsx", excel_buffer.read())

        zip_buffer.seek(0)

        st.success("Archivo ZIP generado correctamente")

        st.download_button(
            "Descargar ZIP con Excel",
            zip_buffer.getvalue(),
            file_name="reporte_unificado.zip"
        )