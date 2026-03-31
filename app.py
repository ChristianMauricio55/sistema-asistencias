import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta
import zipfile
import os
import io

st.title("Sistema de Procesamiento de Asistencias")

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

    if negativo:
        tiempo = "-" + tiempo

    return tiempo


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
# BORRAR DATOS SIN HORAS (🔥)
# -------------------------------------------------

def limpiar_bloques_sin_horas(df):

    pattern = re.compile(r"\d{1,2}:\d{2}")
    dias_validos = ["lunes","martes","miercoles","jueves","viernes","sabado","domingo"]

    resultado = []
    i = 0

    while i < len(df):

        fila = df.iloc[i]
        fila_texto = " ".join(str(x).lower() for x in fila)

        # 🔹 mantener filas de días
        if any(dia in fila_texto for dia in dias_validos):
            resultado.append(list(fila))
            i += 1
            continue

        # 🔹 mantener filas de fechas
        nums = 0
        for x in fila:
            try:
                num = int(float(x))
                if 1 <= num <= 31:
                    nums += 1
            except:
                pass

        if nums >= 3:
            resultado.append(list(fila))
            i += 1
            continue

        # 🔥 detectar inicio de bloque (Nombre / ID)
        if any("nombre" in str(x).lower() or "id" in str(x).lower() for x in fila):

            bloque = [list(fila)]
            i += 1

            tiene_horas = False

            # recorrer bloque completo
            while i < len(df):

                fila_bloque = df.iloc[i]
                texto_bloque = " ".join(str(x).lower() for x in fila_bloque)

                # si encontramos otro encabezado → termina bloque
                if any("nombre" in str(x).lower() or "id" in str(x).lower() for x in fila_bloque):
                    break

                # si encontramos días → termina bloque
                if any(dia in texto_bloque for dia in dias_validos):
                    break

                # verificar si hay horas
                if any(pattern.search(str(x)) for x in fila_bloque):
                    tiene_horas = True

                bloque.append(list(fila_bloque))
                i += 1

            # 🔥 SOLO guardar si tiene horas
            if tiene_horas:
                resultado.extend(bloque)

            continue

        # 🔹 otras filas (por seguridad)
        resultado.append(list(fila))
        i += 1

    return pd.DataFrame(resultado)


# -------------------------------------------------
# CALCULAR HORAS (CON CEDIS 🔥)
# -------------------------------------------------

def calcular_horas(df, es_cedis=False):

    pattern = re.compile(r"-?\d{1,2}:\d{2}")

    resultado = []
    i = 0
    fila_dias_actual = None

    total_cols = len(df.columns) + 1  # 🔥 columna extra fija

    dias_validos = ["lunes","martes","miercoles","jueves","viernes","sabado","domingo"]

    while i < len(df):

        fila = df.iloc[i]

        # 🔥 detectar fila de días
        if any(str(x).strip().lower() in dias_validos for x in fila):

            conteo = sum(str(x).strip().lower() in dias_validos for x in fila)

            if conteo >= 5:
                fila_dias_actual = list(fila)

                fila_lista = list(fila)
                fila_lista += [""] * (total_cols - len(fila_lista))
                resultado.append(fila_lista)

                i += 1
                continue

        # encabezado
        if any("Nombre" in str(x) or "ID" in str(x) for x in fila):

            fila_lista = list(fila)
            fila_lista += [""] * (total_cols - len(fila_lista))
            resultado.append(fila_lista)

            i += 1
            continue

        bloque = []

        while i < len(df):

            fila_actual = df.iloc[i]

            tiene_hora = any(pattern.search(str(x)) for x in fila_actual)

            if tiene_hora:
                bloque.append(fila_actual)

                fila_lista = list(fila_actual)
                fila_lista += [""] * (total_cols - len(fila_lista))
                resultado.append(fila_lista)

                i += 1
            else:
                break

        if len(bloque) >= 2:

            horas_por_columna = [[] for _ in range(total_cols)]

            for fila_bloque in bloque:
                for col, val in enumerate(fila_bloque):

                    val = str(val).strip()
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

                # 🔥 SOLO procesar si hay datos reales
                if len(horas) < 2:
                    continue

                # detectar domingo
                if fila_dias_actual and col < len(fila_dias_actual):
                    es_domingo = str(fila_dias_actual[col]).strip().lower() == "domingo"
                else:
                    es_domingo = False

                # regla CEDIS
                if es_domingo:
                    jornada_base = timedelta(hours=9 if es_cedis else 8)
                else:
                    jornada_base = timedelta(hours=9)

                trabajadas = horas[-1] - horas[0]
                diferencia = jornada_base - trabajadas

                fila_trabajadas[col] = formatear_tiempo(trabajadas)
                fila_jornada[col] = formatear_tiempo(jornada_base)
                fila_diferencia[col] = formatear_tiempo(diferencia)

            # 🔥 suma diferencias (solo donde hay datos reales)
            suma_diferencias = timedelta()

            for val in fila_diferencia:
                if val != "":
                    suma_diferencias += texto_a_timedelta(val)

            fila_diferencia[-1] = formatear_tiempo(suma_diferencias)

            resultado.append(fila_trabajadas)
            resultado.append(fila_jornada)
            resultado.append(fila_diferencia)

        if i < len(df):
            fila_lista = list(df.iloc[i])
            fila_lista += [""] * (total_cols - len(fila_lista))
            resultado.append(fila_lista)
            i += 1

    return pd.DataFrame(resultado)


# -------------------------------------------------
# FUNCION PRINCIPAL
# -------------------------------------------------

def procesar_excel(archivo):

    try:
        df = pd.read_excel(archivo, sheet_name="Reporte de Asistencia", header=None)
    except:
        df = pd.read_excel(archivo, sheet_name=0, header=None)

    df = df.fillna("")

    texto = df.astype(str).stack().str.cat(sep=" ")
    es_cedis = "CEDIS" in texto.upper()

    # 🔥 GENERAR DÍAS DESDE FECHAS (VERSIÓN ROBUSTA)
    periodo = re.search(r"(\d{4})-(\d{2})-\d{2}", texto)

    if periodo:
        anio = int(periodo.group(1))
        mes = int(periodo.group(2))

        dias_es_num = {
            0: "Lunes",
            1: "Martes",
            2: "Miercoles",
            3: "Jueves",
            4: "Viernes",
            5: "Sabado",
            6: "Domingo"
        }

        for i, row in df.iterrows():

            valores_numericos = []

            for x in row:
                try:
                    num = int(float(x))
                    if 1 <= num <= 31:
                        valores_numericos.append(num)
                except:
                    pass

            # 🔥 detectar fila de fechas REAL
            if len(valores_numericos) >= 5:

                fila_dias = [""] * len(df.columns)

                for col, val in enumerate(row):
                    try:
                        num = int(float(val))
                        if 1 <= num <= 31:
                            fecha = datetime(anio, mes, num)
                            fila_dias[col] = dias_es_num[fecha.weekday()]
                    except:
                        pass

                df = pd.concat([
                    df.iloc[:i+1],
                    pd.DataFrame([fila_dias]),
                    df.iloc[i+1:]
                ]).reset_index(drop=True)

                break

    pattern = re.compile(r"\d{1,2}:\d{2}")
    hay_regex = df.astype(str).apply(lambda col: col.str.contains(pattern).any()).any()

    resultado = []

    for _, row in df.iterrows():

        texto_fila = " ".join([str(x) for x in row])

        # 🔥 NO romper fila de fechas (MEJORADO)
        nums_detectados = 0
        for x in row:
            try:
                num = int(float(x))
                if 1 <= num <= 31:
                    nums_detectados += 1
            except:
                pass

        if nums_detectados >= 3:
            resultado.append(list(row))
            continue

        # 🔥 NO romper fila de días
        if any(str(x).strip().lower() in ["lunes","martes","miercoles","jueves","viernes","sabado","domingo"] for x in row):
            resultado.append(list(row))
            continue

        columnas = []

        for celda in row:

            texto = str(int(celda)) if isinstance(celda, float) and celda.is_integer() else str(celda)

            if hay_regex:
                horas = pattern.findall(texto)
                columnas.append(horas if horas else [texto])
            else:
                partes = [x.strip() for x in texto.split("\n") if x.strip()]
                columnas.append(partes if partes else [""])

        max_len = max(len(c) for c in columnas)

        for i in range(max_len):
            resultado.append([c[i] if i < len(c) else "" for c in columnas])

    df_final = pd.DataFrame(resultado)

    # 🔥 LIMPIAR antes de calcular
    df_final = limpiar_bloques_sin_horas(df_final)

    return calcular_horas(df_final, es_cedis)


# -------------------------------------------------
# SUBIR ARCHIVOS
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

        # 🔥 crear excel en memoria
        excel_buffer = io.BytesIO()
        df_unido.to_excel(excel_buffer, index=False, header=False)
        excel_buffer.seek(0)

        # 🔥 ahora crear ZIP
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
