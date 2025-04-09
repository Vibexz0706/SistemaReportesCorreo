import os
import glob
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from fpdf import FPDF
import win32com.client as win32
from datetime import datetime
import numpy as np
import pdfplumber
import re
import unicodedata



# ======================== CONFIGURACIÓN ========================
ruta_carpeta = r"C:\Users\jabar\OneDrive\Escritorio\PruebaFiltro"
df_filtrado = None  # Variable global para almacenar los datos filtrados

# ======================== FUNCIONES ========================
def obtener_ultimo_excel():
    nombre_archivo = "base.xlsx"  # <- Cambialo por el nombre real del archivo
    ruta_archivo = os.path.join(ruta_carpeta, nombre_archivo)
    return ruta_archivo if os.path.exists(ruta_archivo) else None
   # archivos = glob.glob(os.path.join(ruta_carpeta, "*.xlsx"))
   #return max(archivos, key=os.path.getctime) if archivos else None
def procesar_segundo_excel():
    archivo_adicional = os.path.join(ruta_carpeta, "ExportacaoAE_DNT-Ciclo 3.xlsx")  # Cambia por el nombre real

    if not os.path.exists(archivo_adicional):
        messagebox.showerror("Error", "No se encontró el segundo archivo Excel.")
        return

    df_extra = pd.read_excel(archivo_adicional, usecols=["SECTOR", "PERÍODO"])

    # Limpieza de columnas
    df_extra["SECTOR"] = df_extra["SECTOR"].astype(str).str.strip()
    df_extra["PERÍODO"] = df_extra["PERÍODO"].astype(str)

    # Mapeo de SECTOR a nombres
    mapeo_nombres = {
        "50602": "ANA KAREN MORA",
        "50603": "RICARDO  ZUNIGA",
        "50604": "KAROLINA MONTERO",
        "50605": "KEREN CARVAJAL",
        "50606": "JOHAN ALBERTO ARCE NÚÑEZ"
        # Agrega más según lo que necesites
    }
    df_extra["NOMBRE_SECTOR"] = df_extra["SECTOR"].map(mapeo_nombres)

    # Extraer número de horas del texto como "8 horas"
    df_extra["HORAS"] = df_extra["PERÍODO"].apply(lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 0)

    # Suma total de horas por nombre
    resumen_horas = df_extra.groupby("NOMBRE_SECTOR")["HORAS"].sum().reset_index()

    print("Resumen de horas por nombre:")
    print(resumen_horas)

    # Si querés también la suma total de todo:
    total_horas = df_extra["HORAS"].sum()
    print(f"Total de horas (todos los sectores): {total_horas}")

    return resumen_horas, total_horas

def obtener_excel_especifico(nombre_archivo):
    archivo_path = os.path.join(ruta_carpeta, nombre_archivo)
    return archivo_path if os.path.exists(archivo_path) else None

def aplicar_filtros():
    CICLOS = {
        "CICLO 1": (datetime(2025, 1, 1), datetime(2025, 2, 10)),
        "CICLO 2": (datetime(2025, 2, 11), datetime(2025, 3, 13)),
        "CICLO 3": (datetime(2025, 3, 14), datetime(2025, 4, 12)),
        "CICLO 4": (datetime(2025, 4, 13), datetime(2025, 5, 15)),
    }

    def obtener_ciclo_actual(fecha_actual):
        for ciclo, (inicio, fin) in CICLOS.items():
            if inicio <= fecha_actual <= fin:
                return ciclo, inicio, fin
        return None, None, None

    fecha_actual = datetime.now()
    ciclo_actual, fecha_ciclo_inicio, fecha_ciclo_fin = obtener_ciclo_actual(fecha_actual)
    if not ciclo_actual:
        messagebox.showerror("Error", "No se pudo determinar el ciclo actual basado en la fecha.")
        return

    print(f"Ciclo actual: {ciclo_actual} ({fecha_ciclo_inicio.date()} a {fecha_ciclo_fin.date()})")
    global df_filtrado
    archivo_excel = obtener_ultimo_excel()
    if not archivo_excel:
        messagebox.showerror("Error", "No se encontró ningún archivo Excel en la carpeta.")
        return

    df = pd.read_excel(archivo_excel)
    columnas_necesarias = {"SECTOR", "NOMBRE DEL SECTOR", "FECHA DE VISITA", "CLASIFICACIÓN", "TIPO CLIENTE"}
    if not columnas_necesarias.issubset(df.columns):
        messagebox.showerror("Error", f"El archivo Excel no tiene las columnas necesarias: {columnas_necesarias}")
        return

    df["FECHA DE VISITA"] = pd.to_datetime(df["FECHA DE VISITA"], errors="coerce", dayfirst=True)

    fecha_inicio = entrada_fecha_inicio.get().strip()
    fecha_fin = entrada_fecha_fin.get().strip()
    pais_seleccionado = combo_pais.get()

    if fecha_inicio:
        df = df[df["FECHA DE VISITA"] >= pd.to_datetime(fecha_inicio, errors="coerce", dayfirst=True)]
    if fecha_fin:
        df = df[df["FECHA DE VISITA"] <= pd.to_datetime(fecha_fin, errors="coerce", dayfirst=True)]

    if pais_seleccionado == "Costa Rica":
        df = df[df["SECTOR"].astype(str).str.startswith("506")]
    elif pais_seleccionado == "Panamá":
        df = df[df["SECTOR"].astype(str).str.startswith("507")]
    elif pais_seleccionado == "Nicaragua":
        df = df[df["SECTOR"].astype(str).str.startswith("505")]
    elif pais_seleccionado == "Salvador":
        df = df[df["SECTOR"].astype(str).str.startswith("503")]
    elif pais_seleccionado == "Honduras":
        df = df[df["SECTOR"].astype(str).str.startswith("504")]
    elif pais_seleccionado == "Guatemala":
        df = df[df["SECTOR"].astype(str).str.startswith("502")]

    df_filtrado = df.copy()
    for row in tabla.get_children():
        tabla.delete(row)

    if df.empty:
        messagebox.showwarning("Sin Resultados", "No se encontraron registros con esos filtros.")
    else:
        for _, fila in df.iterrows():
            tabla.insert("", "end", values=(fila["SECTOR"], fila["NOMBRE DEL SECTOR"], fila["FECHA DE VISITA"].strftime('%d/%m/%Y'), fila["CLASIFICACIÓN"], fila["TIPO CLIENTE"]))
        total_registros = len(df)
        tabla.insert("", "end", values=("TOTAL", "", "", total_registros, ""))

    # === CALCULAR HORAS A RESTAR DEL SEGUNDO EXCEL ===
    archivo_adicional = obtener_excel_especifico("ExportacaoAE_DNT-Ciclo 3.xlsx")
    if not archivo_adicional:
        return

    df_extra = pd.read_excel(archivo_adicional, usecols=["SECTOR CLIENTE", "PERÍODO", "FECHA_INCLUSION"])
    print("✅ Filas leídas de df_extra:", len(df_extra))
    print("📌 Primeras filas:")
    print(df_extra.head())
    print("📌 Columnas:", df_extra.columns.tolist())

    # Normalización correcta
    df_extra.columns = df_extra.columns.str.strip()
    df_extra["SECTOR CLIENTE"] = df_extra["SECTOR CLIENTE"].astype(str).str.strip()
    df_extra["PERÍODO"] = df_extra["PERÍODO"].astype(str).str.strip()
    df_extra["FECHA_INCLUSION"] = pd.to_datetime(df_extra["FECHA_INCLUSION"], errors="coerce", dayfirst=True)

    # Filtrar solo registros dentro del ciclo actual
    df_extra = df_extra[
        (df_extra["FECHA_INCLUSION"] >= fecha_ciclo_inicio) &
        (df_extra["FECHA_INCLUSION"] <= fecha_ciclo_fin)
    ]

    mapeo_nombres = {
        "50602": "ANA KAREN MORA",
        "50603": "RICARDO  ZUNIGA",
        "50604": "KAROLINA MONTERO",
        "50605": "KEREN CARVAJAL",
        "50606": "JOHAN ALBERTO ARCE NÚÑEZ"
    }

    def normalizar(texto):
        return unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8').strip().upper().replace("  ", " ")

    def extraer_horas(periodo):
        if pd.isnull(periodo):
            return 0
        texto = str(periodo).strip().upper()
        match = re.search(r'(\d+)', texto)
        if match:
            return int(match.group(1))
        else:
            print(f"[WARNING] No se encontró número en PERÍODO: '{texto}'")
            return 0

    global HORAS_POR_CONSULTOR
    HORAS_POR_CONSULTOR = {}

    nombres_en_df = df_filtrado["NOMBRE DEL SECTOR"].dropna().unique()
    for nombre in nombres_en_df:
        nombre_normalizado = normalizar(nombre)
        codigo_sector = None

        for codigo, nombre_mapeado in mapeo_nombres.items():
            if nombre_normalizado == normalizar(nombre_mapeado):
                codigo_sector = codigo
                break

        if not codigo_sector:
            print(f"[ADVERTENCIA] No se encontró código para el consultor '{nombre}'")
            continue

        # Verificación extra
        print(f"\n🔍 DEBUG - Consultor: {nombre_normalizado}")
        print(f"Buscando código: {codigo_sector}")
        print("👉 Valores únicos actuales en SECTOR CLIENTE:", df_extra["SECTOR CLIENTE"].unique())

        df_extra_filtrado = df_extra[df_extra["SECTOR CLIENTE"] == codigo_sector].copy()

        if df_extra_filtrado.empty:
            print(f"[INFO] No hay filas en df_extra para el código {codigo_sector}")
            continue

        df_extra_filtrado["HORAS"] = df_extra_filtrado["PERÍODO"].apply(extraer_horas)

        total_horas = df_extra_filtrado["HORAS"].sum()
        HORAS_POR_CONSULTOR[nombre_normalizado] = total_horas

        print(f"\n✅ Detalle para {nombre_normalizado} (código {codigo_sector}):")
        print(df_extra_filtrado[["FECHA_INCLUSION", "PERÍODO", "HORAS"]])

    print("\n=== HORAS A RESTAR POR CONSULTOR ===")
    for nombre, horas in HORAS_POR_CONSULTOR.items():
        print(f"{nombre}: {horas} horas")



def obtener_excel_especifico(nombre_archivo):
    archivo_path = os.path.join(ruta_carpeta, nombre_archivo)
    return archivo_path if os.path.exists(archivo_path) else None

def aplicar_filtros():
    CICLOS = {
        "CICLO 1": (datetime(2025, 1, 1), datetime(2025, 2, 10)),
        "CICLO 2": (datetime(2025, 2, 11), datetime(2025, 3, 13)),
        "CICLO 3": (datetime(2025, 3, 14), datetime(2025, 4, 12)),
        "CICLO 4": (datetime(2025, 4, 13), datetime(2025, 5, 15)),
    }

    def obtener_ciclo_actual(fecha_actual):
        for ciclo, (inicio, fin) in CICLOS.items():
            if inicio <= fecha_actual <= fin:
                return ciclo, inicio, fin
        return None, None, None

    fecha_actual = datetime.now()
    ciclo_actual, fecha_ciclo_inicio, fecha_ciclo_fin = obtener_ciclo_actual(fecha_actual)
    if not ciclo_actual:
        messagebox.showerror("Error", "No se pudo determinar el ciclo actual basado en la fecha.")
        return

    print(f"Ciclo actual: {ciclo_actual} ({fecha_ciclo_inicio.date()} a {fecha_ciclo_fin.date()})")

    global df_filtrado
    archivo_excel = obtener_ultimo_excel()
    if not archivo_excel:
        messagebox.showerror("Error", "No se encontró ningún archivo Excel en la carpeta.")
        return

    df = pd.read_excel(archivo_excel)
    columnas_necesarias = {"SECTOR", "NOMBRE DEL SECTOR", "FECHA DE VISITA", "CLASIFICACIÓN", "TIPO CLIENTE"}
    if not columnas_necesarias.issubset(df.columns):
        messagebox.showerror("Error", f"El archivo Excel no tiene las columnas necesarias: {columnas_necesarias}")
        return

    df["FECHA DE VISITA"] = pd.to_datetime(df["FECHA DE VISITA"], errors="coerce", dayfirst=True)

    # No aplicar filtros por fecha ni país

    df_filtrado = df.copy()

    for row in tabla.get_children():
        tabla.delete(row)

    if df.empty:
        messagebox.showwarning("Sin Resultados", "No se encontraron registros.")
    else:
        for _, fila in df.iterrows():
            tabla.insert("", "end", values=(fila["SECTOR"], fila["NOMBRE DEL SECTOR"], fila["FECHA DE VISITA"].strftime('%d/%m/%Y'), fila["CLASIFICACIÓN"], fila["TIPO CLIENTE"]))
        total_registros = len(df)
        tabla.insert("", "end", values=("TOTAL", "", "", total_registros, ""))

    # === CALCULAR HORAS A RESTAR DEL SEGUNDO EXCEL ===
    archivo_adicional = obtener_excel_especifico("ExportacaoAE_DNT-Ciclo 3.xlsx")
    if not archivo_adicional:
        return

    df_extra = pd.read_excel(archivo_adicional, usecols=["SECTOR CLIENTE", "PERÍODO", "FECHA_INCLUSION"])
    df_extra.columns = df_extra.columns.str.strip()
    df_extra["SECTOR CLIENTE"] = df_extra["SECTOR CLIENTE"].astype(str).str.strip()
    df_extra["PERÍODO"] = df_extra["PERÍODO"].astype(str).str.strip()
    df_extra["FECHA_INCLUSION"] = pd.to_datetime(df_extra["FECHA_INCLUSION"], errors="coerce", dayfirst=True)

    df_extra = df_extra[
        (df_extra["FECHA_INCLUSION"] >= fecha_ciclo_inicio) &
        (df_extra["FECHA_INCLUSION"] <= fecha_ciclo_fin)
    ]

    mapeo_nombres = {
        "50602": "ANA KAREN MORA",
        "50603": "RICARDO  ZUNIGA",
        "50604": "KAROLINA MONTERO",
        "50605": "KEREN CARVAJAL",
        "50606": "JOHAN ALBERTO ARCE NÚÑEZ"
    }

    def normalizar(texto):
        return unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8').strip().upper().replace("  ", " ")

    def extraer_horas(periodo):
        if pd.isnull(periodo):
            return 0
        texto = str(periodo).strip().upper()
        match = re.search(r'(\d+)', texto)
        if match:
            return int(match.group(1))
        else:
            print(f"[WARNING] No se encontró número en PERÍODO: '{texto}'")
            return 0

    global HORAS_POR_CONSULTOR
    HORAS_POR_CONSULTOR = {}

    nombres_en_df = df_filtrado["NOMBRE DEL SECTOR"].dropna().unique()
    for nombre in nombres_en_df:
        nombre_normalizado = normalizar(nombre)
        codigo_sector = None

        for codigo, nombre_mapeado in mapeo_nombres.items():
            if nombre_normalizado == normalizar(nombre_mapeado):
                codigo_sector = codigo
                break

        if not codigo_sector:
            print(f"[ADVERTENCIA] No se encontró código para el consultor '{nombre}'")
            continue

        df_extra_filtrado = df_extra[df_extra["SECTOR CLIENTE"] == codigo_sector].copy()

        if df_extra_filtrado.empty:
            print(f"[INFO] No hay filas en df_extra para el código {codigo_sector}")
            continue

        df_extra_filtrado["HORAS"] = df_extra_filtrado["PERÍODO"].apply(extraer_horas)

        total_horas = df_extra_filtrado["HORAS"].sum()
        HORAS_POR_CONSULTOR[nombre_normalizado] = total_horas

    print("\n=== HORAS A RESTAR POR CONSULTOR ===")
    for nombre, horas in HORAS_POR_CONSULTOR.items():
        print(f"{nombre}: {horas} horas")

    # === HORAS POR PAÍS Y CONSULTOR ===
    global HORAS_POR_PAIS_Y_CONSULTOR
    HORAS_POR_PAIS_Y_CONSULTOR = {}

    for nombre, horas in HORAS_POR_CONSULTOR.items():
        sectores_del_consultor = df_filtrado[
            df_filtrado["NOMBRE DEL SECTOR"].apply(lambda x: normalizar(str(x))) == nombre
        ]["SECTOR"].astype(str)
        if sectores_del_consultor.empty:
            print(f"[AVISO] No se encontró SECTOR para el consultor {nombre}")
            continue

        codigo_pais = sectores_del_consultor.iloc[0][:3]
        pais_nombre = {
            "506": "Costa Rica",
            "507": "Panamá",
            "505": "Nicaragua",
            "503": "El Salvador",
            "504": "Honduras",
            "502": "Guatemala"
        }.get(codigo_pais, "Otro")

        if pais_nombre not in HORAS_POR_PAIS_Y_CONSULTOR:
            HORAS_POR_PAIS_Y_CONSULTOR[pais_nombre] = {}

        HORAS_POR_PAIS_Y_CONSULTOR[pais_nombre][nombre] = horas

    print("\n=== HORAS A RESTAR POR PAÍS Y CONSULTOR ===")
    for pais, consultores in HORAS_POR_PAIS_Y_CONSULTOR.items():
        print(f"\n🌎 {pais}:")
        for consultor, horas in consultores.items():
            print(f"   👤 {consultor}: {horas} horas")





#------------------------------------------------------------------------------------------------------------------------------------
def generar_pdf_por_pais():
    if df_filtrado is None or df_filtrado.empty:
        messagebox.showwarning("Advertencia", "No hay datos filtrados para generar el PDF.")
        return

    ruta_pdf = os.path.join(ruta_carpeta, "reporte_por_pais.pdf")
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()

    pdf.set_font("Arial", "B", 14)
    pdf.cell(300, 10, "Resumen por País", ln=True, align="C")
    pdf.ln(10)

    df_filtrado["FECHA DE VISITA"] = pd.to_datetime(df_filtrado["FECHA DE VISITA"])
    fecha_inicio = df_filtrado["FECHA DE VISITA"].min().date()
    fecha_fin = df_filtrado["FECHA DE VISITA"].max().date()

    pdf.set_font("Arial", "B", 12)
    pdf.cell(300, 10, f"Datos desde {fecha_inicio.strftime('%d-%m-%Y')} hasta {fecha_fin.strftime('%d-%m-%Y')}", ln=True, align="C")
    pdf.ln(10)

    # Encabezados
    pdf.set_font("Arial", "B", 10)
    pdf.cell(35, 10, "PAÍS", 1, 0, "C")
    pdf.cell(55, 10, "CONSULTORES", 1, 0, "C")
    pdf.cell(50, 10, "PROMEDIO VISITAS", 1, 0, "C")
    pdf.cell(75, 10, "COBERTURA REAL", 1, 0, "C")
    pdf.cell(50, 10, "PROGRAMADOS", 1, 0, "C")
    pdf.cell(25, 10, "PANEL", 1, 1, "C")

    pdf.cell(35, 10, "", 1, 0, "C")
    pdf.cell(55, 10, "", 1, 0, "C")
    pdf.cell(25, 10, "MÉDICOS", 1, 0, "C")
    pdf.cell(25, 10, "FARMACIAS", 1, 0, "C")
    pdf.cell(25, 10, "Médicos", 1, 0, "C")
    pdf.cell(25, 10, "Farmacias", 1, 0, "C")
    pdf.cell(25, 10, "Médicos VIP", 1, 0, "C")
    pdf.cell(25, 10, "Médicos", 1, 0, "C")
    pdf.cell(25, 10, "Farmacias", 1, 0, "C")
    pdf.cell(25, 10, "Médicos VIP", 1, 1, "C")

    pdf.set_font("Arial", "", 10)

    paises = {"506": "Costa Rica", "507": "Panamá", "505": "Nicaragua"}

    for codigo_pais, nombre_pais in paises.items():
        df_pais = df_filtrado[df_filtrado["SECTOR"].astype(str).str.startswith(codigo_pais)]
        consultores_unicos = df_pais["NOMBRE DEL SECTOR"].dropna().unique()

        # Inicializamos acumuladores
        acum_promedio_medicos = 0
        acum_promedio_farmacias = 0
        acum_cobertura_medicos = 0
        acum_cobertura_farmacias = 0
        acum_cobertura_vip = 0
        acum_prog_medicos = 0
        acum_prog_farmacias = 0
        acum_prog_vip = 0
        total_consultores = 0

        for consultor in consultores_unicos:
            df_consultor = df_pais[df_pais["NOMBRE DEL SECTOR"] == consultor]

            total_medicos = len(df_consultor)
            total_vip = len(df_consultor[df_consultor["CLASIFICACIÓN"] == "VIP"])
            total_regentes = df_consultor["CLASIFICACIÓN"].isna().sum()
            var_arreglo = total_medicos - total_regentes

            total_medicos_vip_estandar = df_consultor[df_consultor["CLASIFICACIÓN"].isin(["VIP", "ESTANDAR"])].shape[0]
            total_medicos_vip = df_consultor[df_consultor["CLASIFICACIÓN"] == "VIP"].shape[0]
            total_medicos_regentes = df_consultor[df_consultor["TIPO CLIENTE"] == "PDV"].shape[0]

            fechas_unicas = sorted(df_consultor["FECHA DE VISITA"].dropna().unique())
            total_dias = len(fechas_unicas) if len(fechas_unicas) > 0 else 1

            RESTAR_HORAS = HORAS_POR_CONSULTOR.get(consultor, 0)

            dias_restados = RESTAR_HORAS / 8

            cobertura_medicos = round(var_arreglo / total_dias, 2)
            cobertura_regentes = round(total_regentes / total_dias, 2)

            if consultor == "ANA KAREN MORA":
                var_arreglo_real = 7.5 * (total_dias - dias_restados)
                var_arreglo_regentes = 4 * (total_dias - dias_restados)
                var_arreglo_vip = 1.8 * (total_dias - dias_restados)
            else:
                var_arreglo_real = 9 * (total_dias - dias_restados)
                var_arreglo_regentes = 6 * (total_dias - dias_restados)
                var_arreglo_vip = 4 * (total_dias - dias_restados)

            cobertura_medicos_pct = round((var_arreglo / var_arreglo_real) * 100, 2) if var_arreglo_real > 0 else 0
            cobertura_regentes_pct = round((total_regentes / var_arreglo_regentes) * 100, 2) if var_arreglo_regentes > 0 else 0
            cobertura_vip_pct = round((total_vip / var_arreglo_vip) * 100, 2) if var_arreglo_vip > 0 else 0

            # Acumulamos
            acum_promedio_medicos += cobertura_medicos
            acum_promedio_farmacias += cobertura_regentes
            acum_cobertura_medicos += cobertura_medicos_pct
            acum_cobertura_farmacias += cobertura_regentes_pct
            acum_cobertura_vip += cobertura_vip_pct
            acum_prog_medicos += total_medicos_vip_estandar
            acum_prog_farmacias += total_medicos_regentes
            acum_prog_vip += total_medicos_vip
            total_consultores += 1

            # Fila consultor
            pdf.cell(35, 10, nombre_pais, 1, 0, "C")
            pdf.cell(55, 10, consultor, 1, 0, "C")
            pdf.cell(25, 10, str(cobertura_medicos), 1, 0, "C")
            pdf.cell(25, 10, str(cobertura_regentes), 1, 0, "C")
            pdf.cell(25, 10, f"{cobertura_medicos_pct:.2f}%", 1, 0, "C")
            pdf.cell(25, 10, f"{cobertura_regentes_pct:.2f}%", 1, 0, "C")
            pdf.cell(25, 10, f"{cobertura_vip_pct:.2f}%", 1, 0, "C")
            pdf.cell(25, 10, str(total_medicos_vip_estandar), 1, 0, "C")
            pdf.cell(25, 10, str(total_medicos_regentes), 1, 0, "C")
            pdf.cell(25, 10, str(total_medicos_vip), 1, 1, "C")

        # Fila de totales del país
        if total_consultores > 0:
            pdf.set_font("Arial", "B", 10)
            pdf.cell(35, 10, nombre_pais, 1, 0, "C")
            pdf.cell(55, 10, "TOTAL", 1, 0, "C")
            pdf.cell(25, 10, str(round(acum_promedio_medicos / total_consultores, 2)), 1, 0, "C")
            pdf.cell(25, 10, str(round(acum_promedio_farmacias / total_consultores, 2)), 1, 0, "C")
            pdf.cell(25, 10, f"{round(acum_cobertura_medicos / total_consultores, 2)}%", 1, 0, "C")
            pdf.cell(25, 10, f"{round(acum_cobertura_farmacias / total_consultores, 2)}%", 1, 0, "C")
            pdf.cell(25, 10, f"{round(acum_cobertura_vip / total_consultores, 2)}%", 1, 0, "C")
            pdf.cell(25, 10, str(acum_prog_medicos), 1, 0, "C")
            pdf.cell(25, 10, str(acum_prog_farmacias), 1, 0, "C")
            pdf.cell(25, 10, str(acum_prog_vip), 1, 1, "C")
            pdf.set_font("Arial", "", 10)

    pdf.output(ruta_pdf)
    messagebox.showinfo("PDF Generado", f"El reporte se ha guardado en:\n{ruta_pdf}")
    return ruta_pdf


#--------------------------------------------------------------------
#///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
def extraer_todas_las_tablas_pdf(ruta_pdf):
    import pdfplumber

    tabla_html = ""

    with pdfplumber.open(ruta_pdf) as pdf:
        for page in pdf.pages:
            tablas = page.extract_tables()

            if tablas:
                for i, tabla in enumerate(tablas):
                    #tabla_html += f"<h3>Resumen País {i + 1}</h3>"
                    tabla_html += "<table border='1' style='border-collapse: collapse; width: 100%; text-align: center;'>"

                    for fila_index, fila in enumerate(tabla):
                        tabla_html += "<tr>"
                        skip_celdas = 0

                        for col_index, celda in enumerate(fila):
                            if skip_celdas > 0:
                                skip_celdas -= 1
                                continue

                            if celda is None:
                                celda = ""

                            if fila_index == 0:  # Primera fila (cabecera principal)
                                colspan = 1
                                for next_col in fila[col_index + 1:]:
                                    if next_col is None:
                                        colspan += 1
                                    else:
                                        break

                                if colspan > 1:
                                    tabla_html += f"<th colspan='{colspan}'>{celda}</th>"
                                    skip_celdas = colspan - 1
                                else:
                                    tabla_html += f"<th>{celda}</th>"

                            elif fila_index == 1:  # Segunda fila (subencabezado)
                                tabla_html += f"<th>{celda}</th>"

                            else:  # Resto de filas (datos)
                                tabla_html += f"<td>{celda}</td>"

                        tabla_html += "</tr>"

                    tabla_html += "</table><br>"

    return tabla_html if tabla_html else "<p>No se encontraron tablas en el PDF.</p>"


def es_correo_valido(correo):
    """Verifica si el correo tiene un formato válido usando una expresión regular."""
    patron = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
    return re.match(patron, correo) is not None

def enviar_correo():
    #ruta_pdf = generar_pdf_extra()
    correo_destino = entrada_correo.get().strip()
    if not correo_destino:
        messagebox.showwarning("Advertencia", "No se ingresó un correo. El reporte no se envió.")
        return
    if not es_correo_valido(correo_destino):
        messagebox.showerror("Error", "El correo ingresado no es válido. Inténtalo de nuevo.")
        return
    ruta_pdf = generar_pdf_por_pais()



    tabla_html = extraer_todas_las_tablas_pdf(ruta_pdf)  # Extrae todas las tablas como HTML


    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    lista_destinatarios = [correo_destino, "xotoyip215@buides.com", "jose03pif@gmail.com"]
    mail.To = "; ".join(lista_destinatarios)

    #mail.To = correo_destino
    mail.Subject = "AGD - Resumen de Gerentes Distritales"
    mail.HTMLBody = f"""
    <p>A continuación se envía un resumen del proceso</p>

    {tabla_html}

    <p>Se marcan en rojo promedios de visitas a médicos que no alcanzan el objetivo de contactos programados.</p> 
    <p>Se marcan en rojo promedios de visitas a farmacias que no alcanzan el objetivo de contactos programados</p>
    <p>Se marcan en rojo coberturas menores a la cobertura teórica</p>

    """


    #mail.Body = "Adjunto el reporte filtrado."
    mail.Attachments.Add(ruta_pdf)
    mail.Send()
    messagebox.showinfo("Correo Enviado", "El correo ha sido enviado correctamente.")

# ======================== VENTANAS ========================


def abrir_reporte_por_pais():
   # messagebox.showinfo("Otra Funcionalidad", "Aquí puedes agregar otra funcionalidad.")
    #ventana_inicio.withdraw()
    global root, tabla, entrada_nombre, entrada_fecha_inicio, entrada_fecha_fin, combo_pais, combo_tipo_cliente, clasificacion_listbox, entrada_correo

    root = tk.Toplevel()
    root.title("Sistema de Reportes")
    root.geometry("1020x450")

    

    # Botón para aplicar filtros
    btn_filtrar = ttk.Button(root, text="🔍 Traer Datos", command=aplicar_filtros)
    btn_filtrar.grid(row=6, column=0, columnspan=2, pady=10)

       # Botón para aplicar filtros
    #btn_filtrar = ttk.Button(root, text="🔍 Aplicar Filtros 2", command=aplicar_filtros_extra)   #--------------------------------
    #btn_filtrar.grid(row=6, column=0, columnspan=2, pady=10)

    # Tabla para mostrar resultados
    columnas = ("Sector", "Nombre del Sector", "Fecha de Visita", "Clasificación", "TIPO CLIENTE")
    tabla = ttk.Treeview(root, columns=columnas, show="headings")
    for col in columnas:
        tabla.heading(col, text=col)
    tabla.grid(row=7, column=0, columnspan=2, padx=10, pady=5)


     # Marco para contener la tabla y el scroll
    frame_tabla = ttk.Frame(root)
    frame_tabla.grid(row=7, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")

    # Scrollbars
    scrollbar_y = ttk.Scrollbar(frame_tabla, orient="vertical")
    scrollbar_x = ttk.Scrollbar(frame_tabla, orient="horizontal")

    columnas = ("Sector", "Nombre del Sector", "Fecha de Visita", "Clasificación", "TIPO CLIENTE")
    tabla = ttk.Treeview(frame_tabla, columns=columnas, show="headings", yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

    for col in columnas:
        tabla.heading(col, text=col)
        tabla.column(col, width=200)  # Ajusta el ancho de las columnas

    tabla.pack(side="top", fill="both", expand=True)
    
    scrollbar_y.pack(side="right", fill="y")
    scrollbar_x.pack(side="bottom", fill="x")

    scrollbar_y.config(command=tabla.yview)
    scrollbar_x.config(command=tabla.xview)


    # Entrada para correo
    ttk.Label(root, text="Correo Destino:").grid(row=10, column=0, padx=10, pady=5)
    entrada_correo = ttk.Entry(root)
    entrada_correo.grid(row=8, column=0, padx=10, pady=5)

    btn_enviar = ttk.Button(root, text="📧 Enviar Correo", command=enviar_correo)
    btn_enviar.grid(row=9, column=0)

  #  btn_generar_pdf_sin_duplicados = ttk.Button(root, text="📄 Reporte Total", command=generar_pdf_por_pais)
   #  btn_generar_pdf_sin_duplicados.grid(row=8, column=1, padx=10, pady=5)  # Posicionarlo a la par del otro botón


    #btn_regresar = ttk.Button(root, text="⬅ Volver", command=lambda: (root.destroy(), ventana_inicio.deiconify()))
    #btn_regresar.grid(row=15, column=0)




