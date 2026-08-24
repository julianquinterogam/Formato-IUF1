import calendar
from datetime import date
from io import BytesIO

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Generador Reporte SUI", layout="centered")

# ---------------------------------------------------------------------------
# Orden exacto de columnas del archivo de salida (IUF1_MM_AAAA)
# ---------------------------------------------------------------------------
OUTPUT_COLUMNS = [
    "NIU", "COD LOCALIDAD", "ALTITUD", "LONGITUD", "LATITUD", "ID FACTURA",
    "ENERGIA GEN MES", "TIPO CORRI SALIDA", "DIAS PRES MES", "IUC", "MP",
    "COR", "GIO", "PE", "GAOM", "DIRECCION", "FECH EXP FACT",
    "FECH INI PERIO", "DIAS FACT", "ESTRATO", "TIPO LECT", "FACT CONSUMO",
    "VAL REFACT", "VAL MORA", "INT MORA", "VAL SUBS", "PORCE SUBS",
    "TARIFA", "VAL TOTAL FACT",
]

MESES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre",
    11: "Noviembre", 12: "Diciembre",
}


PREFIJO_ZNISISFV = "ZNISISFV_54787_54_"


def nombre_znisisfv_esperado(mes: int, anio: int) -> str:
    """El archivo Formato54 se nombra con el mes SIGUIENTE al periodo reportado."""
    if mes == 12:
        mes_sig, anio_sig = 1, anio + 1
    else:
        mes_sig, anio_sig = mes + 1, anio
    return f"{PREFIJO_ZNISISFV}{mes_sig:02d}{anio_sig}.xlsx"


def generar_reporte(df_znisisfv: pd.DataFrame, df_usuarios: pd.DataFrame,
                     mes: int, anio: int) -> pd.DataFrame:
    """Construye el DataFrame de salida siguiendo las reglas del reporte SUI."""

    df_znisisfv = df_znisisfv.copy()
    df_usuarios = df_usuarios.copy()

    # Normalizar la columna llave para evitar fallos de cruce por tipo de dato
    df_znisisfv["NIU"] = df_znisisfv["NIU"].astype(str).str.strip()
    df_usuarios["NIU_SUI"] = df_usuarios["NIU_SUI"].astype(str).str.strip()

    homolog = df_usuarios.set_index("NIU_SUI")[["LONGITUD", "LATITUD", "VEREDA"]]

    out = pd.DataFrame(index=df_znisisfv.index)

    # --- Campos tomados directamente de ZNISISFV ---
    out["NIU"] = df_znisisfv["NIU"]
    out["COD LOCALIDAD"] = pd.to_numeric(
        df_znisisfv["COD_LOCALIDAD"], errors="coerce"
    ).round(0).astype("Int64")
    out["ID FACTURA"] = df_znisisfv["ID_FACTURA"]
    out["ENERGIA GEN MES"] = df_znisisfv["CONSUMO_ENERGIA"].round(0).astype("Int64")
    out["DIAS PRES MES"] = df_znisisfv["DIAS_PRESTACION"]
    out["DIAS FACT"] = df_znisisfv["DIAS_PRESTACION"]
    out["FACT CONSUMO"] = df_znisisfv["FACT_CONSUMO"]
    out["VAL MORA"] = df_znisisfv["VALOR_MORA"]
    out["VAL SUBS"] = df_znisisfv["VALOR_SUBSIDIO"]
    out["TARIFA"] = df_znisisfv["VALOR_TARIFA"]

    # --- Campos homologados desde Usuarios (por NIU = NIU_SUI) ---
    merged = df_znisisfv["NIU"].map(homolog["LONGITUD"])
    out["LONGITUD"] = merged
    out["LATITUD"] = df_znisisfv["NIU"].map(homolog["LATITUD"])
    out["DIRECCION"] = df_znisisfv["NIU"].map(homolog["VEREDA"])

    # --- Valores constantes / calculados ---
    ultimo_dia = calendar.monthrange(anio, mes)[1]
    fecha_ini = date(anio, mes, 1)
    fecha_fin = date(anio, mes, ultimo_dia)

    out["ALTITUD"] = 0
    out["TIPO CORRI SALIDA"] = 1
    out["IUC"] = ""
    out["MP"] = ""
    out["COR"] = ""
    out["GIO"] = ""
    out["PE"] = ""
    out["GAOM"] = ""
    out["FECH EXP FACT"] = fecha_fin.strftime("%d-%m-%Y")
    out["FECH INI PERIO"] = fecha_ini.strftime("%d-%m-%Y")
    out["ESTRATO"] = 1
    out["TIPO LECT"] = 3
    out["VAL REFACT"] = 0
    out["INT MORA"] = 0
    out["PORCE SUBS"] = (out["VAL SUBS"] / out["FACT CONSUMO"]).replace(
        [float("inf"), float("-inf")], 0
    ).fillna(0).round(4)
    out["VAL TOTAL FACT"] = out["FACT CONSUMO"]

    return out[OUTPUT_COLUMNS]


CAMPOS_TOTAL = ["TARIFA", "VAL SUBS", "FACT CONSUMO"]


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Hoja1")

        worksheet = writer.sheets["Hoja1"]

        # Evitar notación científica en columnas de códigos numéricos largos
        columnas_enteras_largas = ["NIU", "COD LOCALIDAD"]
        for campo in columnas_enteras_largas:
            col_idx = df.columns.get_loc(campo) + 1  # 1-based
            col_letter = worksheet.cell(row=1, column=col_idx).column_letter
            for fila in range(2, len(df) + 2):
                worksheet[f"{col_letter}{fila}"].number_format = "0"

        # Formato con 4 decimales
        columnas_4_decimales = [
            "FACT CONSUMO", "VAL REFACT", "VAL MORA", "INT MORA",
            "VAL SUBS", "PORCE SUBS", "TARIFA", "VAL TOTAL FACT",
        ]
        for campo in columnas_4_decimales:
            col_idx = df.columns.get_loc(campo) + 1  # 1-based
            col_letter = worksheet.cell(row=1, column=col_idx).column_letter
            for fila in range(2, len(df) + 2):
                worksheet[f"{col_letter}{fila}"].number_format = "0.0000"

    return buffer.getvalue()


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("Generador de Reporte SUI (IUF1)")
st.write(
    "Sube el archivo **ZNISISFV** (fuente principal) y el archivo **Usuarios** "
    "(homologación por NIU) para generar el reporte con el formato IUF1."
)

col1, col2 = st.columns(2)
with col1:
    archivo_znisisfv = st.file_uploader("Archivo ZNISISFV (Formato54)", type=["xlsx"])
with col2:
    archivo_usuarios = st.file_uploader("Archivo Usuarios (BD_Usuarios)", type=["xlsx"])

col3, col4 = st.columns(2)
with col3:
    mes = st.selectbox("Mes a generar", options=list(MESES.keys()),
                        format_func=lambda m: f"{m:02d} - {MESES[m]}")
with col4:
    anio = st.number_input("Año a generar", min_value=2000, max_value=2100,
                            value=date.today().year, step=1)

if st.button("Generar reporte", type="primary"):
    if archivo_znisisfv is None or archivo_usuarios is None:
        st.error("Debes cargar los dos archivos antes de generar el reporte.")
    else:
        nombre_esperado = nombre_znisisfv_esperado(mes, anio)
        if archivo_znisisfv.name != nombre_esperado:
            st.error(
                f"El archivo ZNISISFV cargado se llama **{archivo_znisisfv.name}**, "
                f"pero para generar el reporte de {MESES[mes]} {anio} se espera el "
                f"archivo **{nombre_esperado}** (Formato 54 usa el mes siguiente al "
                "periodo reportado). Sube el archivo correcto para continuar."
            )
        else:
            try:
                df_znisisfv = pd.read_excel(archivo_znisisfv, sheet_name="Formato54")
                df_usuarios = pd.read_excel(archivo_usuarios, sheet_name="BD_Usuarios")

                resultado = generar_reporte(df_znisisfv, df_usuarios, mes, anio)

                faltantes = resultado["LONGITUD"].isna().sum()
                if faltantes:
                    st.warning(
                        f"{faltantes} registro(s) no tuvieron coincidencia en el archivo "
                        "Usuarios (LONGITUD/LATITUD/DIRECCION quedaron vacíos)."
                    )

                st.success(f"Reporte generado con {len(resultado)} registros.")

                st.write("**Totales de validación** (no se incluyen en el Excel descargado):")
                t1, t2, t3 = st.columns(3)
                t1.metric("Suma TARIFA", f"${resultado['TARIFA'].sum():,.0f}")
                t2.metric("Suma VAL SUBS", f"${resultado['VAL SUBS'].sum():,.0f}")
                t3.metric("Suma FACT CONSUMO", f"${resultado['FACT CONSUMO'].sum():,.0f}")

                st.dataframe(resultado.head(20))

                nombre_archivo = f"IUF1_{mes:02d}_{anio}.xlsx"
                st.download_button(
                    label=f"Descargar {nombre_archivo}",
                    data=to_excel_bytes(resultado),
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Ocurrió un error generando el reporte: {e}")
