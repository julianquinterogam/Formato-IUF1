import calendar
from datetime import date
from io import BytesIO

import pandas as pd
import streamlit as st

st.set_page_config(page_title=Generador Reporte SUI, layout=centered)

# ---------------------------------------------------------------------------
# Orden exacto de columnas del archivo de salida (IUF1_MM_AAAA)
# ---------------------------------------------------------------------------
OUTPUT_COLUMNS = [
    NIU, COD LOCALIDAD, ALTITUD, LONGITUD, LATITUD, ID FACTURA,
    ENERGIA GEN MES, TIPO CORRI SALIDA, DIAS PRES MES, IUC, MP,
    COR, GIO, PE, GAOM, DIRECCION, FECH EXP FACT,
    FECH INI PERIO, DIAS FACT, ESTRATO, TIPO LECT, FACT CONSUMO,
    VAL REFACT, VAL MORA, INT MORA, VAL SUBS, PORCE SUBS,
    TARIFA, VAL TOTAL FACT,
]

MESES = {
    1 Enero, 2 Febrero, 3 Marzo, 4 Abril, 5 Mayo, 6 Junio,
    7 Julio, 8 Agosto, 9 Septiembre, 10 Octubre,
    11 Noviembre, 12 Diciembre,
}


def generar_reporte(df_znisisfv pd.DataFrame, df_usuarios pd.DataFrame,
                     mes int, anio int) - pd.DataFrame
    Construye el DataFrame de salida siguiendo las reglas del reporte SUI.

    df_znisisfv = df_znisisfv.copy()
    df_usuarios = df_usuarios.copy()

    # Normalizar la columna llave para evitar fallos de cruce por tipo de dato
    df_znisisfv[NIU] = df_znisisfv[NIU].astype(str).str.strip()
    df_usuarios[NIU_SUI] = df_usuarios[NIU_SUI].astype(str).str.strip()

    homolog = df_usuarios.set_index(NIU_SUI)[[LONGITUD, LATITUD, VEREDA]]

    out = pd.DataFrame(index=df_znisisfv.index)

    # --- Campos tomados directamente de ZNISISFV ---
    out[NIU] = df_znisisfv[NIU]
    out[COD LOCALIDAD] = df_znisisfv[COD_LOCALIDAD]
    out[ID FACTURA] = df_znisisfv[ID_FACTURA]
    out[ENERGIA GEN MES] = df_znisisfv[CONSUMO_ENERGIA].round(0).astype(Int64)
    out[DIAS PRES MES] = df_znisisfv[DIAS_PRESTACION]
    out[DIAS FACT] = df_znisisfv[DIAS_PRESTACION]
    out[FACT CONSUMO] = df_znisisfv[FACT_CONSUMO]
    out[VAL MORA] = df_znisisfv[VALOR_MORA]
    out[VAL SUBS] = df_znisisfv[VALOR_SUBSIDIO]
    out[TARIFA] = df_znisisfv[VALOR_TARIFA]

    # --- Campos homologados desde Usuarios (por NIU = NIU_SUI) ---
    merged = df_znisisfv[NIU].map(homolog[LONGITUD])
    out[LONGITUD] = merged
    out[LATITUD] = df_znisisfv[NIU].map(homolog[LATITUD])
    out[DIRECCION] = df_znisisfv[NIU].map(homolog[VEREDA])

    # --- Valores constantes  calculados ---
    ultimo_dia = calendar.monthrange(anio, mes)[1]
    fecha_ini = date(anio, mes, 1)
    fecha_fin = date(anio, mes, ultimo_dia)

    out[ALTITUD] = 0
    out[TIPO CORRI SALIDA] = 1
    out[IUC] = 
    out[MP] = 
    out[COR] = 
    out[GIO] = 
    out[PE] = 
    out[GAOM] = 
    out[FECH EXP FACT] = fecha_fin.strftime(%d-%m-%Y)
    out[FECH INI PERIO] = fecha_ini.strftime(%d-%m-%Y)
    out[ESTRATO] = 1
    out[TIPO LECT] = 3
    out[VAL REFACT] = 0
    out[INT MORA] = 0
    out[PORCE SUBS] = out[VAL SUBS]  out[FACT CONSUMO]
    out[VAL TOTAL FACT] = out[FACT CONSUMO]

    return out[OUTPUT_COLUMNS]


def to_excel_bytes(df pd.DataFrame) - bytes
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine=openpyxl) as writer
        df.to_excel(writer, index=False, sheet_name=Hoja1)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title(Generador de Reporte SUI (IUF1))
st.write(
    Sube el archivo ZNISISFV (fuente principal) y el archivo Usuarios 
    (homologación por NIU) para generar el reporte con el formato IUF1.
)

col1, col2 = st.columns(2)
with col1
    archivo_znisisfv = st.file_uploader(Archivo ZNISISFV (Formato54), type=[xlsx])
with col2
    archivo_usuarios = st.file_uploader(Archivo Usuarios (BD_Usuarios), type=[xlsx])

col3, col4 = st.columns(2)
with col3
    mes = st.selectbox(Mes a generar, options=list(MESES.keys()),
                        format_func=lambda m f{m02d} - {MESES[m]})
with col4
    anio = st.number_input(Año a generar, min_value=2000, max_value=2100,
                            value=date.today().year, step=1)

if st.button(Generar reporte, type=primary)
    if archivo_znisisfv is None or archivo_usuarios is None
        st.error(Debes cargar los dos archivos antes de generar el reporte.)
    else
        try
            df_znisisfv = pd.read_excel(archivo_znisisfv, sheet_name=Formato54)
            df_usuarios = pd.read_excel(archivo_usuarios, sheet_name=BD_Usuarios)

            resultado = generar_reporte(df_znisisfv, df_usuarios, mes, anio)

            faltantes = resultado[LONGITUD].isna().sum()
            if faltantes
                st.warning(
                    f{faltantes} registro(s) no tuvieron coincidencia en el archivo 
                    Usuarios (LONGITUDLATITUDDIRECCION quedaron vacíos).
                )

            st.success(fReporte generado con {len(resultado)} registros.)
            st.dataframe(resultado.head(20))

            nombre_archivo = fIUF1_{mes02d}_{anio}.xlsx
            st.download_button(
                label=fDescargar {nombre_archivo},
                data=to_excel_bytes(resultado),
                file_name=nombre_archivo,
                mime=applicationvnd.openxmlformats-officedocument.spreadsheetml.sheet,
            )
        except Exception as e
            st.error(fOcurrió un error generando el reporte {e})
