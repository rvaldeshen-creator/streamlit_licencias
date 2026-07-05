import streamlit as st
import pandas as pd
import plotly.express as px
import seaborn as sns
import matplotlib.pyplot as plt
from io import BytesIO
import base64

try:
    import pygwalker as pyg
    import streamlit.components.v1 as components
    PYGWALKER_OK = True
except Exception:
    PYGWALKER_OK = False

def agregar_fondo_transparente(ruta_imagen, mostrar_imagen=True):
    if mostrar_imagen:
        with open(ruta_imagen, "rb") as archivo:
            imagen_base64 = base64.b64encode(archivo.read()).decode()

        fondo = f"""
            background-image:
                linear-gradient(rgba(5, 10, 20, 0.95), rgba(5, 10, 20, 0.95)),
                url("data:image/jpg;base64,{imagen_base64}");
            background-size: cover;
            background-position: center;
            background-repeat: no-repeat;
            background-attachment: fixed;
        """
    else:
        fondo = """
            background-color: #050a14;
            background-image: none;
        """

    st.markdown(
        f"""
        <style>
        [data-testid="stAppViewContainer"] {{
            {fondo}
        }}

        [data-testid="stSidebar"] {{
            background-color: rgba(13, 17, 23, 0.96);
        }}
        </style>
        """,
        unsafe_allow_html=True
    )
    

# ============================================================
# CONFIGURACIÓN GENERAL DE LA APP
# ============================================================

st.set_page_config(
    page_title="Dashboard Licencias Médicas",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

mostrar_fondo_licencias = st.sidebar.checkbox(
    "Mostrar imagen de fondo",
    value=True
)

agregar_fondo_transparente(
    "fondo_licencias.jpg",
    mostrar_imagen=mostrar_fondo_licencias
)

# ============================================================
# ESTILO VISUAL DEL DASHBOARD
# ============================================================

st.markdown("""
<style>

.kpi-card {
    background: linear-gradient(135deg, #1e2130, #252d42);
    border-left: 4px solid #4fc3f7;
    border-radius: 10px;
    padding: 18px 20px;
    margin: 8px 0;
    height: 160px;
    width: 100%;
    box-sizing: border-box;
    display: flex;
    flex-direction: column;
    justify-content: center;
}

.kpi-card h2 {
    color: #4fc3f7;
    font-size: 2rem;
    margin: 0;
    font-weight: 700;
}

.kpi-card p {
    color: #c9d1d9;
    margin: 8px 0 0;
    font-size: 0.9rem;
}

.info-box {
    background: #1a2236;
    border-left: 4px solid #43a047;
    border-radius: 8px;
    padding: 14px 18px;
    margin: 10px 0;
    color: #c8e6c9;
    font-size: 0.95rem;
}

.warning-box {
    background: #332701;
    border-left: 4px solid #f9a825;
    border-radius: 8px;
    padding: 14px 18px;
    margin: 10px 0;
    color: #fff3cd;
    font-size: 0.95rem;
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def formato_miles(valor):
    try:
        return f"{valor:,.0f}".replace(",", ".")
    except Exception:
        return "0"


def formato_pesos(valor):
    try:
        return "$" + f"{valor:,.0f}".replace(",", ".")
    except Exception:
        return "$0"


def formato_decimal(valor, decimales=1):
    try:
        return f"{valor:.{decimales}f}"
    except Exception:
        return "0"


def convertir_a_excel(df_exportar):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_exportar.to_excel(writer, index=False, sheet_name="Licencias")
    return output.getvalue()


def aplicar_layout(fig, altura=480):
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#e6edf3"),
        margin=dict(l=20, r=20, t=70, b=40),
        height=altura
    )
    return fig


def kpi_card(col, numero, texto, icono=""):
    col.markdown(
        f"""
        <div class="kpi-card">
            <h2>{numero}</h2>
            <p>{icono} {texto}</p>
        </div>
        """,
        unsafe_allow_html=True
    )


def info(texto):
    st.markdown(
        f"""
        <div class="info-box">
            {texto}
        </div>
        """,
        unsafe_allow_html=True
    )


def warning_box(texto):
    st.markdown(
        f"""
        <div class="warning-box">
            {texto}
        </div>
        """,
        unsafe_allow_html=True
    )



# ============================================================
# CARGA DE DATOS
# ============================================================

@st.cache_data
def cargar_datos():
    df = pd.read_csv("licencias_limpias_streamlit.csv", low_memory=False, encoding="utf-8-sig")
    df.columns = df.columns.str.strip()

    columnas_numericas = [
        "ANIO_EMISION",
        "TRIMESTRE_EMISION",
        "NUMERO_DE_DIAS",
        "NUMERO_DE_DIAS_AUTORIZADOS",
        "MONTO_SUBSIDIO_LIQUIDO"
    ]

    for col in columnas_numericas:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if "LICENCIA_LARGA" not in df.columns and "NUMERO_DE_DIAS_AUTORIZADOS" in df.columns:
        df["LICENCIA_LARGA"] = df["NUMERO_DE_DIAS_AUTORIZADOS"].apply(
            lambda x: "Sí" if pd.notna(x) and x > 30 else "No"
        )

    return df


df = cargar_datos()

# Dejar solo años válidos para el análisis
df["ANIO_EMISION"] = pd.to_numeric(df["ANIO_EMISION"], errors="coerce")

df = df[df["ANIO_EMISION"].isin([2025, 2026])]


# ============================================================
# VALIDACIÓN DE COLUMNAS NECESARIAS
# ============================================================

columnas_necesarias = [
    "ANIO_EMISION",
    "TRIMESTRE_EMISION",
    "ISAPRE",
    "SEXO_TRABAJADOR",
    "NUMERO_DE_DIAS",
    "NUMERO_DE_DIAS_AUTORIZADOS",
    "MONTO_SUBSIDIO_LIQUIDO"
]

columnas_faltantes = [col for col in columnas_necesarias if col not in df.columns]

if len(columnas_faltantes) > 0:
    st.error("Faltan columnas necesarias en la base:")
    st.write(columnas_faltantes)

    st.write("Columnas disponibles en tu archivo:")
    st.write(df.columns.tolist())

    st.stop()


# ============================================================
# SIDEBAR: FILTROS DEL PANEL
# ============================================================

st.sidebar.markdown("### 🏥 Filtros del Panel")
st.sidebar.caption("Dashboard de Licencias Médicas")
st.sidebar.divider()

anio = st.sidebar.multiselect(
    "Año de emisión",
    options=[2025, 2026],
    default=[2025, 2026]
)

trimestre = st.sidebar.multiselect(
    "Trimestre",
    options=sorted(df["TRIMESTRE_EMISION"].dropna().unique()),
    default=sorted(df["TRIMESTRE_EMISION"].dropna().unique())
)

isapre = st.sidebar.multiselect(
    "Isapre",
    options=sorted(df["ISAPRE"].dropna().unique()),
    default=sorted(df["ISAPRE"].dropna().unique())
)

sexo = st.sidebar.multiselect(
    "Sexo del trabajador",
    options=sorted(df["SEXO_TRABAJADOR"].dropna().unique()),
    default=sorted(df["SEXO_TRABAJADOR"].dropna().unique())
)

if "TIPO_DE_LICENCIA" in df.columns:
    tipo_licencia = st.sidebar.multiselect(
        "Tipo de licencia",
        options=sorted(df["TIPO_DE_LICENCIA"].dropna().unique()),
        default=sorted(df["TIPO_DE_LICENCIA"].dropna().unique())
    )
else:
    tipo_licencia = None

if "TIPO_DE_RESOLUCION" in df.columns:
    tipo_resolucion = st.sidebar.multiselect(
        "Tipo de resolución",
        options=sorted(df["TIPO_DE_RESOLUCION"].dropna().unique()),
        default=sorted(df["TIPO_DE_RESOLUCION"].dropna().unique())
    )
else:
    tipo_resolucion = None

if "LICENCIA_LARGA" in df.columns:
    licencia_larga = st.sidebar.multiselect(
        "Licencia larga",
        options=sorted(df["LICENCIA_LARGA"].dropna().unique()),
        default=sorted(df["LICENCIA_LARGA"].dropna().unique())
    )
else:
    licencia_larga = None


# ============================================================
# APLICACIÓN DE FILTROS
# ============================================================

df_filtrado = df[
    (df["ANIO_EMISION"].isin(anio)) &
    (df["TRIMESTRE_EMISION"].isin(trimestre)) &
    (df["ISAPRE"].isin(isapre)) &
    (df["SEXO_TRABAJADOR"].isin(sexo))
]

if tipo_licencia is not None:
    df_filtrado = df_filtrado[
        df_filtrado["TIPO_DE_LICENCIA"].isin(tipo_licencia)
    ]

if tipo_resolucion is not None:
    df_filtrado = df_filtrado[
        df_filtrado["TIPO_DE_RESOLUCION"].isin(tipo_resolucion)
    ]

if licencia_larga is not None:
    df_filtrado = df_filtrado[
        df_filtrado["LICENCIA_LARGA"].isin(licencia_larga)
    ]


if len(df_filtrado) == 0:
    st.warning("No hay registros con los filtros seleccionados. Modifica los filtros del panel lateral.")
    st.stop()


# ============================================================
# ENCABEZADO PRINCIPAL
# ============================================================

st.title("🏥 Dashboard Licencias Médicas")

st.markdown(
    """
    Aplicativo para analizar licencias médicas, resoluciones, diagnósticos,
    días autorizados, subsidios e Isapres a partir de la base limpia.
    """
)

st.divider()


# ============================================================
# INDICADORES GENERALES
# ============================================================

total_licencias = len(df_filtrado)

promedio_dias_solicitados = df_filtrado["NUMERO_DE_DIAS"].mean()
promedio_dias_autorizados = df_filtrado["NUMERO_DE_DIAS_AUTORIZADOS"].mean()
promedio_subsidio_liquido = df_filtrado["MONTO_SUBSIDIO_LIQUIDO"].mean()

if "TIPO_DE_RESOLUCION" in df_filtrado.columns and len(df_filtrado) > 0:
    autorizadas = df_filtrado[
        df_filtrado["TIPO_DE_RESOLUCION"]
        .astype(str)
        .str.upper()
        .str.contains("AUTORIZ", na=False)
    ]

    tasa_autorizacion = len(autorizadas) / len(df_filtrado) * 100
else:
    tasa_autorizacion = 0

if "LICENCIA_LARGA" in df_filtrado.columns:
    cantidad_largas = len(
        df_filtrado[
            df_filtrado["LICENCIA_LARGA"]
            .astype(str)
            .str.upper()
            .isin(["SI", "SÍ", "TRUE", "1"])
        ]
    )
else:
    cantidad_largas = len(
        df_filtrado[df_filtrado["NUMERO_DE_DIAS_AUTORIZADOS"] > 30]
    )

if total_licencias > 0:
    porcentaje_licencias_largas = cantidad_largas / total_licencias * 100
else:
    porcentaje_licencias_largas = 0



c1, c2, c3, c4 = st.columns(4)

kpi_card(c1, formato_miles(total_licencias), "Licencias totales", "📄")
kpi_card(c2, formato_decimal(promedio_dias_solicitados, 1), "Promedio días solicitados", "📅")
kpi_card(c3, formato_decimal(promedio_dias_autorizados, 1), "Promedio días autorizados", "✅")
kpi_card(c4, formato_pesos(promedio_subsidio_liquido), "Promedio subsidio líquido", "💰")

c5, c6 = st.columns(2)

kpi_card(c5, f"{tasa_autorizacion:.1f}%", "Tasa de autorización", "⚖️")
kpi_card(c6, f"{porcentaje_licencias_largas:.1f}%", "Licencias largas", "⏱️")

st.divider()


# ============================================================
# VISTA PREVIA DE LA BASE
# ============================================================

with st.expander("👁️ Ver vista previa de la base filtrada"):
    st.dataframe(df_filtrado.head(20), use_container_width=True)


# ============================================================
# MENÚ PRINCIPAL
# ============================================================

menu = st.selectbox(
    "Selecciona una sección",
    [
        "Resumen General",
        "Análisis por Isapre",
        "Diagnósticos",
        "Licencias Largas",
        "Trabajadores y Profesionales",
        "Exploración con PyGWalker",
    ]
)


# ============================================================
# SECCIÓN 1: RESUMEN GENERAL
# ============================================================

if menu == "Resumen General":

    st.header("📈 Resumen General")

    tab1, tab2, tab3 = st.tabs([
        "Evolución temporal",
        "Distribución por sexo",
        "Tipo de licencia"
    ])

    with tab1:
        resumen_periodo = (
            df_filtrado
            .groupby(["ANIO_EMISION", "TRIMESTRE_EMISION"], as_index=False)
            .size()
            .rename(columns={"size": "cantidad_licencias"})
        )

        resumen_periodo["PERIODO"] = (
            resumen_periodo["ANIO_EMISION"].astype(str).str.replace(".0", "", regex=False)
            + " - T"
            + resumen_periodo["TRIMESTRE_EMISION"].astype(str).str.replace(".0", "", regex=False)
        )

        fig = px.bar(
            resumen_periodo,
            x="PERIODO",
            y="cantidad_licencias",
            color="ANIO_EMISION",
            text="cantidad_licencias",
            title="Cantidad de licencias por año y trimestre",
            color_discrete_sequence=px.colors.qualitative.Set2
        )

        fig.update_traces(textposition="outside")

        fig.update_layout(
            xaxis_title="Periodo",
            yaxis_title="Cantidad de licencias",
            legend_title="Año"
        )

        fig = aplicar_layout(fig, 520)

        st.plotly_chart(fig, use_container_width=True)

        if len(resumen_periodo) > 0:
            periodo_max = resumen_periodo.sort_values(
                "cantidad_licencias",
                ascending=False
            ).iloc[0]

            info(
                f"🔍 <b>Info:</b> El periodo con mayor cantidad de licencias es "
                f"<b>{periodo_max['PERIODO']}</b>, con "
                f"<b>{formato_miles(periodo_max['cantidad_licencias'])}</b> registros."
            )

    with tab2:
        col_a, col_b = st.columns(2)

        with col_a:
            sexo_resumen = (
                df_filtrado["SEXO_TRABAJADOR"]
                .value_counts()
                .reset_index()
            )

            sexo_resumen.columns = ["Sexo", "Cantidad"]

            df_sexo_grafico = df_filtrado[
                df_filtrado["SEXO_TRABAJADOR"]
                .astype(str)
                .str.strip()
                .str.lower()
                .isin(["femenino", "masculino"])
            ]

            fig2 = px.pie(
                df_sexo_grafico,
                names="SEXO_TRABAJADOR",
                title="Distribución de licencias por sexo",
                color="SEXO_TRABAJADOR",
                color_discrete_sequence=["#f48fb1", "#4fc3f7"]
            )

            fig2.update_traces(
                textposition="inside",
                textinfo="percent",
                marker=dict(line=dict(color="#0d1117", width=1))
            )

            fig2.update_layout(
                legend_title="Sexo"
            )

            fig2 = aplicar_layout(fig2, 500)

            st.plotly_chart(fig2, use_container_width=True)

        with col_b:
            dias_sexo = (
                df_sexo_grafico
                .groupby("SEXO_TRABAJADOR", as_index=False)
                .agg(dias_promedio=("NUMERO_DE_DIAS_AUTORIZADOS", "mean"))
            )
            
 
            fig_dias_sexo = px.bar(
                dias_sexo,
                x="SEXO_TRABAJADOR",
                y="dias_promedio",
                color="SEXO_TRABAJADOR",
                text="dias_promedio",
                title="Días autorizados promedio por sexo",
                color_discrete_sequence=["#f48fb1", "#4fc3f7", "#9e9e9e"]
            )

            fig_dias_sexo.update_traces(
                texttemplate="%{y:.1f}",
                textposition="outside"
            )

            fig_dias_sexo.update_layout(
                showlegend=False,
                xaxis_title="Sexo",
                yaxis_title="Días promedio"
            )

            fig_dias_sexo = aplicar_layout(fig_dias_sexo, 500)

            st.plotly_chart(fig_dias_sexo, use_container_width=True)

        if len(sexo_resumen) > 0:
            sexo_top = sexo_resumen.iloc[0]

            info(
                f"👥 <b>Info:</b> El grupo con mayor cantidad de licencias es "
                f"<b>{sexo_top['Sexo']}</b>, con "
                f"<b>{formato_miles(sexo_top['Cantidad'])}</b> registros."
            )

    with tab3:
        if "TIPO_DE_LICENCIA" in df_filtrado.columns:
            tipo_resumen = (
                df_filtrado["TIPO_DE_LICENCIA"]
                .value_counts()
                .head(12)
                .reset_index()
            )

            tipo_resumen.columns = ["Tipo de licencia", "Cantidad"]

            fig_tipo = px.bar(
                tipo_resumen,
                x="Cantidad",
                y="Tipo de licencia",
                orientation="h",
                color="Cantidad",
                color_continuous_scale="Blues",
                text="Cantidad",
                title="Principales tipos de licencia"
            )

            fig_tipo.update_traces(textposition="outside")

            fig_tipo.update_layout(
                yaxis=dict(autorange="reversed"),
                xaxis_title="Cantidad",
                yaxis_title="Tipo de licencia",
                coloraxis_showscale=False
            )

            fig_tipo = aplicar_layout(fig_tipo, 560)

            st.plotly_chart(fig_tipo, use_container_width=True)

            if len(tipo_resumen) > 0:
                tipo_top = tipo_resumen.iloc[0]

                info(
                    f"📌 <b>Info:</b> El tipo de licencia más frecuente es "
                    f"<b>{tipo_top['Tipo de licencia']}</b>, con "
                    f"<b>{formato_miles(tipo_top['Cantidad'])}</b> registros."
                )
        else:
            warning_box("No existe la columna TIPO_DE_LICENCIA en la base.")


# ============================================================
# SECCIÓN 2: ANÁLISIS POR ISAPRE
# ============================================================

elif menu == "Análisis por Isapre":

    st.header("🏢 Análisis por Isapre")

    resumen_isapre = (
        df_filtrado
        .groupby("ISAPRE", as_index=False)
        .agg(
            cantidad_licencias=("ISAPRE", "count"),
            dias_autorizados=("NUMERO_DE_DIAS_AUTORIZADOS", "sum"),
            dias_promedio=("NUMERO_DE_DIAS_AUTORIZADOS", "mean"),
            subsidio_total=("MONTO_SUBSIDIO_LIQUIDO", "sum")
        )
        .sort_values("cantidad_licencias", ascending=False)
    )

    tab1, tab2, tab3 = st.tabs([
        "Licencias por Isapre",
        "Días autorizados",
        "Subsidio por Isapre"
    ])

    with tab1:
        fig = px.bar(
            resumen_isapre,
            x="ISAPRE",
            y="cantidad_licencias",
            color="cantidad_licencias",
            color_continuous_scale="Blues",
            text="cantidad_licencias",
            title="Cantidad de licencias por Isapre"
        )

        fig.update_traces(textposition="outside")

        fig.update_layout(
            xaxis_title="Isapre",
            yaxis_title="Cantidad de licencias",
            coloraxis_showscale=False,
            xaxis_tickangle=-25
        )

        fig = aplicar_layout(fig, 520)

        st.plotly_chart(fig, use_container_width=True)

        if len(resumen_isapre) > 0:
            top_isapre = resumen_isapre.iloc[0]

            info(
                f"🏢 <b>Info:</b> La Isapre con mayor cantidad de licencias es "
                f"<b>{top_isapre['ISAPRE']}</b>, con "
                f"<b>{formato_miles(top_isapre['cantidad_licencias'])}</b> registros."
            )

    with tab2:
        fig_dias = px.bar(
            resumen_isapre,
            x="ISAPRE",
            y="dias_promedio",
            color="dias_promedio",
            color_continuous_scale="Oranges",
            text="dias_promedio",
            title="Promedio de días autorizados por Isapre"
        )

        fig_dias.update_traces(
            texttemplate="%{y:.1f}",
            textposition="outside"
        )

        fig_dias.update_layout(
            xaxis_title="Isapre",
            yaxis_title="Días promedio",
            coloraxis_showscale=False,
            xaxis_tickangle=-25
        )

        fig_dias = aplicar_layout(fig_dias, 520)

        st.plotly_chart(fig_dias, use_container_width=True)

    with tab3:
        fig_subsidio = px.bar(
            resumen_isapre,
            x="ISAPRE",
            y="subsidio_total",
            color="subsidio_total",
            color_continuous_scale="Greens",
            text="subsidio_total",
            title="Subsidio líquido total por Isapre"
        )

        fig_subsidio.update_traces(
            texttemplate="$%{y:,.0f}",
            textposition="outside"
        )

        fig_subsidio.update_layout(
            xaxis_title="Isapre",
            yaxis_title="Subsidio líquido total",
            coloraxis_showscale=False,
            xaxis_tickangle=-25
        )

        fig_subsidio = aplicar_layout(fig_subsidio, 520)

        st.plotly_chart(fig_subsidio, use_container_width=True)

    st.subheader("Tabla resumen por Isapre")
    st.dataframe(resumen_isapre, use_container_width=True, hide_index=True)


# ============================================================
# SECCIÓN 3: DIAGNÓSTICOS
# ============================================================

elif menu == "Diagnósticos":

    st.header("🧠 Diagnósticos principales")

    if "SIGNIFICADO_CIE10" in df_filtrado.columns:
        columna_diag = "SIGNIFICADO_CIE10"
    elif "DIAGNOSTICO_PRINCIPAL" in df_filtrado.columns:
        columna_diag = "DIAGNOSTICO_PRINCIPAL"
    else:
        columna_diag = None

    if columna_diag is not None:

  

        df_diag_grafico = df_filtrado[
            df_filtrado[columna_diag]
            .astype(str)
            .str.strip()
            .str.lower()
            != "sin significado en diccionario externo"
        ]

        top_diagnosticos = (
            df_diag_grafico
            .groupby(columna_diag, as_index=False)
            .size()
            .rename(columns={"size": "cantidad"})
            .sort_values("cantidad", ascending=False)
            .head(12)
        )

        top_diagnosticos["porcentaje"] = (
            top_diagnosticos["cantidad"] / len(df_diag_grafico) * 100
        ).round(1)

        fig = px.bar(
            top_diagnosticos,
            x="cantidad",
            y=columna_diag,
            orientation="h",
            color="porcentaje",
            color_continuous_scale="Blues",
            text=top_diagnosticos["porcentaje"].apply(lambda x: f"{x}%"),
            title="Top 12 diagnósticos más frecuentes"
        )

        fig.update_traces(textposition="outside")

        fig.update_layout(
            yaxis=dict(autorange="reversed"),
            xaxis_title="Cantidad",
            yaxis_title="Diagnóstico",
            coloraxis_showscale=False
        )

        fig = aplicar_layout(fig, 620)

        st.plotly_chart(fig, use_container_width=True)

        st.divider()

        col_a, col_b = st.columns(2)

        with col_a:
            dias_diag = (
                df_filtrado
                .groupby(columna_diag, as_index=False)
                .agg(dias_promedio=("NUMERO_DE_DIAS_AUTORIZADOS", "mean"))
                .dropna()
                .sort_values("dias_promedio", ascending=False)
                .head(10)
            )

            fig_dias = px.bar(
                dias_diag,
                x="dias_promedio",
                y=columna_diag,
                orientation="h",
                color="dias_promedio",
                color_continuous_scale="Oranges",
                text="dias_promedio",
                title="Diagnósticos con mayor duración promedio"
            )

            fig_dias.update_traces(
                texttemplate="%{x:.1f}",
                textposition="outside"
            )

            fig_dias.update_layout(
                yaxis=dict(autorange="reversed"),
                xaxis_title="Días promedio",
                yaxis_title="Diagnóstico",
                coloraxis_showscale=False
            )

            fig_dias = aplicar_layout(fig_dias, 520)

            st.plotly_chart(fig_dias, use_container_width=True)

        with col_b:
            top8 = top_diagnosticos[columna_diag].head(8).tolist()

            df_top = df_filtrado[
                df_filtrado[columna_diag].isin(top8)
            ]

            cruce = (
                df_top
                .groupby([columna_diag, "SEXO_TRABAJADOR"], as_index=False)
                .size()
                .rename(columns={"size": "cantidad"})
            )

            fig_sexo_diag = px.bar(
                cruce,
                x="cantidad",
                y=columna_diag,
                color="SEXO_TRABAJADOR",
                orientation="h",
                barmode="group",
                title="Top diagnósticos por sexo",
                color_discrete_sequence=["#f48fb1", "#4fc3f7", "#9e9e9e"]
            )

            fig_sexo_diag.update_layout(
                yaxis=dict(autorange="reversed"),
                xaxis_title="Cantidad",
                yaxis_title="Diagnóstico",
                legend_title="Sexo"
            )

            fig_sexo_diag = aplicar_layout(fig_sexo_diag, 520)

            st.plotly_chart(fig_sexo_diag, use_container_width=True)

        if len(top_diagnosticos) > 0:
            diag_top = top_diagnosticos.iloc[0]

            info(
                f"🧠 <b>Info:</b> El diagnóstico más frecuente es "
                f"<b>{diag_top[columna_diag]}</b>, con "
                f"<b>{formato_miles(diag_top['cantidad'])}</b> licencias, "
                f"equivalente al <b>{diag_top['porcentaje']}%</b> de la base filtrada."
            )

    else:
        warning_box("No existe SIGNIFICADO_CIE10 ni DIAGNOSTICO_PRINCIPAL en la base.")


# ============================================================
# SECCIÓN 4: LICENCIAS LARGAS
# ============================================================

elif menu == "Licencias Largas":

    st.header("⏱️ Análisis de Licencias Largas")

    resumen_larga = (
        df_filtrado
        .groupby("LICENCIA_LARGA", as_index=False)
        .size()
        .rename(columns={"size": "cantidad"})
    )

    col_a, col_b = st.columns(2)

    with col_a:
        fig = px.bar(
            resumen_larga,
            x="LICENCIA_LARGA",
            y="cantidad",
            color="LICENCIA_LARGA",
            text="cantidad",
            title="Cantidad de licencias largas y no largas",
            color_discrete_sequence=["#43a047", "#f9a825", "#e53935"]
        )

        fig.update_traces(textposition="outside")

        fig.update_layout(
            xaxis_title="Licencia larga",
            yaxis_title="Cantidad",
            showlegend=False
        )

        fig = aplicar_layout(fig, 500)

        st.plotly_chart(fig, use_container_width=True)

    with col_b:
        df_larga_sexo_grafico = df_filtrado[
            df_filtrado["SEXO_TRABAJADOR"]
            .astype(str)
            .str.strip()
            .str.lower()
            .isin(["femenino", "masculino"])
        ]

        larga_sexo = (
            df_larga_sexo_grafico
            .groupby(["LICENCIA_LARGA", "SEXO_TRABAJADOR"], as_index=False)
            .size()
            .rename(columns={"size": "cantidad"})
        )

        fig_sexo = px.bar(
            larga_sexo,
            x="SEXO_TRABAJADOR",
            y="cantidad",
            color="LICENCIA_LARGA",
            barmode="group",
            text="cantidad",
            title="Licencias largas por sexo",
            color_discrete_sequence=px.colors.qualitative.Set2
        )

        fig_sexo.update_traces(textposition="outside")

        fig_sexo.update_layout(
            xaxis_title="Sexo",
            yaxis_title="Cantidad",
            legend_title="Licencia larga"
        )

        fig_sexo = aplicar_layout(fig_sexo, 500)

        st.plotly_chart(fig_sexo, use_container_width=True)

    st.subheader("Tabla resumen")
    st.dataframe(resumen_larga, use_container_width=True, hide_index=True)

    info(
        f"⏱️ <b>Info:</b> La base filtrada contiene "
        f"<b>{formato_miles(cantidad_largas)}</b> licencias largas."
    )


# ============================================================
# SECCIÓN 5: TRABAJADORES Y PROFESIONALES
# ============================================================

elif menu == "Trabajadores y Profesionales":

    st.header("👨‍⚕️ Trabajadores y Profesionales")

    st.subheader("Top 10 profesionales con más licencias emitidas")

    if "RUN_PROFESIONAL" in df_filtrado.columns:
        profesional = (
            df_filtrado
            .groupby("RUN_PROFESIONAL", as_index=False)
            .size()
            .rename(columns={"size": "cantidad_licencias"})
            .sort_values("cantidad_licencias", ascending=False)
            .head(10)
        )

        st.dataframe(
            profesional,
            use_container_width=True,
            hide_index=True
        )
    else:
        warning_box("No existe la columna RUN_PROFESIONAL.")

    st.divider()

    st.subheader("Licencias por calidad del trabajador")

    if "CALIDAD_TRABAJADOR" in df_filtrado.columns:
        calidad = (
            df_filtrado["CALIDAD_TRABAJADOR"]
            .value_counts()
            .head(12)
            .reset_index()
        )

        calidad.columns = ["Calidad trabajador", "Cantidad"]

        fig_calidad = px.bar(
            calidad,
            x="Cantidad",
            y="Calidad trabajador",
            orientation="h",
            color="Cantidad",
            color_continuous_scale="Purples",
            text="Cantidad",
            title="Licencias por calidad del trabajador"
        )

        fig_calidad.update_traces(textposition="outside")

        fig_calidad.update_layout(
            yaxis=dict(autorange="reversed"),
            xaxis_title="Cantidad",
            yaxis_title="Calidad trabajador",
            coloraxis_showscale=False,
            height=520
        )

        fig_calidad = aplicar_layout(fig_calidad, 520)

        st.plotly_chart(fig_calidad, use_container_width=True)

        if len(calidad) > 0:
            calidad_top = calidad.iloc[0]
            porcentaje_calidad_top = calidad_top["Cantidad"] / len(df_filtrado) * 100

            info(
                f"💼 <b>Info:</b> La calidad de trabajador con más licencias es "
                f"<b>{calidad_top['Calidad trabajador']}</b>, con "
                f"<b>{formato_miles(calidad_top['Cantidad'])}</b> licencias, "
                f"equivalente al <b>{porcentaje_calidad_top:.1f}%</b> de la base filtrada."
            )
        
        
    else:
        warning_box("No existe la columna CALIDAD_TRABAJADOR.")


# ============================================================
# SECCIÓN 6: EXPLORACIÓN CON PYGWALKER
# ============================================================

elif menu == "Exploración con PyGWalker":

    st.header("🧭 Exploración visual con PyGWalker")

    if PYGWALKER_OK:
        st.write(
            """
            Esta sección permite explorar la base filtrada de forma visual.
            Para evitar que la app se ponga lenta, se usa una muestra de registros.
            """
        )

        max_filas = min(len(df_filtrado), 5000)

        if max_filas > 0:
            muestra = df_filtrado.head(max_filas)

            generated_html = pyg.to_html(
                muestra,
                return_html=True,
                dark="light"
            )

            components.html(
                generated_html,
                height=800,
                scrolling=True
            )
        else:
            st.warning("No hay datos disponibles con los filtros actuales.")
    else:
        warning_box("PyGWalker no está instalado. Puedes instalarlo con: pip install pygwalker")




# ============================================================
# FOOTER
# ============================================================

st.sidebar.divider()
st.sidebar.markdown("Proyecto Licencias Médicas")
st.sidebar.markdown("Ingeniería en Información y Control de Gestión")
st.sidebar.markdown("Universidad Santo Tomás")
