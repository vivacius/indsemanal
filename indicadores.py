import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO

# ==========================================
# CONFIGURACIÓN GLOBAL DE ESTILO
# ==========================================
plt.rcParams.update({
    'font.family': 'sans-serif',
    'font.size': 11,
    'axes.titlesize': 13,
    'axes.labelsize': 11,
    'xtick.labelsize': 10,
    'ytick.labelsize': 10,
    'figure.titlesize': 14,
    'figure.dpi': 120
})

# Configuración de Streamlit
st.set_page_config(
    page_title="Dashboard Ejecutivo - Productividad de Equipos",
    page_icon="📊",
    layout="wide"
)

# Estilo CSS personalizado
st.markdown("""
    <style>
    .block-container { padding: 2rem 3rem; }
    h1, h2, h3 { color: #2c3e50; font-weight: 700; }
    h3 { border-left: 4px solid #3498db; padding-left: 10px; margin-top: 1.5rem; }
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; background-color: #f8f9fa; border-radius: 6px 6px 0 0; }
    .stTabs [aria-selected="true"] { background-color: #2c3e50; color: white; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Dashboard Ejecutivo: Productividad de Equipos")
st.markdown("<small style='color: #7f8c8d;'>Análisis comparativo con y sin operación de cierre. Datos actualizados en tiempo real.</small>", unsafe_allow_html=True)

# --- Carga de archivo ---
uploaded_file = st.file_uploader("📂 Sube tu archivo Excel", type=["xlsx"], label_visibility="collapsed")

if uploaded_file is not None:
    # Cargar datos
    df = pd.read_excel(uploaded_file)

    # --- Crear columna 'Frente' ---
    df["Frente"] = df["Descripción del grupo de equipos"].apply(
        lambda x: "TRACTOMULA" if x == "FRENTE TRACTOR MULA - CAMPO" else "TRACTORES"
    )

    # --- Calcular duración en horas ---
    df["Duracion_h"] = (
        pd.to_datetime(df["Hora de finalización"]) - pd.to_datetime(df["Hora de inicio"])
    ).dt.total_seconds() / 3600

    # --- Normalizar GOP ---
    df["GOP"] = df["Descripción del grupo de operaciones"].apply(
        lambda x: "PRODUCTIVO" if str(x).upper() == "AUXILIAR" else x
    )

    # --- Crear dos versiones de los datos ---
    df_con_fin = df.copy()  # Con FIN DE OPERACION
    df_sin_fin = df[df["Descripción de la operación"] != "FIN DE OPERACION DE LA MAQUINA"].copy()  # Sin FIN DE OPERACION

    # --- Crear pestañas ---
    tab1, tab2 = st.tabs(["📈 CON Fin de Operación", "📉 SIN Fin de Operación"])

    # ===================================================
    # PESTAÑA 1: CON FIN DE OPERACION
    # ===================================================
    with tab1:
        st.header("📈 Análisis CON 'FIN DE OPERACION DE LA MAQUINA'")
        df_actual = df_con_fin.copy()

        # --- Gráfico 1: Equipos con >60% en Pérdida o Mantenimiento (DISEÑO GERENCIAL) ---
        st.subheader("1. Equipos con más del 60% de su tiempo en Pérdida o Mantenimiento")

        # Agrupar por equipo y frente, sumando horas por GOP
        horas_por_gop = df_actual.groupby(["Frente", "Código de equipo", "GOP"])["Duracion_h"].sum().reset_index()
        totales_por_equipo = df_actual.groupby(["Frente", "Código de equipo"])["Duracion_h"].sum().reset_index(name="Total_h")

        # Unir para calcular %
        horas_por_gop = horas_por_gop.merge(totales_por_equipo, on=["Frente", "Código de equipo"])
        horas_por_gop["%_GOP"] = (horas_por_gop["Duracion_h"] / horas_por_gop["Total_h"]) * 100

        # Filtrar equipos con >60% en Pérdida o Mantenimiento
        equipos_60 = horas_por_gop[
            (horas_por_gop["GOP"].isin(["PERDIDA", "MANTENIMIENTO"])) &
            (horas_por_gop["%_GOP"] > 60)
        ]

        # Contar equipos por Frente y GOP
        resumen = equipos_60.groupby(["Frente", "GOP"])["Código de equipo"].count().reset_index(name="Cantidad_Equipos")

        # Obtener lista de frentes únicos
        frentes_unicos = resumen["Frente"].unique()

        # Crear un gráfico por frente
        for frente in frentes_unicos:
            subset_frente = resumen[resumen["Frente"] == frente]

            # Configurar figura
            fig, ax = plt.subplots(figsize=(8, 2))
            colores = {"PERDIDA": "#e74c3c", "MANTENIMIENTO": "#3498db"}

            # Barras
            for gop in subset_frente["GOP"].unique():
                valor = subset_frente[subset_frente["GOP"] == gop]["Cantidad_Equipos"].iloc[0]
                barra = ax.bar(
                    gop,
                    valor,
                    color=colores[gop],
                    edgecolor='white',
                    linewidth=1.5,
                    width=0.4
                )
                # Etiqueta encima de la barra
                ax.text(
                    barra[0].get_x() + barra[0].get_width() / 2,
                    valor + 0.1,
                    f'{int(valor)}',
                    ha='center',
                    va='bottom',
                    fontsize=12,
                    fontweight='bold',
                    color=colores[gop]
                )

            # Estilo minimalista
            ax.set_title(f"{frente.upper()}", fontsize=14, fontweight='bold', pad=15, loc='left', color='#2c3e50')
            ax.set_ylabel("Cantidad de Equipos", fontsize=9, color='#7f8c8d')
            ax.set_xlabel("")
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_color('#bdc3c7')
            ax.spines['bottom'].set_color('#bdc3c7')
            ax.grid(axis='y', linestyle='--', alpha=0.5, color='#ecf0f1')
            ax.set_axisbelow(True)
            ax.set_ylim(0, max(subset_frente["Cantidad_Equipos"]) * 1.2 if not subset_frente.empty else 1)

            # KPI resaltado debajo del gráfico
            total_equipos = subset_frente["Cantidad_Equipos"].sum()
            st.markdown(f"<div style='text-align: center; font-size: 16px; font-weight: bold; color: #2c3e50; padding: 12px; "
                        f"background-color: #f8f9fa; border-left: 4px solid #e67e22; border-radius: 0 8px 8px 0; margin: 10px 0;'>"
                        f"⚠️ Equipos en riesgo operativo: <span style='color: #e67e22; font-size: 18px;'>{int(total_equipos)}</span></div>",
                        unsafe_allow_html=True)

            # Mostrar gráfico
            st.pyplot(fig)
            st.markdown("---")

        # --- Gráfico 2: Top 5 actividades en pérdida (COLOR ROJO) ---
        st.subheader("2. Top 5 Actividades en Pérdida (% horas)")
        df_perdida = df_actual[df_actual["GOP"] == "PERDIDA"]
        resumen_perdida = (
            df_perdida.groupby(["Frente", "Descripción de la operación"])["Duracion_h"]
            .sum()
            .reset_index()
            .rename(columns={"Descripción de la operación": "Actividad", "Duracion_h": "Horas"})
        )
        resumen_perdida["%"] = resumen_perdida.groupby("Frente")["Horas"].transform(lambda x: 100 * x / x.sum())
        top5 = resumen_perdida.groupby("Frente").apply(
            lambda x: x.sort_values("%", ascending=False).head(5)
        ).reset_index(drop=True)

        frentes = top5["Frente"].unique()
        n_frentes = len(frentes)
        fig_height = max(5, n_frentes * 3.5)
        fig, axes = plt.subplots(n_frentes, 1, figsize=(12, fig_height))
        if n_frentes == 1:
            axes = [axes]

        for ax, frente in zip(axes, frentes):
            subset = top5[top5["Frente"] == frente].copy()
            bars = ax.barh(
                subset["Actividad"],
                subset["%"],
                color="#e74c3c",  # 🔴 ROJO
                edgecolor='white',
                linewidth=1.0,
                height=0.6
            )
            ax.set_title(f"{frente.upper()}", fontsize=14, fontweight='bold', pad=15, loc='left', color='#2c3e50')
            ax.set_xlabel("% de horas perdidas", fontsize=11, color='#7f8c8d')
            ax.set_ylabel("Actividad", fontsize=11, color='#7f8c8d')
            ax.invert_yaxis()
            ax.grid(axis='x', linestyle='--', alpha=0.6, linewidth=0.8, color='#ecf0f1')
            ax.set_axisbelow(True)
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)

            # Etiquetas de valor
            for bar in bars:
                width = bar.get_width()
                ax.text(
                    width + 0.8,
                    bar.get_y() + bar.get_height()/2,
                    f'{width:.1f}%',
                    va='center',
                    fontsize=10,
                    fontweight='bold',
                    color='#c0392b'
                )

        plt.tight_layout(pad=3.0)
        st.pyplot(fig)

        # --- Tabla de Rangos de Productividad ---
        st.subheader("3. Distribución por Rango de Productividad")
        totales_equipo = df_actual.groupby(["Frente", "Código de equipo"], as_index=False)["Duracion_h"].sum().rename(columns={"Duracion_h":"Total_h"})
        horas_equipo = df_actual.groupby(["Frente","Código de equipo","GOP"], as_index=False)["Duracion_h"].sum()
        horas_equipo = horas_equipo.merge(totales_equipo, on=["Frente","Código de equipo"])
        horas_equipo["%_GOP"] = (horas_equipo["Duracion_h"] / horas_equipo["Total_h"]) * 100

        bins = [0,10,20,30,40,50,60,100]
        labels = ["0-10%", "10-20%","20-30%","30-40%","40-50%", "50-60%", "60-100%"]
        horas_equipo["Rango_Productividad"] = pd.cut(horas_equipo["%_GOP"], bins=bins, labels=labels, include_lowest=True)

        tabla_rangos = horas_equipo.groupby(["Frente","GOP","Rango_Productividad"]).agg(
            Equipos=("Código de equipo","nunique"),
            Promedio=("%_GOP","mean")
        ).reset_index()

        # Mostrar tabla por frente con estilo
        for frente in tabla_rangos["Frente"].unique():
            tabla = (
                tabla_rangos[tabla_rangos["Frente"]==frente]
                .pivot(index="GOP", columns="Rango_Productividad", values="Equipos")
                .fillna(0)
                .astype(int)
            )
            st.markdown(f"### {frente.upper()}")
            st.dataframe(
                tabla.style
                .background_gradient(cmap="Blues", axis=None)
                .format("{:.0f}")
                .set_properties(**{'text-align': 'center', 'font-weight': 'bold'})
                .set_table_styles([
                    {'selector': 'th', 'props': [('background-color', '#f1f3f6'), ('color', '#2c3e50')]},
                    {'selector': 'td', 'props': [('border', '1px solid #e0e0e0')]}
                ])
            )

        # --- Gráfico 4: Distribución PRODUCTIVOS (COLOR VERDE) ---
        st.subheader("4. Equipos Productivos por Rango de Eficiencia")
        df_prod = tabla_rangos[tabla_rangos["GOP"] == "PRODUCTIVO"]
        orden = ["0-10%", "10-20%","20-30%","30-40%","40-50%", "50-60%", "60-100%"]
        df_prod["Rango_Productividad"] = pd.Categorical(df_prod["Rango_Productividad"], categories=orden, ordered=True)

        g = sns.catplot(
            data=df_prod,
            x="Rango_Productividad",
            y="Equipos",
            kind="bar",
            col="Frente",
            col_wrap=2,
            height=5,
            aspect=1.2,
            palette=["#27ae60"] * len(df_prod),  # 🟩 VERDE
            edgecolor='white',
            linewidth=1.2
        )
        g.set_titles("{col_name}", size=13, weight='bold', color='#2c3e50')
        g.set_axis_labels("Rango de Productividad", "Número de Equipos", size=11, color='#7f8c8d')
        g.set_xticklabels(rotation=45, size=10)

        # Agregar etiquetas en cada subplot
        for ax in g.axes.flat:
            for bar in ax.patches:
                height = bar.get_height()
                if height > 0:
                    ax.text(
                        bar.get_x() + bar.get_width() / 2,
                        height + 0.2,
                        f'{int(height)}',
                        ha='center',
                        va='bottom',
                        fontsize=10,
                        fontweight='bold',
                        color='#2c3e50'
                    )
            ax.grid(axis='y', linestyle='--', alpha=0.6, linewidth=0.8, color='#ecf0f1')
            ax.set_axisbelow(True)
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)

        plt.tight_layout()
        st.pyplot(g)

    # ===================================================
    # PESTAÑA 2: SIN FIN DE OPERACION
    # ===================================================
    with tab2:
        st.header("📉 Análisis SIN 'FIN DE OPERACION DE LA MAQUINA'")
        df_actual = df_sin_fin.copy()

        # --- Gráfico 1: Equipos con >60% en Pérdida o Mantenimiento (DISEÑO GERENCIAL) ---
        st.subheader("1. Equipos con más del 60% de su tiempo en Pérdida o Mantenimiento")

        # Agrupar por equipo y frente, sumando horas por GOP
        horas_por_gop = df_actual.groupby(["Frente", "Código de equipo", "GOP"])["Duracion_h"].sum().reset_index()
        totales_por_equipo = df_actual.groupby(["Frente", "Código de equipo"])["Duracion_h"].sum().reset_index(name="Total_h")

        # Unir para calcular %
        horas_por_gop = horas_por_gop.merge(totales_por_equipo, on=["Frente", "Código de equipo"])
        horas_por_gop["%_GOP"] = (horas_por_gop["Duracion_h"] / horas_por_gop["Total_h"]) * 100

        # Filtrar equipos con >60% en Pérdida o Mantenimiento
        equipos_60 = horas_por_gop[
            (horas_por_gop["GOP"].isin(["PERDIDA", "MANTENIMIENTO"])) &
            (horas_por_gop["%_GOP"] > 60)
        ]

        # Contar equipos por Frente y GOP
        resumen = equipos_60.groupby(["Frente", "GOP"])["Código de equipo"].count().reset_index(name="Cantidad_Equipos")

        # Obtener lista de frentes únicos
        frentes_unicos = resumen["Frente"].unique()

        # Crear un gráfico por frente
        for frente in frentes_unicos:
            subset_frente = resumen[resumen["Frente"] == frente]

            # Configurar figura
            fig, ax = plt.subplots(figsize=(8, 2))
            colores = {"PERDIDA": "#e74c3c", "MANTENIMIENTO": "#3498db"}

            # Barras
            for gop in subset_frente["GOP"].unique():
                valor = subset_frente[subset_frente["GOP"] == gop]["Cantidad_Equipos"].iloc[0]
                barra = ax.bar(
                    gop,
                    valor,
                    color=colores[gop],
                    edgecolor='white',
                    linewidth=1.5,
                    width=0.6
                )
                # Etiqueta encima de la barra
                ax.text(
                    barra[0].get_x() + barra[0].get_width() / 2,
                    valor + 0.1,
                    f'{int(valor)}',
                    ha='center',
                    va='bottom',
                    fontsize=12,
                    fontweight='bold',
                    color=colores[gop]
                )

            # Estilo minimalista
            ax.set_title(f"{frente.upper()}", fontsize=14, fontweight='bold', pad=15, loc='left', color='#2c3e50')
            ax.set_ylabel("Cantidad de Equipos", fontsize=9, color='#7f8c8d')
            ax.set_xlabel("")
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_color('#bdc3c7')
            ax.spines['bottom'].set_color('#bdc3c7')
            ax.grid(axis='y', linestyle='--', alpha=0.5, color='#ecf0f1')
            ax.set_axisbelow(True)
            ax.set_ylim(0, max(subset_frente["Cantidad_Equipos"]) * 1.2 if not subset_frente.empty else 1)

            # KPI resaltado debajo del gráfico
            total_equipos = subset_frente["Cantidad_Equipos"].sum()
            st.markdown(f"<div style='text-align: center; font-size: 16px; font-weight: bold; color: #2c3e50; padding: 12px; "
                        f"background-color: #f8f9fa; border-left: 4px solid #e67e22; border-radius: 0 8px 8px 0; margin: 10px 0;'>"
                        f"⚠️ Equipos en riesgo operativo: <span style='color: #e67e22; font-size: 18px;'>{int(total_equipos)}</span></div>",
                        unsafe_allow_html=True)

            # Mostrar gráfico
            st.pyplot(fig)
            st.markdown("---")

        # --- Gráfico 2: Top 5 actividades en pérdida (COLOR ROJO) ---
        st.subheader("2. Top 5 Actividades en Pérdida (% horas)")
        df_perdida = df_actual[df_actual["GOP"] == "PERDIDA"]
        resumen_perdida = (
            df_perdida.groupby(["Frente", "Descripción de la operación"])["Duracion_h"]
            .sum()
            .reset_index()
            .rename(columns={"Descripción de la operación": "Actividad", "Duracion_h": "Horas"})
        )
        resumen_perdida["%"] = resumen_perdida.groupby("Frente")["Horas"].transform(lambda x: 100 * x / x.sum())
        top5 = resumen_perdida.groupby("Frente").apply(
            lambda x: x.sort_values("%", ascending=False).head(5)
        ).reset_index(drop=True)

        frentes = top5["Frente"].unique()
        n_frentes = len(frentes)
        fig_height = max(5, n_frentes * 3.5)
        fig, axes = plt.subplots(n_frentes, 1, figsize=(12, fig_height))
        if n_frentes == 1:
            axes = [axes]

        for ax, frente in zip(axes, frentes):
            subset = top5[top5["Frente"] == frente].copy()
            bars = ax.barh(
                subset["Actividad"],
                subset["%"],
                color="#e74c3c",  # 🔴 ROJO
                edgecolor='white',
                linewidth=1.0,
                height=0.6
            )
            ax.set_title(f"{frente.upper()}", fontsize=14, fontweight='bold', pad=15, loc='left', color='#2c3e50')
            ax.set_xlabel("% de horas perdidas", fontsize=11, color='#7f8c8d')
            ax.set_ylabel("Actividad", fontsize=11, color='#7f8c8d')
            ax.invert_yaxis()
            ax.grid(axis='x', linestyle='--', alpha=0.6, linewidth=0.8, color='#ecf0f1')
            ax.set_axisbelow(True)
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)

            # Etiquetas de valor
            for bar in bars:
                width = bar.get_width()
                ax.text(
                    width + 0.8,
                    bar.get_y() + bar.get_height()/2,
                    f'{width:.1f}%',
                    va='center',
                    fontsize=10,
                    fontweight='bold',
                    color='#c0392b'
                )

        plt.tight_layout(pad=3.0)
        st.pyplot(fig)

        # --- Tabla de Rangos de Productividad ---
        st.subheader("3. Distribución por Rango de Productividad")
        totales_equipo = df_actual.groupby(["Frente", "Código de equipo"], as_index=False)["Duracion_h"].sum().rename(columns={"Duracion_h":"Total_h"})
        horas_equipo = df_actual.groupby(["Frente","Código de equipo","GOP"], as_index=False)["Duracion_h"].sum()
        horas_equipo = horas_equipo.merge(totales_equipo, on=["Frente","Código de equipo"])
        horas_equipo["%_GOP"] = (horas_equipo["Duracion_h"] / horas_equipo["Total_h"]) * 100

        bins = [0,10,20,30,40,50,60,100]
        labels = ["0-10%", "10-20%","20-30%","30-40%","40-50%", "50-60%", "60-100%"]
        horas_equipo["Rango_Productividad"] = pd.cut(horas_equipo["%_GOP"], bins=bins, labels=labels, include_lowest=True)

        tabla_rangos = horas_equipo.groupby(["Frente","GOP","Rango_Productividad"]).agg(
            Equipos=("Código de equipo","nunique"),
            Promedio=("%_GOP","mean")
        ).reset_index()

        # Mostrar tabla por frente con estilo
        for frente in tabla_rangos["Frente"].unique():
            tabla = (
                tabla_rangos[tabla_rangos["Frente"]==frente]
                .pivot(index="GOP", columns="Rango_Productividad", values="Equipos")
                .fillna(0)
                .astype(int)
            )
            st.markdown(f"### {frente.upper()}")
            st.dataframe(
                tabla.style
                .background_gradient(cmap="Blues", axis=None)
                .format("{:.0f}")
                .set_properties(**{'text-align': 'center', 'font-weight': 'bold'})
                .set_table_styles([
                    {'selector': 'th', 'props': [('background-color', '#f1f3f6'), ('color', '#2c3e50')]},
                    {'selector': 'td', 'props': [('border', '1px solid #e0e0e0')]}
                ])
            )

        # --- Gráfico 4: Distribución PRODUCTIVOS (COLOR VERDE) ---
        st.subheader("4. Equipos Productivos por Rango de Eficiencia")
        df_prod = tabla_rangos[tabla_rangos["GOP"] == "PRODUCTIVO"]
        orden = ["0-10%", "10-20%","20-30%","30-40%","40-50%", "50-60%", "60-100%"]
        df_prod["Rango_Productividad"] = pd.Categorical(df_prod["Rango_Productividad"], categories=orden, ordered=True)

        g = sns.catplot(
            data=df_prod,
            x="Rango_Productividad",
            y="Equipos",
            kind="bar",
            col="Frente",
            col_wrap=2,
            height=5,
            aspect=1.2,
            palette=["#27ae60"] * len(df_prod),  # 🟩 VERDE
            edgecolor='white',
            linewidth=1.2
        )
        g.set_titles("{col_name}", size=13, weight='bold', color='#2c3e50')
        g.set_axis_labels("Rango de Productividad", "Número de Equipos", size=11, color='#7f8c8d')
        g.set_xticklabels(rotation=45, size=10)

        # Agregar etiquetas en cada subplot
        for ax in g.axes.flat:
            for bar in ax.patches:
                height = bar.get_height()
                if height > 0:
                    ax.text(
                        bar.get_x() + bar.get_width() / 2,
                        height + 0.2,
                        f'{int(height)}',
                        ha='center',
                        va='bottom',
                        fontsize=10,
                        fontweight='bold',
                        color='#2c3e50'
                    )
            ax.grid(axis='y', linestyle='--', alpha=0.6, linewidth=0.8, color='#ecf0f1')
            ax.set_axisbelow(True)
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)

        plt.tight_layout()
        st.pyplot(g)

else:
    st.info("👆 Por favor, sube un archivo Excel para comenzar el análisis ejecutivo.")
    st.image("https://www.manuelita.com/wp-content/uploads/2017/01/Logo-manuelita.jpg", width=400)