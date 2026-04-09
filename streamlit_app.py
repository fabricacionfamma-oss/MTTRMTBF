import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from fpdf import FPDF
import tempfile
import os

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA
# ==========================================
st.set_page_config(page_title="KPIs Mantenimiento - Matricería", layout="wide", page_icon="⚙️")

st.markdown("""
<style>
    .metric-card {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #1f77b4;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.1);
    }
</style>
""", unsafe_allow_html=True)

col_title, col_btn = st.columns([4, 1])
with col_title:
    st.title("⚙️ Análisis de MTBF y MTTR - Matricería")
    st.write("Indicadores de Confiabilidad y Mantenibilidad con seguimiento Mensual (Formato PD).")
with col_btn:
    if st.button("Limpiar Caché", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

st.divider()

# ==========================================
# OBJETIVOS (TARGETS) SEGÚN EXCEL PD
# ==========================================
TARGET_MTTR_MIN = 30      # 30 minutos (0.5 hs)
TARGET_MTBF_HS = 500      # 500 horas

# ==========================================
# FILTROS
# ==========================================
anio_actual = pd.to_datetime("today").year
anio_sel = st.selectbox("Seleccione el Año para el Análisis y Reporte:", range(2023, anio_actual + 2), index=anio_actual-2023)

# ==========================================
# EXTRACCIÓN Y PROCESAMIENTO DE DATOS ANUAL
# ==========================================
@st.cache_data(ttl=300)
def fetch_annual_data(anio):
    try:
        conn = st.connection("wii_bi", type="sql")
        
        # 1. Uptime Agrupado por Mes
        q_uptime = f"""
            SELECT MONTH(p.Date) as Mes, 
                   SUM(p.ProductiveTime) as Tiempo_Productivo_Min
            FROM PROD_D_03 p
            JOIN CELL c ON p.CellId = c.CellId
            WHERE YEAR(p.Date) = {anio}
            GROUP BY MONTH(p.Date)
        """
        df_uptime = conn.query(q_uptime)
        
        # 2. Fallas Matricería Agrupadas por Mes
        q_fallas = f"""
            SELECT MONTH(e.Date) as Mes, 
                   COUNT(e.Id) as Cantidad_Fallas,
                   SUM(e.Interval) as Tiempo_Reparacion_Min
            FROM EVENT_01 e
            LEFT JOIN EVENTTYPE t1 ON e.EventTypeLevel1 = t1.EventTypeId
            LEFT JOIN EVENTTYPE t2 ON e.EventTypeLevel2 = t2.EventTypeId
            LEFT JOIN EVENTTYPE t3 ON e.EventTypeLevel3 = t3.EventTypeId
            LEFT JOIN EVENTTYPE t4 ON e.EventTypeLevel4 = t4.EventTypeId
            WHERE YEAR(e.Date) = {anio}
              AND (
                  UPPER(t1.Name) LIKE '%MATRI%' OR UPPER(t2.Name) LIKE '%MATRI%' OR UPPER(t3.Name) LIKE '%MATRI%' OR UPPER(t4.Name) LIKE '%MATRI%'
                  OR UPPER(t1.Name) LIKE '%HERRAMENTAL%' OR UPPER(t2.Name) LIKE '%HERRAMENTAL%'
              )
            GROUP BY MONTH(e.Date)
        """
        df_fallas = conn.query(q_fallas)
        
        # Generar estructura de 12 meses
        df_meses = pd.DataFrame({'Mes': range(1, 13)})
        
        # Unir datos
        df_anual = pd.merge(df_meses, df_uptime, on='Mes', how='left')
        df_anual = pd.merge(df_anual, df_fallas, on='Mes', how='left').fillna(0)
        
        # Cálculos en Horas
        df_anual['Uptime_Hs'] = df_anual['Tiempo_Productivo_Min'] / 60.0
        df_anual['Downtime_Hs'] = df_anual['Tiempo_Reparacion_Min'] / 60.0
        df_anual['Downtime_Min'] = df_anual['Tiempo_Reparacion_Min']
        
        df_anual['MTBF (Hs)'] = df_anual.apply(
            lambda r: r['Uptime_Hs'] / r['Cantidad_Fallas'] if r['Cantidad_Fallas'] > 0 else (r['Uptime_Hs'] if r['Uptime_Hs'] > 0 else 0), axis=1
        )
        
        # Calculamos MTTR en Minutos para que coincida con el objetivo del Excel (30 min)
        df_anual['MTTR (Min)'] = df_anual.apply(
            lambda r: r['Downtime_Min'] / r['Cantidad_Fallas'] if r['Cantidad_Fallas'] > 0 else 0, axis=1
        )
        
        return df_anual
    except Exception as e:
        st.error(f"Error consultando BD: {e}")
        return pd.DataFrame()

df_anual = fetch_annual_data(anio_sel)

# ==========================================
# GENERADOR PDF (FORMATO PESTAÑA 'PD')
# ==========================================
class ReportePD(FPDF):
    def header(self):
        self.set_font("Arial", 'B', 14)
        self.set_text_color(15, 76, 129)
        self.cell(0, 10, f"Reporte de Indicadores de Mantenimiento Matrices (PD) - Año {anio_sel}", ln=True, align='C')
        self.set_draw_color(15, 76, 129)
        self.set_line_width(0.5)
        self.line(10, self.get_y(), 287, self.get_y()) # Landscape format
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.set_text_color(128)
        self.cell(0, 10, f"Página {self.page_no()}", 0, 0, "C")

def crear_pdf_pd(df_data, anio):
    # Crear PDF apaisado (Landscape) para que entren los 12 meses cómodamente
    pdf = ReportePD(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    
    meses_nombres = ['E', 'F', 'M', 'A', 'M', 'J', 'J', 'A', 'S', 'O', 'N', 'D']
    
    # --- FUNCION PARA DIBUJAR UNA TABLA DE INDICADOR (ESTILO EXCEL) ---
    def dibujar_tabla_indicador(titulo, objetivo_val, col_real, is_mtbf=True):
        pdf.set_font("Arial", 'B', 10)
        pdf.set_text_color(50, 50, 50)
        pdf.cell(0, 8, titulo, ln=True)
        
        # Configuraciones de celda
        w_indicador = 45
        w_tipo = 20
        w_mes = 16
        w_estado = 25
        h_cell = 6
        
        # CABECERAS MESES
        pdf.set_font("Arial", 'B', 8)
        pdf.set_fill_color(220, 230, 241) # Azul clarito Excel
        pdf.cell(w_indicador, h_cell, "Indicador", border=1, align='C', fill=True)
        pdf.cell(w_tipo, h_cell, "Tipo", border=1, align='C', fill=True)
        for m in meses_nombres:
            pdf.cell(w_mes, h_cell, m, border=1, align='C', fill=True)
        pdf.cell(w_estado, h_cell, "Promedio", border=1, align='C', fill=True, ln=True)
        
        # FILA OBJETIVO (T)
        pdf.set_font("Arial", '', 8)
        pdf.set_fill_color(255, 255, 255)
        pdf.cell(w_indicador, h_cell, "Matricería / Herramental", border=1, align='L')
        pdf.cell(w_tipo, h_cell, "Objetivo (T)", border=1, align='C')
        for _ in range(12):
            pdf.cell(w_mes, h_cell, f"{objetivo_val}", border=1, align='C')
        pdf.cell(w_estado, h_cell, f"{objetivo_val}", border=1, align='C', ln=True)
        
        # FILA REAL (C)
        pdf.cell(w_indicador, h_cell, "Valores Registrados", border=1, align='L')
        pdf.cell(w_tipo, h_cell, "Real (C)", border=1, align='C')
        
        suma_real = 0
        meses_activos = 0
        
        for i in range(1, 13):
            val_real = df_data[df_data['Mes'] == i][col_real].values[0]
            if val_real > 0:
                suma_real += val_real
                meses_activos += 1
                
            val_str = f"{val_real:.1f}" if val_real > 0 else "-"
            
            # Formato Condicional de Color (K = Verde, L = Rojo)
            if val_real > 0:
                if is_mtbf:
                    if val_real >= objetivo_val: pdf.set_text_color(33, 195, 84) # Verde
                    else: pdf.set_text_color(220, 20, 20) # Rojo
                else:
                    if val_real <= objetivo_val: pdf.set_text_color(33, 195, 84) # Verde
                    else: pdf.set_text_color(220, 20, 20) # Rojo
            else:
                pdf.set_text_color(50, 50, 50)
                
            pdf.cell(w_mes, h_cell, val_str, border=1, align='C')
            pdf.set_text_color(50, 50, 50) # Reset
        
        # Promedio Anual
        promedio = suma_real / meses_activos if meses_activos > 0 else 0
        pdf.set_font("Arial", 'B', 8)
        pdf.cell(w_estado, h_cell, f"{promedio:.1f}", border=1, align='C', ln=True)
        pdf.ln(5)

    # 1. TABLA MTBF
    dibujar_tabla_indicador(f"MTBF - Tiempo Medio Entre Fallas MATRICERIA (Objetivo: {TARGET_MTBF_HS} Hs)", TARGET_MTBF_HS, 'MTBF (Hs)', is_mtbf=True)
    
    # 2. TABLA MTTR
    dibujar_tabla_indicador(f"MTTR - Tiempo Medio de Parada MATRICERIA (Objetivo: {TARGET_MTTR_MIN} Min)", TARGET_MTTR_MIN, 'MTTR (Min)', is_mtbf=False)

    # --- GRAFICOS PLOTLY AL PDF ---
    pdf.ln(5)
    y_base = pdf.get_y()
    
    # Grafico Evolución Mensual
    df_plot = df_data[df_data['Uptime_Hs'] > 0].copy()
    if not df_plot.empty:
        meses_map = {1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'}
        df_plot['Mes_Str'] = df_plot['Mes'].map(meses_map)
        
        fig = go.Figure()
        fig.add_trace(go.Bar(x=df_plot['Mes_Str'], y=df_plot['MTBF (Hs)'], name="MTBF (Hs)", marker_color='#1f77b4', text=df_plot['MTBF (Hs)'].round(1), textposition='auto'))
        fig.add_trace(go.Scatter(x=df_plot['Mes_Str'], y=df_plot['MTTR (Min)'], name="MTTR (Min)", mode='lines+markers', yaxis='y2', line=dict(color='#ff7f0e', width=3), marker=dict(size=8)))
        
        fig.update_layout(
            title="Evolución Mensual MTBF vs MTTR",
            yaxis=dict(title="MTBF (Horas)"),
            yaxis2=dict(title="MTTR (Minutos)", overlaying='y', side='right'),
            legend=dict(x=0.01, y=0.99),
            margin=dict(l=40, r=40, t=40, b=30),
            height=300, width=900, plot_bgcolor='rgba(0,0,0,0)'
        )
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_chart:
            fig.write_image(tmp_chart.name, engine="kaleido")
            pdf.image(tmp_chart.name, x=10, y=y_base, w=270)
            os.remove(tmp_chart.name)

    return pdf.output(dest='S').encode('latin-1')

# ==========================================
# VISUALIZACIÓN DASHBOARD Y DESCARGA
# ==========================================
if not df_anual.empty:
    st.subheader(f"📊 Tablero Mensual - Año {anio_sel}")
    
    meses_map = {1:'Ene', 2:'Feb', 3:'Mar', 4:'Abr', 5:'May', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dic'}
    df_visual = df_anual.copy()
    df_visual['Mes'] = df_visual['Mes'].map(meses_map)
    df_visual = df_visual[['Mes', 'Cantidad_Fallas', 'Uptime_Hs', 'Downtime_Min', 'MTBF (Hs)', 'MTTR (Min)']].round(2)
    
    # Botón de Descarga PDF
    col_v1, col_v2 = st.columns([3, 1])
    with col_v1:
        st.dataframe(df_visual.set_index('Mes').T, use_container_width=True) # Mostrar tabla transpuesta estilo Excel
    with col_v2:
        st.write("📥 **Exportar Documento PD**")
        st.info("Descarga el reporte en formato tabla mensual con comparativa Objetivo vs Real.")
        try:
            pdf_bytes = crear_pdf_pd(df_anual, anio_sel)
            st.download_button(
                label="📄 Descargar PDF (Formato PD)",
                data=pdf_bytes,
                file_name=f"Indicadores_Matriceria_PD_{anio_sel}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"Error al generar PDF: {e}")
            
    st.divider()
    
    # Gráfico interactivo en Streamlit
    st.subheader("Tendencia Anual de Confiabilidad")
    df_chart = df_anual[df_anual['Uptime_Hs'] > 0].copy()
    if not df_chart.empty:
        df_chart['Mes_Str'] = df_chart['Mes'].map(meses_map)
        
        fig = go.Figure()
        fig.add_trace(go.Bar(x=df_chart['Mes_Str'], y=df_chart['MTBF (Hs)'], name="MTBF Real (Hs)", marker_color='#3498DB'))
        fig.add_trace(go.Scatter(x=df_chart['Mes_Str'], y=[TARGET_MTBF_HS]*len(df_chart), name="Objetivo MTBF", mode='lines', line=dict(color='green', dash='dash')))
        
        fig.add_trace(go.Scatter(x=df_chart['Mes_Str'], y=df_chart['MTTR (Min)'], name="MTTR Real (Min)", yaxis='y2', mode='lines+markers', line=dict(color='#E74C3C', width=3)))
        fig.add_trace(go.Scatter(x=df_chart['Mes_Str'], y=[TARGET_MTTR_MIN]*len(df_chart), name="Objetivo MTTR", yaxis='y2', mode='lines', line=dict(color='orange', dash='dash')))

        fig.update_layout(
            yaxis=dict(title="MTBF (Horas)"),
            yaxis2=dict(title="MTTR (Minutos)", overlaying='y', side='right'),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            plot_bgcolor='white', hovermode="x unified"
        )
        st.plotly_chart(fig, use_container_width=True)

else:
    st.warning("No hay datos disponibles para el año seleccionado.")
