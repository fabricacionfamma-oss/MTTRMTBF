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
    st.title("⚙️ Análisis de MTBF, MTTR y Down Time")
    st.write("Indicadores de Confiabilidad y Mantenibilidad en MINUTOS con formato matricial (PD Excel).")
with col_btn:
    if st.button("Limpiar Caché", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

st.divider()

# ==========================================
# OBJETIVOS (TARGETS) SEGÚN EXCEL PD
# ==========================================
TARGET_DT_PCT = 5.2       # 5.2% (0.052 en el Excel)
TARGET_MTTR_MIN = 30      # 30 minutos
TARGET_MTBF_MIN = 500     # 500 minutos (ajustar si el objetivo real era diferente)

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
        
        q_uptime = f"""
            SELECT MONTH(p.Date) as Mes, 
                   SUM(p.ProductiveTime) as Tiempo_Productivo_Min,
                   SUM(p.ProductiveTime + p.DownTime) as Tiempo_Total_Disponible_Min
            FROM PROD_D_03 p
            JOIN CELL c ON p.CellId = c.CellId
            WHERE YEAR(p.Date) = {anio}
            GROUP BY MONTH(p.Date)
        """
        df_uptime = conn.query(q_uptime)
        
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
        
        df_meses = pd.DataFrame({'Mes': range(1, 13)})
        df_anual = pd.merge(df_meses, df_uptime, on='Mes', how='left')
        df_anual = pd.merge(df_anual, df_fallas, on='Mes', how='left').fillna(0)
        
        # Mantener los cálculos base estrictamente en Minutos
        df_anual['Uptime_Min'] = df_anual['Tiempo_Productivo_Min']
        df_anual['Downtime_Min'] = df_anual['Tiempo_Reparacion_Min']
        
        # Cálculos de KPI
        df_anual['DT (%)'] = df_anual.apply(lambda r: (r['Downtime_Min'] / r['Tiempo_Total_Disponible_Min'] * 100) if r['Tiempo_Total_Disponible_Min'] > 0 else 0, axis=1)
        
        # MTBF en Minutos
        df_anual['MTBF (Min)'] = df_anual.apply(lambda r: r['Uptime_Min'] / r['Cantidad_Fallas'] if r['Cantidad_Fallas'] > 0 else (r['Uptime_Min'] if r['Uptime_Min'] > 0 else 0), axis=1)
        
        # MTTR en Minutos
        df_anual['MTTR (Min)'] = df_anual.apply(lambda r: r['Downtime_Min'] / r['Cantidad_Fallas'] if r['Cantidad_Fallas'] > 0 else 0, axis=1)
        
        return df_anual
    except Exception as e:
        st.error(f"Error consultando BD: {e}")
        return pd.DataFrame()

df_anual = fetch_annual_data(anio_sel)

# ==========================================
# GENERADOR PDF (FORMATO ESTILO EXCEL PD)
# ==========================================
class ReportePD(FPDF):
    def header(self):
        self.set_font("Arial", 'B', 14)
        self.set_text_color(15, 76, 129)
        self.cell(0, 10, f"Reporte de Indicadores de Mantenimiento Matrices (Formato PD) - Año {anio_sel}", ln=True, align='C')
        self.set_draw_color(15, 76, 129)
        self.set_line_width(0.5)
        self.line(10, self.get_y(), 287, self.get_y())
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.set_text_color(128)
        self.cell(0, 10, f"Página {self.page_no()}", 0, 0, "C")

def crear_pdf_pd_excel(df_data, anio):
    pdf = ReportePD(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    
    def dibujar_bloque_excel(x, y, titulo, objetivo_val, col_real, is_lower_better, is_pct=False):
        pdf.set_xy(x, y)
        w_lbl = 8    
        w_m = 9.5    
        w_tot = w_lbl + (w_m * 12) 
        
        # 1. FILA DE TÍTULO
        pdf.set_font("Arial", 'B', 8)
        pdf.set_text_color(255, 255, 255)
        pdf.set_fill_color(31, 78, 121) 
        pdf.set_draw_color(0, 0, 0)
        pdf.set_line_width(0.2)
        pdf.cell(w_tot - 24, 5, " " + titulo, border=1, align='L', fill=True)
        pdf.set_fill_color(189, 195, 199)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(12, 5, "Estado", border=1, align='C', fill=True)
        pdf.cell(12, 5, "Tend.", border=1, align='C', fill=True)
        
        # 2. FILA DE MESES
        pdf.set_xy(x, y + 5)
        pdf.cell(w_lbl, 5, "", border=0, align='C') 
        pdf.set_fill_color(221, 235, 247) 
        meses = ['E', 'F', 'M', 'A', 'M', 'J', 'J', 'A', 'S', 'O', 'N', 'D']
        for m in meses:
            pdf.cell(w_m, 5, m, border=1, align='C', fill=True)
            
        # 3. FILA T (OBJETIVO)
        pdf.set_xy(x, y + 10)
        pdf.set_font("Arial", 'B', 8)
        pdf.cell(w_lbl, 5, "T", border=1, align='C', fill=True)
        pdf.set_font("Arial", '', 7)
        pdf.set_fill_color(255, 255, 255)
        obj_str = f"{objetivo_val}%" if is_pct else f"{objetivo_val}"
        for _ in range(12):
            pdf.cell(w_m, 5, obj_str, border=1, align='C')
            
        # 4. FILA C (REAL)
        pdf.set_xy(x, y + 15)
        pdf.set_font("Arial", 'B', 8)
        pdf.set_fill_color(221, 235, 247)
        pdf.set_text_color(0,0,0)
        pdf.cell(w_lbl, 5, "C", border=1, align='C', fill=True)
        pdf.set_font("Arial", 'B', 7)
        pdf.set_fill_color(255, 255, 255)
        
        for i in range(1, 13):
            val = df_data[df_data['Mes'] == i][col_real].values[0]
            if val > 0:
                # Si no es porcentaje, mostramos sin decimales para los minutos que pueden ser grandes
                val_str = f"{val:.1f}%" if is_pct else f"{val:.0f}" 
                
                if is_lower_better:
                    if val <= objetivo_val: pdf.set_text_color(33, 195, 84) # Verde
                    else: pdf.set_text_color(220, 20, 20) # Rojo
                else:
                    if val >= objetivo_val: pdf.set_text_color(33, 195, 84) # Verde
                    else: pdf.set_text_color(220, 20, 20) # Rojo
            else:
                val_str = "-"
                pdf.set_text_color(150, 150, 150) 
                
            pdf.cell(w_m, 5, val_str, border=1, align='C')
        pdf.set_text_color(0,0,0) 
        
        return y + 25

    # Bloque 1: Down Time (Izquierda Arriba)
    dibujar_bloque_excel(x=15, y=30, titulo="Down Time Matriceria", objetivo_val=TARGET_DT_PCT, col_real='DT (%)', is_lower_better=True, is_pct=True)
    
    # Bloque 2: MTTR (Derecha Arriba)
    dibujar_bloque_excel(x=155, y=30, titulo="MTTR - Tiempo medio parada (Min)", objetivo_val=TARGET_MTTR_MIN, col_real='MTTR (Min)', is_lower_better=True)
    
    # Bloque 3: MTBF (Izquierda Abajo)
    dibujar_bloque_excel(x=15, y=60, titulo="MTBF - Tiempo medio entre fallas (Min)", objetivo_val=TARGET_MTBF_MIN, col_real='MTBF (Min)', is_lower_better=False)

    # --- GRÁFICO PLOTLY DEBAJO ---
    y_base_grafico = 95
    df_plot = df_data[df_data['Tiempo_Total_Disponible_Min'] > 0].copy()
    if not df_plot.empty:
        meses_map = {1:'E', 2:'F', 3:'M', 4:'A', 5:'M', 6:'J', 7:'J', 8:'A', 9:'S', 10:'O', 11:'N', 12:'D'}
        df_plot['Mes_Str'] = df_plot['Mes'].map(meses_map)
        
        fig = go.Figure()
        fig.add_trace(go.Bar(x=df_plot['Mes_Str'], y=df_plot['MTBF (Min)'], name="MTBF (Min)", marker_color='#1f77b4', text=df_plot['MTBF (Min)'].round(0), textposition='auto'))
        fig.add_trace(go.Scatter(x=df_plot['Mes_Str'], y=df_plot['MTTR (Min)'], name="MTTR (Min)", mode='lines+markers', yaxis='y2', line=dict(color='#ff7f0e', width=3), marker=dict(size=8)))
        
        fig.update_layout(
            title="Evolución Mensual: MTBF vs MTTR (en Minutos)",
            yaxis=dict(title="MTBF (Minutos)"),
            yaxis2=dict(title="MTTR (Minutos)", overlaying='y', side='right'),
            legend=dict(x=0.01, y=1.1, orientation="h"),
            margin=dict(l=40, r=40, t=40, b=20),
            height=300, width=950, plot_bgcolor='rgba(0,0,0,0)'
        )
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_chart:
            fig.write_image(tmp_chart.name, engine="kaleido")
            pdf.image(tmp_chart.name, x=15, y=y_base_grafico, w=260)
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
    df_visual = df_visual[['Mes', 'Cantidad_Fallas', 'DT (%)', 'MTBF (Min)', 'MTTR (Min)']].round(2)
    
    col_v1, col_v2 = st.columns([3, 1])
    with col_v1:
        df_show = df_visual.set_index('Mes').T
        st.dataframe(df_show, use_container_width=True) 
    with col_v2:
        st.write("📥 **Exportar Documento PD**")
        try:
            pdf_bytes = crear_pdf_pd_excel(df_anual, anio_sel)
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
    
    st.subheader("Tendencia Anual de Confiabilidad")
    df_chart = df_anual[df_anual['Tiempo_Total_Disponible_Min'] > 0].copy()
    if not df_chart.empty:
        df_chart['Mes_Str'] = df_chart['Mes'].map(meses_map)
        
        fig = go.Figure()
        fig.add_trace(go.Bar(x=df_chart['Mes_Str'], y=df_chart['MTBF (Min)'], name="MTBF Real (Min)", marker_color='#3498DB'))
        fig.add_trace(go.Scatter(x=df_chart['Mes_Str'], y=[TARGET_MTBF_MIN]*len(df_chart), name="Objetivo MTBF", mode='lines', line=dict(color='green', dash='dash')))
        
        fig.add_trace(go.Scatter(x=df_chart['Mes_Str'], y=df_chart['MTTR (Min)'], name="MTTR Real (Min)", yaxis='y2', mode='lines+markers', line=dict(color='#E74C3C', width=3)))
        fig.add_trace(go.Scatter(x=df_chart['Mes_Str'], y=[TARGET_MTTR_MIN]*len(df_chart), name="Objetivo MTTR", yaxis='y2', mode='lines', line=dict(color='orange', dash='dash')))

        fig.update_layout(
            yaxis=dict(title="MTBF (Minutos)"),
            yaxis2=dict(title="MTTR (Minutos)", overlaying='y', side='right'),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            plot_bgcolor='white', hovermode="x unified"
        )
        st.plotly_chart(fig, use_container_width=True)

else:
    st.warning("No hay datos disponibles para el año seleccionado.")
