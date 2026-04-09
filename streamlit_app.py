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
    st.write("Indicadores con Gráficos Individuales sobre Tablas Matriciales (PD Excel).")
with col_btn:
    if st.button("Limpiar Caché", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

st.divider()

# ==========================================
# OBJETIVOS (TARGETS) SEGÚN EXCEL PD
# ==========================================
TARGET_DT_PCT = 5.2       # 5.2%
TARGET_MTTR_MIN = 30      # 30 minutos
TARGET_MTBF_MIN = 500     # 500 minutos

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
        
        df_anual['Uptime_Min'] = df_anual['Tiempo_Productivo_Min']
        df_anual['Downtime_Min'] = df_anual['Tiempo_Reparacion_Min']
        
        df_anual['DT (%)'] = df_anual.apply(lambda r: (r['Downtime_Min'] / r['Tiempo_Total_Disponible_Min'] * 100) if r['Tiempo_Total_Disponible_Min'] > 0 else 0, axis=1)
        df_anual['MTBF (Min)'] = df_anual.apply(lambda r: r['Uptime_Min'] / r['Cantidad_Fallas'] if r['Cantidad_Fallas'] > 0 else (r['Uptime_Min'] if r['Uptime_Min'] > 0 else 0), axis=1)
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
        self.cell(0, 8, f"Reporte de Indicadores de Mantenimiento Matrices - Año {anio_sel}", ln=True, align='C')
        self.set_draw_color(15, 76, 129)
        self.set_line_width(0.5)
        self.line(10, self.get_y(), 287, self.get_y())
        self.ln(2)

    def footer(self):
        self.set_y(-10)
        self.set_font("Arial", "I", 8)
        self.set_text_color(128)
        self.cell(0, 10, f"Página {self.page_no()}", 0, 0, "C")

def crear_pdf_pd_excel(df_data, anio):
    pdf = ReportePD(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    
    meses_nombres = ['E', 'F', 'M', 'A', 'M', 'J', 'J', 'A', 'S', 'O', 'N', 'D']
    
    df_plot = df_data.copy()
    df_plot['Mes_Str'] = df_plot['Mes'].map(dict(zip(range(1, 13), meses_nombres)))
    
    def generar_grafico(df, col_real, objetivo_val, titulo_grafico, is_pct=False, is_lower_better=True):
        fig = go.Figure()
        
        df_filtered = df[df['Tiempo_Total_Disponible_Min'] > 0].copy() if col_real == 'DT (%)' else df.copy()
        color_barra = '#1f77b4' 
        
        text_format = [f"{v:.1f}%" if is_pct else f"{v:.0f}" for v in df_filtered[col_real]]
        
        fig.add_trace(go.Bar(
            x=df_filtered['Mes_Str'], 
            y=df_filtered[col_real], 
            name="Real", 
            marker_color=color_barra, 
            text=text_format, 
            textposition='auto',
            textfont=dict(size=10)
        ))
        
        fig.add_trace(go.Scatter(
            x=df['Mes_Str'], 
            y=[objetivo_val] * 12, 
            name="Objetivo", 
            mode='lines', 
            line=dict(color='red', dash='dash', width=2)
        ))
        
        y_title = "Porcentaje (%)" if is_pct else "Minutos"
        
        # FIX: Se corrigió la sintaxis de 'titlefont' adaptándola a las versiones nuevas de Plotly
        fig.update_layout(
            title=dict(text=titulo_grafico, font=dict(size=14)),
            yaxis=dict(
                title=dict(text=y_title, font=dict(size=10)), 
                tickfont=dict(size=9)
            ),
            xaxis=dict(tickfont=dict(size=10)),
            margin=dict(l=30, r=10, t=30, b=20),
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=10)),
            height=200, width=450, 
            plot_bgcolor='white'
        )
        
        fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='LightGray')
        
        tmp_chart = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        fig.write_image(tmp_chart.name, engine="kaleido")
        return tmp_chart.name

    def dibujar_bloque_con_grafico(x, y, titulo, objetivo_val, col_real, is_lower_better, is_pct=False):
        img_path = generar_grafico(df_plot, col_real, objetivo_val, titulo, is_pct, is_lower_better)
        pdf.image(img_path, x=x, y=y, w=125, h=55)
        os.remove(img_path)
        
        y_tabla = y + 55 
        pdf.set_xy(x, y_tabla)
        
        w_lbl = 8    
        w_m = 9.5    
        w_tot = w_lbl + (w_m * 12) 
        
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
        
        pdf.set_xy(x, y_tabla + 5)
        pdf.cell(w_lbl, 5, "", border=0, align='C') 
        pdf.set_fill_color(221, 235, 247) 
        for m in meses_nombres:
            pdf.cell(w_m, 5, m, border=1, align='C', fill=True)
            
        pdf.set_xy(x, y_tabla + 10)
        pdf.set_font("Arial", 'B', 8)
        pdf.cell(w_lbl, 5, "T", border=1, align='C', fill=True)
        pdf.set_font("Arial", '', 7)
        pdf.set_fill_color(255, 255, 255)
        obj_str = f"{objetivo_val}%" if is_pct else f"{objetivo_val}"
        for _ in range(12):
            pdf.cell(w_m, 5, obj_str, border=1, align='C')
            
        pdf.set_xy(x, y_tabla + 15)
        pdf.set_font("Arial", 'B', 8)
        pdf.set_fill_color(221, 235, 247)
        pdf.set_text_color(0,0,0)
        pdf.cell(w_lbl, 5, "C", border=1, align='C', fill=True)
        pdf.set_font("Arial", 'B', 7)
        pdf.set_fill_color(255, 255, 255)
        
        for i in range(1, 13):
            val = df_data[df_data['Mes'] == i][col_real].values[0]
            if val > 0:
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

    y_fila_1 = 15
    dibujar_bloque_con_grafico(x=15, y=y_fila_1, titulo="Down Time Matriceria", objetivo_val=TARGET_DT_PCT, col_real='DT (%)', is_lower_better=True, is_pct=True)
    dibujar_bloque_con_grafico(x=155, y=y_fila_1, titulo="MTTR - Tiempo medio parada (Min)", objetivo_val=TARGET_MTTR_MIN, col_real='MTTR (Min)', is_lower_better=True)
    
    y_fila_2 = 105
    dibujar_bloque_con_grafico(x=15, y=y_fila_2, titulo="MTBF - Tiempo medio entre fallas (Min)", objetivo_val=TARGET_MTBF_MIN, col_real='MTBF (Min)', is_lower_better=False)

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
    
    st.subheader("Gráficos de Tendencia (Vista Web)")
    df_chart = df_anual[df_anual['Tiempo_Total_Disponible_Min'] > 0].copy()
    
    if not df_chart.empty:
        df_chart['Mes_Str'] = df_chart['Mes'].map(meses_map)
        
        c1, c2, c3 = st.columns(3)
        
        with c1:
            fig_dt = px.bar(df_chart, x='Mes_Str', y='DT (%)', title='Down Time (%)')
            fig_dt.add_hline(y=TARGET_DT_PCT, line_dash="dash", line_color="red", annotation_text="Objetivo")
            st.plotly_chart(fig_dt, use_container_width=True)
            
        with c2:
            fig_mttr = px.bar(df_chart, x='Mes_Str', y='MTTR (Min)', title='MTTR (Min)')
            fig_mttr.add_hline(y=TARGET_MTTR_MIN, line_dash="dash", line_color="red", annotation_text="Objetivo")
            st.plotly_chart(fig_mttr, use_container_width=True)
            
        with c3:
            fig_mtbf = px.bar(df_chart, x='Mes_Str', y='MTBF (Min)', title='MTBF (Min)')
            fig_mtbf.add_hline(y=TARGET_MTBF_MIN, line_dash="dash", line_color="red", annotation_text="Objetivo")
            st.plotly_chart(fig_mtbf, use_container_width=True)

else:
    st.warning("No hay datos disponibles para el año seleccionado.")
