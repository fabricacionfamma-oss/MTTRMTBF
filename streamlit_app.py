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
    st.write("Dashboard con Gráficos y Tablas Matriciales (T = Límite Sup, C = Límite Inf, A = Real Mensual).")
with col_btn:
    if st.button("Limpiar Caché", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

st.divider()

# ==========================================
# OBJETIVOS (TARGETS T y C)
# ==========================================
TARGET_DT_T = 5.2       # Límite Superior Down Time
TARGET_DT_C = 3.0       # Límite Inferior Down Time

TARGET_MTTR_T = 30      # Límite Superior MTTR
TARGET_MTTR_C = 20      # Límite Inferior MTTR

TARGET_MTBF_T = 600     # Límite Superior MTBF
TARGET_MTBF_C = 500     # Límite Inferior MTBF

# ==========================================
# FILTROS
# ==========================================
anio_actual = pd.to_datetime("today").year
anio_sel = st.selectbox("Seleccione el Año para el Análisis y Reporte:", range(2023, anio_actual + 2), index=anio_actual-2023)

# ==========================================
# EXTRACCIÓN Y PROCESAMIENTO DE DATOS ANUALES
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

    def generar_grafico_tendencia_pdf(df, col_real, obj_t, obj_c, is_pct):
        df_plot = df.copy()
        df_plot['Mes_Str'] = df_plot['Mes'].map(dict(zip(range(1, 13), meses_nombres)))
        df_filtered = df_plot[df_plot['Tiempo_Total_Disponible_Min'] > 0].copy() if col_real == 'DT (%)' else df_plot.copy()
        
        fig = go.Figure()
        text_format = [f"{v:.1f}%" if is_pct else f"{v:.0f}" for v in df_filtered[col_real]]
        
        fig.add_trace(go.Bar(
            x=df_filtered['Mes_Str'], y=df_filtered[col_real], name="Real (A)",
            marker_color='#1f77b4', text=text_format, textposition='auto', textfont=dict(size=12)
        ))
        
        fig.add_trace(go.Scatter(
            x=df_plot['Mes_Str'], y=[obj_t] * 12, name="Sup. (T)",
            mode='lines', line=dict(color='red', dash='dash', width=2)
        ))
        
        fig.add_trace(go.Scatter(
            x=df_plot['Mes_Str'], y=[obj_c] * 12, name="Inf. (C)",
            mode='lines', line=dict(color='orange', dash='dot', width=2)
        ))
        
        y_title = "Porcentaje (%)" if is_pct else "Minutos"
        
        fig.update_layout(
            yaxis=dict(title=dict(text=y_title, font=dict(size=10)), tickfont=dict(size=9)),
            xaxis=dict(tickfont=dict(size=10)),
            # Se aumentó el margen superior (t=25) para darle aire internamente al gráfico
            margin=dict(l=40, r=10, t=25, b=30), 
            height=200, width=550, 
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=10)),
            plot_bgcolor='white'
        )
        fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='LightGray')
        
        tmp_chart = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        fig.write_image(tmp_chart.name, engine="kaleido")
        return tmp_chart.name

    def dibujar_bloque_completo(x, y, titulo, obj_t, obj_c, col_real, is_lower_better, is_pct=False):
        # --- 1. DIBUJAR TÍTULO OSCURO ---
        pdf.set_xy(x, y)
        w_lbl = 8; w_m = 9.5; w_tot = w_lbl + (w_m * 12) 
        pdf.set_font("Arial", 'B', 8)
        pdf.set_text_color(255, 255, 255); pdf.set_fill_color(31, 78, 121); pdf.set_draw_color(0, 0, 0); pdf.set_line_width(0.2)
        pdf.cell(w_tot, 6, "  " + titulo, border=1, align='L', fill=True)

        # --- 2. DIBUJAR GRÁFICO (Con más espacio respecto al título) ---
        img_path = generar_grafico_tendencia_pdf(df_data, col_real, obj_t, obj_c, is_pct)
        # Bajamos la imagen a y+12 para separarla del bloque azul
        pdf.image(img_path, x=x, y=y + 12, w=w_tot, h=45)
        os.remove(img_path)

        # --- 3. DIBUJAR TABLA (T, C, A) ---
        # Bajamos la tabla a y+62 para que no pise el gráfico
        y_tabla = y + 62 
        pdf.set_xy(x, y_tabla)
        pdf.set_fill_color(221, 235, 247) 
        
        # Cabecera de Meses
        pdf.cell(w_lbl, 5, "", border=0, align='C') 
        for m in meses_nombres: 
            pdf.cell(w_m, 5, m, border=1, align='C', fill=True)
            
        # Fila T 
        pdf.set_xy(x, y_tabla + 5)
        pdf.set_font("Arial", 'B', 8)
        pdf.cell(w_lbl, 5, "T", border=1, align='C', fill=True)
        pdf.set_font("Arial", '', 7); pdf.set_fill_color(255, 255, 255)
        t_str = f"{obj_t}%" if is_pct else f"{obj_t}"
        for _ in range(12): 
            pdf.cell(w_m, 5, t_str, border=1, align='C')
            
        # Fila C 
        pdf.set_xy(x, y_tabla + 10)
        pdf.set_font("Arial", 'B', 8)
        pdf.set_fill_color(221, 235, 247); pdf.set_text_color(0,0,0)
        pdf.cell(w_lbl, 5, "C", border=1, align='C', fill=True)
        pdf.set_font("Arial", '', 7); pdf.set_fill_color(255, 255, 255)
        c_str = f"{obj_c}%" if is_pct else f"{obj_c}"
        for _ in range(12): 
            pdf.cell(w_m, 5, c_str, border=1, align='C')
            
        # Fila A 
        pdf.set_xy(x, y_tabla + 15)
        pdf.set_font("Arial", 'B', 8)
        pdf.set_fill_color(221, 235, 247); pdf.set_text_color(0,0,0)
        pdf.cell(w_lbl, 5, "A", border=1, align='C', fill=True)
        pdf.set_font("Arial", 'B', 7); pdf.set_fill_color(255, 255, 255)
        
        for i in range(1, 13):
            val_a = df_data[df_data['Mes'] == i][col_real].values[0]
            if df_data[df_data['Mes'] == i]['Tiempo_Total_Disponible_Min'].values[0] > 0:
                val_str = f"{val_a:.1f}%" if is_pct else f"{val_a:.0f}" 
                
                if is_lower_better:
                    if val_a <= obj_c: pdf.set_text_color(33, 195, 84)       
                    elif val_a > obj_t: pdf.set_text_color(220, 20, 20)      
                    else: pdf.set_text_color(200, 150, 0)                    
                else:
                    if val_a >= obj_t: pdf.set_text_color(33, 195, 84)       
                    elif val_a < obj_c: pdf.set_text_color(220, 20, 20)      
                    else: pdf.set_text_color(200, 150, 0)                    
            else:
                val_str = "-"
                pdf.set_text_color(150, 150, 150) 
            pdf.cell(w_m, 5, val_str, border=1, align='C')

        pdf.set_text_color(0,0,0) 

    # --- DIBUJAR LOS BLOQUES ESPACIADOS ---
    dibujar_bloque_completo(x=15, y=15, titulo="Down Time Matriceria", obj_t=TARGET_DT_T, obj_c=TARGET_DT_C, col_real='DT (%)', is_lower_better=True, is_pct=True)
    dibujar_bloque_completo(x=150, y=15, titulo="MTTR - Tiempo medio parada (Min)", obj_t=TARGET_MTTR_T, obj_c=TARGET_MTTR_C, col_real='MTTR (Min)', is_lower_better=True)
    
    # Bajamos la Fila 2 a Y=110 para asegurar que no roce nada de la tabla superior
    dibujar_bloque_completo(x=15, y=110, titulo="MTBF - Tiempo medio entre fallas (Min)", obj_t=TARGET_MTBF_T, obj_c=TARGET_MTBF_C, col_real='MTBF (Min)', is_lower_better=False)

    return pdf.output(dest='S').encode('latin-1')

# ==========================================
# VISUALIZACIÓN DASHBOARD EN WEB Y DESCARGA
# ==========================================
if not df_anual.empty:
    st.subheader(f"📊 Dashboard Mantenimiento Matrices - Año {anio_sel}")
    
    def renderizar_grafico_web(df, col_real, obj_t, obj_c, titulo, is_pct):
        meses_nombres = ['E', 'F', 'M', 'A', 'M', 'J', 'J', 'A', 'S', 'O', 'N', 'D']
        df_plot = df.copy()
        df_plot['Mes_Str'] = df_plot['Mes'].map(dict(zip(range(1, 13), meses_nombres)))
        df_filtered = df_plot[df_plot['Tiempo_Total_Disponible_Min'] > 0].copy() if col_real == 'DT (%)' else df_plot.copy()
        
        fig = go.Figure()
        text_format = [f"{v:.1f}%" if is_pct else f"{v:.0f}" for v in df_filtered[col_real]]
        
        fig.add_trace(go.Bar(
            x=df_filtered['Mes_Str'], y=df_filtered[col_real], name="Real (A)",
            marker_color='#3498DB', text=text_format, textposition='auto', textfont=dict(size=12)
        ))
        fig.add_trace(go.Scatter(
            x=df_plot['Mes_Str'], y=[obj_t] * 12, name="Sup. (T)",
            mode='lines', line=dict(color='red', dash='dash', width=2)
        ))
        fig.add_trace(go.Scatter(
            x=df_plot['Mes_Str'], y=[obj_c] * 12, name="Inf. (C)",
            mode='lines', line=dict(color='orange', dash='dot', width=2)
        ))
        
        fig.update_layout(
            title=titulo,
            margin=dict(l=20, r=20, t=40, b=20), height=250,
            showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            plot_bgcolor='rgba(0,0,0,0)'
        )
        fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='LightGray')
        return fig

    # --- Vista en la Web ---
    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(renderizar_grafico_web(df_anual, 'DT (%)', TARGET_DT_T, TARGET_DT_C, "Down Time Matricería", True), use_container_width=True)
    with c2:
        st.plotly_chart(renderizar_grafico_web(df_anual, 'MTTR (Min)', TARGET_MTTR_T, TARGET_MTTR_C, "MTTR (Min)", False), use_container_width=True)
    
    c3, _ = st.columns(2)
    with c3:
        st.plotly_chart(renderizar_grafico_web(df_anual, 'MTBF (Min)', TARGET_MTBF_T, TARGET_MTBF_C, "MTBF (Min)", False), use_container_width=True)
    
    st.divider()
    st.write("📥 **Exportar Documento PD (Formato Excel T, C, A)**")
    try:
        pdf_bytes = crear_pdf_pd_excel(df_anual, anio_sel)
        st.download_button(
            label="📄 Descargar Reporte PDF",
            data=pdf_bytes,
            file_name=f"Indicadores_Matriceria_PD_{anio_sel}.pdf",
            mime="application/pdf"
        )
    except Exception as e:
        st.error(f"Error al generar PDF: {e}")

else:
    st.warning("No hay datos disponibles para el año seleccionado.")
