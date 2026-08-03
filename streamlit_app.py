# ==========================================
# EXTRACCIÓN Y PROCESAMIENTO DE DATOS
# ==========================================

@st.cache_data(ttl=3600) # Cacheamos las exclusiones por 1 hora
def obtener_exclusiones_sheet():
    # Convertimos la URL de /edit a /export?format=csv para leerla como tabla
    url = "https://docs.google.com/spreadsheets/d/1mLnIC8B7mwmFZwthO0A32H3ZFfXSKt7vIUMBXEZxDJ0/export?format=csv&gid=0"
    try:
        df = pd.read_csv(url)
        # Tomamos la primera columna, quitamos nulos, pasamos a texto y eliminamos espacios
        exclusiones = df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
        # Escapamos comillas simples para que no rompan el formato SQL (ej: O'Connor -> O''Connor)
        exclusiones = [e.replace("'", "''") for e in exclusiones if e]
        return exclusiones
    except Exception as e:
        st.warning(f"No se pudieron cargar las exclusiones del Google Sheet: {e}")
        return []

@st.cache_data(ttl=300)
def fetch_annual_data(anio):
    # 1. Obtener la lista dinámica del sheet
    exclusiones = obtener_exclusiones_sheet()
    
    # 2. Armar la cláusula SQL para excluir los ítems del Sheet
    filtro_exclusiones_sql = ""
    if exclusiones:
        lista_sql = "'" + "','".join(exclusiones) + "'"
        filtro_exclusiones_sql = f"""
            AND (t1.Name NOT IN ({lista_sql}) OR t1.Name IS NULL)
            AND (t2.Name NOT IN ({lista_sql}) OR t2.Name IS NULL)
            AND (t3.Name NOT IN ({lista_sql}) OR t3.Name IS NULL)
            AND (t4.Name NOT IN ({lista_sql}) OR t4.Name IS NULL)
        """

    try:
        conn = st.connection("wii_bi", type="sql")
        
        q_uptime = f"""
            SELECT MONTH(p.Date) as Mes, 
                   SUM(p.ProductiveTime) as Tiempo_Productivo_Min,
                   SUM(p.ProductiveTime + p.DownTime) as Tiempo_Total_Disponible_Min
            FROM PROD_D_03 p
            JOIN CELL c ON p.CellId = c.CellId
            WHERE YEAR(p.Date) = {anio}
              AND c.Name IN ({sql_maquinas_in})
            GROUP BY MONTH(p.Date)
        """
        df_uptime = conn.query(q_uptime)
        
        # 3. Consulta de fallas modificada
        q_fallas = f"""
            SELECT MONTH(e.Date) as Mes, 
                   COUNT(e.Id) as Cantidad_Fallas,
                   SUM(e.Interval) as Tiempo_Reparacion_Min
            FROM EVENT_01 e
            LEFT JOIN EVENTTYPE t1 ON e.EventTypeLevel1 = t1.EventTypeId
            LEFT JOIN EVENTTYPE t2 ON e.EventTypeLevel2 = t2.EventTypeId
            LEFT JOIN EVENTTYPE t3 ON e.EventTypeLevel3 = t3.EventTypeId
            LEFT JOIN EVENTTYPE t4 ON e.EventTypeLevel4 = t4.EventTypeId
            LEFT JOIN CELL c ON e.CellId = c.CellId
            WHERE YEAR(e.Date) = {anio}
              AND c.Name IN ({sql_maquinas_in})
              -- EVENTOS DE MATRICERÍA/HERRAMENTAL
              AND (
                  UPPER(t1.Name) LIKE '%MATRI%' OR UPPER(t2.Name) LIKE '%MATRI%' OR UPPER(t3.Name) LIKE '%MATRI%' OR UPPER(t4.Name) LIKE '%MATRI%'
                  OR UPPER(t1.Name) LIKE '%HERRAMENTAL%' OR UPPER(t2.Name) LIKE '%HERRAMENTAL%'
              )
              -- EXCLUSIÓN 1: La palabra "PROYECTO" (COALESCE evita que los NULL anulen el filtro)
              AND UPPER(COALESCE(t1.Name, '')) NOT LIKE '%PROYECTO%'
              AND UPPER(COALESCE(t2.Name, '')) NOT LIKE '%PROYECTO%'
              AND UPPER(COALESCE(t3.Name, '')) NOT LIKE '%PROYECTO%'
              AND UPPER(COALESCE(t4.Name, '')) NOT LIKE '%PROYECTO%'
              -- EXCLUSIÓN 2: Los eventos listados en el Google Sheet
              {filtro_exclusiones_sql}
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
        
        df_anual['Cum_Uptime'] = df_anual['Uptime_Min'].cumsum()
        df_anual['Cum_Downtime'] = df_anual['Downtime_Min'].cumsum()
        df_anual['Cum_TotalTime'] = df_anual['Tiempo_Total_Disponible_Min'].cumsum()
        df_anual['Cum_Fallas'] = df_anual['Cantidad_Fallas'].cumsum()

        df_anual['A_DT (%)'] = df_anual.apply(lambda r: (r['Cum_Downtime'] / r['Cum_TotalTime'] * 100) if r['Cum_TotalTime'] > 0 else 0, axis=1)
        df_anual['A_MTBF (Min)'] = df_anual.apply(lambda r: r['Cum_Uptime'] / r['Cum_Fallas'] if r['Cum_Fallas'] > 0 else (r['Cum_Uptime'] if r['Cum_Uptime'] > 0 else 0), axis=1)
        df_anual['A_MTTR (Min)'] = df_anual.apply(lambda r: r['Cum_Downtime'] / r['Cum_Fallas'] if r['Cum_Fallas'] > 0 else 0, axis=1)

        return df_anual
    except Exception as e:
        st.error(f"Error consultando BD: {e}")
        return pd.DataFrame()

df_anual = fetch_annual_data(anio_sel)
