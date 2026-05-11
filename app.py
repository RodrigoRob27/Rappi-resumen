import streamlit as st
import pandas as pd
from sqlalchemy import create_engine

# ==========================================
# CONFIGURACIÓN DE PÁGINA
# ==========================================
st.set_page_config(page_title="Pedidos Ya", layout="centered")

# ==========================================
# CONFIGURACIÓN DE BASE DE DATOS (Segura)
# ==========================================
try:
    # Intenta leer las credenciales guardadas en Streamlit Cloud
    secrets = st.secrets["db"]
    DB_CONFIG = {
        "server": secrets["server"],
        "database": secrets["database"],
        "user": secrets["user"],
        "password": secrets["password"],
        "driver": secrets["driver"]
    }
except:
    # Modo local (si no hay secrets configurados)
    DB_CONFIG = {
        "server": "minerva.fridaysperu.com",
        "database": "DW_Simphony_Prod",
        "user": "rrobles",
        "password": "TU_CONTRASEÑA_AQUÍ",
        "driver": "ODBC Driver 17 for SQL Server"
    }

def get_engine():
    conn_str = f"mssql+pyodbc://{DB_CONFIG['user']}:{DB_CONFIG['password']}@{DB_CONFIG['server']}/{DB_CONFIG['database']}?driver={DB_CONFIG['driver']}"
    return create_engine(conn_str, fast_executemany=True)

# ==========================================
# VALIDACIÓN DE DUPLICADOS (CONTRA ARCHIVO Y BD)
# ==========================================
def validar_archivo_sin_duplicados(xls, engine=None):
    hojas_map = {
        "Lista de ordenes": ("pedidos_lista", "Número de pedido"),
        "Cargos por cancelaciones": ("pedidos_cancelaciones", "Número de pedido"),
        "Cargos por reclamos": ("pedidos_reclamos", ["Número de pedido", "Razon", "Monto"]),
        "Reintegros": ("pedidos_reintegros", "Número de pedido")
    }
    
    errores = []
    advertencias = []
    dataframes = {}
    
    for hoja, (tabla, col_id) in hojas_map.items():
        if hoja in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=hoja)
            df.columns = df.columns.str.strip()
            
            # 1. Validar duplicados INTERNOS en el Excel
            if isinstance(col_id, list):
                cols_existentes = [c for c in col_id if c in df.columns]
                if len(cols_existentes) == len(col_id):
                    duplicados_internos = df[df.duplicated(subset=cols_existentes, keep=False)]
                else:
                    continue
            else:
                if col_id in df.columns:
                    duplicados_internos = df[df.duplicated(subset=[col_id], keep=False)]
                else:
                    continue
            
            if not duplicados_internos.empty:
                errores.append(f"Hoja '{hoja}': Se detectaron registros duplicados internamente")
                continue
            
            # 2. Si hay conexión a BD, validar contra base de datos
            if engine is not None:
                try:
                    if tabla == "pedidos_reclamos":
                        query = f"SELECT [Número de pedido], Razon, Monto FROM dbo.{tabla}"
                        df_bd = pd.read_sql(query, con=engine)
                        
                        df['llave'] = df['Número de pedido'].astype(str) + '|' + df['Razon'].astype(str) + '|' + df['Monto'].astype(str)
                        df_bd['llave'] = df_bd['Número de pedido'].astype(str) + '|' + df_bd['Razon'].astype(str) + '|' + df_bd['Monto'].astype(str)
                        
                        existentes = df[df['llave'].isin(df_bd['llave'])]
                        nuevos = df[~df['llave'].isin(df_bd['llave'])]
                        
                        df = df.drop(columns=['llave'])
                    else:
                        query = f"SELECT [Número de pedido] FROM dbo.{tabla}"
                        df_bd = pd.read_sql(query, con=engine)
                        
                        existentes = df[df[col_id].isin(df_bd[col_id])]
                        nuevos = df[~df[col_id].isin(df_bd[col_id])]
                    
                    if len(existentes) > 0:
                        advertencias.append(f"'{hoja}': {len(existentes)} registros ya existen en la base de datos (se omitirán)")
                    
                    dataframes[tabla] = nuevos
                    
                except Exception as e:
                    dataframes[tabla] = df
            else:
                dataframes[tabla] = df
    
    if errores:
        return False, errores, advertencias, {}
    
    return True, errores, advertencias, dataframes

# ==========================================
# CARGA A BASE DE DATOS
# ==========================================
def cargar_a_bd(dataframes, engine):
    resultados = {}
    
    for nombre_tabla, df in dataframes.items():
        if df.empty:
            resultados[nombre_tabla] = {"status": "warning", "msg": "No hay registros nuevos para procesar"}
            continue
            
        try:
            if nombre_tabla == "pedidos_reclamos":
                query = f"SELECT [Número de pedido], Razon, Monto FROM dbo.{nombre_tabla}"
                df_existentes = pd.read_sql(query, con=engine)
                
                df['llave'] = df['Número de pedido'].astype(str) + '|' + df['Razon'].astype(str) + '|' + df['Monto'].astype(str)
                df_existentes['llave'] = df_existentes['Número de pedido'].astype(str) + '|' + df_existentes['Razon'].astype(str) + '|' + df_existentes['Monto'].astype(str)
                
                df_nuevos = df[~df['llave'].isin(df_existentes['llave'])].drop(columns=['llave'])
            else:
                query = f"SELECT [Número de pedido] FROM dbo.{nombre_tabla}"
                df_existentes = pd.read_sql(query, con=engine)
                df_nuevos = df[~df['Número de pedido'].isin(df_existentes['Número de pedido'])]
            
            total_nuevos = len(df_nuevos)
            
            if total_nuevos > 0:
                with engine.begin() as conn:
                    df_nuevos.to_sql(nombre_tabla, con=conn, if_exists='append', index=False, method='multi')
                
                resultados[nombre_tabla] = {"status": "success", "nuevos": total_nuevos}
            else:
                resultados[nombre_tabla] = {"status": "warning", "msg": "Todos los registros ya se encuentran en la base de datos"}
                
        except Exception as e:
            resultados[nombre_tabla] = {"status": "error", "msg": str(e)}
    
    return resultados

# ==========================================
# INTERFAZ PRINCIPAL
# ==========================================
def main():
    st.title("Pedidos Ya | Carga de Informes")
    st.markdown("Seleccione el archivo Excel semanal. El sistema validará automáticamente la existencia de duplicados.")
    
    uploaded_file = st.file_uploader("Seleccionar archivo Excel", type=['xls', 'xlsx'])
    
    if uploaded_file:
        st.info(f"Archivo seleccionado: **{uploaded_file.name}**")
        
        try:
            engine = get_engine()
        except Exception as e:
            st.error(f"Error de conexión a la base de datos: {e}")
            engine = None
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("Validar archivo", use_container_width=True):
                with st.spinner("Ejecutando validación..."):
                    try:
                        xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
                        es_valido, errores, advertencias, dataframes = validar_archivo_sin_duplicados(xls, engine)
                        
                        st.session_state['validacion_ok'] = es_valido
                        st.session_state['errores'] = errores
                        st.session_state['advertencias'] = advertencias
                        st.session_state['dataframes'] = dataframes
                        
                        if not es_valido:
                            st.error("El archivo contiene registros duplicados internamente:")
                            for msg in errores:
                                st.markdown(f"- {msg}")
                        elif advertencias:
                            st.warning(f"Validación completada con observaciones ({len(advertencias)} hojas con registros existentes):")
                            for msg in advertencias:
                                st.markdown(f"- {msg}")
                            
                            total_nuevos = sum(len(df) for df in dataframes.values()) if dataframes else 0
                            st.info(f"Total de registros nuevos a insertar: **{total_nuevos}**")
                        else:
                            st.success("Archivo válido. Todos los registros son nuevos.")
                            total_registros = sum(len(df) for df in dataframes.values()) if dataframes else 0
                            st.info(f"Total de registros a insertar: **{total_registros}**")
                            
                    except Exception as e:
                        st.error(f"Error durante la validación: {e}")
        
        with col2:
            if st.button("Cargar datos", use_container_width=True, type="primary"):
                if engine is None:
                    st.error("No se estableció conexión con la base de datos")
                    st.stop()
                
                if 'validacion_ok' not in st.session_state:
                    st.warning("Por favor, valide el archivo antes de proceder con la carga.")
                    st.stop()
                
                if not st.session_state['validacion_ok']:
                    st.error("No es posible cargar el archivo: se detectaron duplicados internos.")
                    st.stop()
                
                with st.spinner("Procesando carga de datos..."):
                    try:
                        dataframes = st.session_state.get('dataframes', {})
                        
                        if not dataframes or all(df.empty for df in dataframes.values()):
                            st.warning("No hay registros nuevos disponibles para insertar.")
                            st.stop()
                        
                        resultados = cargar_a_bd(dataframes, engine)
                        
                        st.divider()
                        st.subheader("Resumen de Ejecución")
                        
                        for tabla, resultado in resultados.items():
                            if resultado["status"] == "success":
                                st.success(f"**{tabla}**: {resultado['nuevos']} registros insertados correctamente.")
                            elif resultado["status"] == "warning":
                                st.warning(f"**{tabla}**: {resultado.get('msg', 'Sin cambios')}")
                            else:
                                st.error(f"**{tabla}**: {resultado['msg']}")
                        
                        st.success("Proceso de carga finalizado correctamente.")
                        
                        st.session_state.pop('validacion_ok', None)
                        st.session_state.pop('dataframes', None)
                        
                    except Exception as e:
                        st.error(f"Error crítico durante la carga: {e}")

if __name__ == "__main__":
    main()