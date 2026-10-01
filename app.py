import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="Carga Rappi - Asientos Contables", layout="centered")
st.title("📊 Carga de Datos Rappi")
st.markdown("Sube el archivo original de Rappi. El sistema generará el Excel con detalle, resumen y asiento contable.")

uploaded_file = st.file_uploader("Seleccionar archivo Excel de Rappi", type=['xlsx', 'xls'])

MAPEO_CODIGOS = {
    "Fridays Comas": "0000000016",
    "Fridays Mall del Sur": "0000000012",
    "Fridays Óvalo Gutiérrez": "0000000001",
    "Fridays Ovalo Gutierrez": "0000000001",
    "Fridays Ovalo Monitor": "0000000019",
    "Fridays Óvalo Monitor": "0000000019",
    "Fridays Plaza Lima Norte": "0000000017",
    "Fridays Primavera": "0000000009",
    "Fridays Rambla": "0000000010",
    "Fridays Salaverry": "0000000008",
    "Fridays San Miguel": "0000000003",
    "Fridays Santa Anita": "0000000015",
    "Fridays DK La Molina": "0000000018",
    "Fridays Dk La Molina": "0000000018",
    "Fridays Puruchuco": "0000000014",
    "Fridays San Juan de Lurigancho": "0000000020",
    "Fridays Piura": "0000000024",
    "Fridays - Real Plaza Arequipa": "0000000006",
    "Fridays - Mall Aventura Porongoche": "0000000023",
    "Fridays Mall Aventura Porongoche": "0000000023",
    "Fridays El Polo": "0000000027",
    "Fridays - Cusco": "0000000025",
    "Fridays Piura - Real Plaza": "0000000024",
    "Smash Burger by Fridays Óvalo Gutiérrez": "0000000001",
    "Smash Burger by Fridays Ovalo Gutierrez": "0000000001",
    "Smash Burger by Fridays Ovalo Monitor": "0000000019",
    "Smash Burger by Fridays Óvalo Monitor": "0000000019",
    "Smash Burger by Fridays San Juan de Lurigancho": "0000000020",
    "Smash Burger by Fridays Dk la Molina": "0000000018",
    "Smash Burger by Fridays DK La Molina": "0000000018",
    "Smash Burger by Fridays Santa Anita": "0000000015",
    "Smash Burger by Fridays Piura": "0000000024",
    "Smash Burger by Fridays Plaza Lima Norte": "0000000017",
    "Smash Burger by Fridays San Miguel": "0000000003",
    "Smash Burger By Fridays - Mall Aventura Porongoche": "0000000023",
    "Smash Burger by Fridays Mall Aventura Porongoche": "0000000023",
    "Smash Burger by Fridays Salaverry": "0000000008",
    "Smash Burger by Fridays Mall del Sur": "0000000012",
    "Smash Burger by Fridays Rambla": "0000000010",
    "Smash Burger by Fridays Comas": "0000000016",
    "Smash Burger by Fridays Puruchuco": "0000000014",
    "Smash Burger by Fridays Arequipa": "0000000006",
    "Smash Burger by Fridays Primavera": "0000000009",
    "Smash Burger by Fridays El Polo": "0000000027",
    "Smash Burger by Fridays Cusco": "0000000025"
}

def obtener_codigo_tienda(nombre_tienda):
    if nombre_tienda in MAPEO_CODIGOS:
        return MAPEO_CODIGOS[nombre_tienda]
    nombre_normalizado = nombre_tienda.strip().lower()
    for nombre_mapeado, codigo in MAPEO_CODIGOS.items():
        if nombre_mapeado.lower() == nombre_normalizado:
            return codigo
    return ""

def aplicar_formato_detalle(ws):
    fill_header = PatternFill(start_color="F4B183", end_color="F4B183", fill_type="solid")
    font_header = Font(name='Calibri', bold=True, size=11, color="000000")
    font_normal = Font(name='Calibri', size=10)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for cell in ws[1]:
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.font = font_normal
            cell.border = thin_border
            cell.alignment = Alignment(vertical='center')
            if cell.column > 5:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal='right', vertical='center')
    anchos = {'A': 22, 'B': 18, 'C': 35, 'D': 18, 'E': 20, 'F': 25, 'G': 18, 'H': 18, 'I': 25, 'J': 22}
    for col, ancho in anchos.items():
        ws.column_dimensions[col].width = ancho
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions

def aplicar_formato_resumen(ws):
    fill_header = PatternFill(start_color="F4B183", end_color="F4B183", fill_type="solid")
    fill_gral = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    font_header = Font(name='Calibri', bold=True, size=11, color="000000")
    font_normal = Font(name='Calibri', size=10)
    font_total = Font(name='Calibri', bold=True, size=10)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for cell in ws[1]:
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    max_row = ws.max_row
    for row_idx in range(2, max_row + 1):
        for cell in ws[row_idx]:
            cell.font = font_normal
            cell.border = thin_border
            cell.alignment = Alignment(vertical='center')
            if cell.column > 1:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal='right', vertical='center')
        if row_idx == max_row:
            for cell in ws[row_idx]:
                cell.fill = fill_gral
                cell.font = font_total
    anchos = {'A': 35, 'B': 18, 'C': 22, 'D': 18, 'E': 22, 'F': 22, 'G': 22, 'H': 18}
    for col, ancho in anchos.items():
        ws.column_dimensions[col].width = ancho
    ws.freeze_panes = 'A2'

def aplicar_formato_asiento(ws):
    fill_header = PatternFill(start_color="F4B183", end_color="F4B183", fill_type="solid")
    fill_gral = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    font_header = Font(name='Calibri', bold=True, size=11, color="000000")
    font_normal = Font(name='Calibri', size=10)
    font_total = Font(name='Calibri', bold=True, size=10)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for cell in ws[1]:
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    max_row = ws.max_row
    for row_idx in range(2, max_row + 1):
        for cell in ws[row_idx]:
            cell.font = font_normal
            cell.border = thin_border
            cell.alignment = Alignment(vertical='center')
            if cell.column in [2, 4, 5]:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif cell.column in [3, 6]:
                cell.alignment = Alignment(horizontal='center', vertical='center')
        if row_idx == max_row:
            for cell in ws[row_idx]:
                cell.fill = fill_gral
                cell.font = font_total
    anchos = {'A': 45, 'B': 18, 'C': 20, 'D': 18, 'E': 18, 'F': 15}
    for col, ancho in anchos.items():
        ws.column_dimensions[col].width = ancho
    ws.freeze_panes = 'A2'

if uploaded_file:
    st.info(f"Archivo cargado: **{uploaded_file.name}**")
    
    if st.button("Procesar y Generar Excel", use_container_width=True, type="primary"):
        with st.spinner("Procesando datos..."):
            try:
                df_completo = pd.read_excel(uploaded_file, sheet_name="Detalle", header=1)
                df_completo.columns = df_completo.columns.str.strip()
                
                columnas_necesarias = [
                    "Fecha de creación orden", "ID de la órden", "Nombre de la tienda",
                    "Estado de la órden", "Tipo de transacción",
                    "Ventas base por Uso y alquiler de plataforma Rappi (informativo)",
                    "Compensaciones", "Costo Canceladas",
                    "Uso y alquiler de plataforma Rappi", "Valor Ajustes Manuales"
                ]
                
                columnas_existentes = [col for col in columnas_necesarias if col in df_completo.columns]
                df_filtrado = df_completo[columnas_existentes].copy()
                
                resumen = df_filtrado.groupby("Nombre de la tienda").agg({
                    "Ventas base por Uso y alquiler de plataforma Rappi (informativo)": "sum",
                    "Valor Ajustes Manuales": "sum",
                    "Uso y alquiler de plataforma Rappi": "sum",
                    "Compensaciones": "sum",
                    "Costo Canceladas": "sum"
                }).reset_index()
                
                resumen["IGV Comisión (18%)"] = resumen["Uso y alquiler de plataforma Rappi"] * 0.18
                resumen["Total"] = (
                    resumen["Ventas base por Uso y alquiler de plataforma Rappi (informativo)"] +
                    resumen["Valor Ajustes Manuales"] +
                    resumen["Uso y alquiler de plataforma Rappi"] +
                    resumen["IGV Comisión (18%)"] +
                    resumen["Compensaciones"] +
                    resumen["Costo Canceladas"]
                )
                
                resumen.columns = [
                    "Nombre de la tienda", "Suma de VENTAS", "Suma de Valor Ajustes Manuales",
                    "Suma de Comisión", "Suma de Compensaciones", "Suma de Costo Canceladas",
                    "Suma de IVA Comisión (18%)", "Total"
                ]
                
                cols_order = ["Nombre de la tienda", "Suma de VENTAS", "Suma de Valor Ajustes Manuales", 
                             "Suma de Comisión", "Suma de IVA Comisión (18%)", 
                             "Suma de Compensaciones", "Suma de Costo Canceladas", "Total"]
                resumen = resumen[cols_order]
                
                for col in resumen.columns[1:]:
                    resumen[col] = resumen[col].round(2)
                
                fila_total = resumen.iloc[:, 1:].sum()
                fila_total["Nombre de la tienda"] = "TOTAL GENERAL"
                resumen = pd.concat([resumen, pd.DataFrame([fila_total])], ignore_index=True)
                
                asiento = pd.DataFrame()
                asiento["Tiendas"] = resumen["Nombre de la tienda"].copy()
                asiento["Suma de Comision"] = resumen["Suma de Comisión"].copy()
                asiento["Suma de Descuento por inversión de Rappi a aplicar sobre Uso y alquiler de plataforma Rappi DAR"] = "-"
                asiento["Comision_Abs_1"] = resumen["Suma de Comisión"].abs()
                asiento["Comision_Abs_2"] = resumen["Suma de Comisión"].abs()
                asiento["Code"] = asiento["Tiendas"].apply(obtener_codigo_tienda)
                
                asiento["Suma de Comision"] = asiento["Suma de Comision"].round(2)
                asiento["Comision_Abs_1"] = asiento["Comision_Abs_1"].round(2)
                asiento["Comision_Abs_2"] = asiento["Comision_Abs_2"].round(2)
                
                excel_buffer = io.BytesIO()
                
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_filtrado.to_excel(writer, index=False, sheet_name='Detalle')
                    resumen.to_excel(writer, index=False, sheet_name='Resumen')
                    asiento.to_excel(writer, index=False, sheet_name='Asiento')
                    
                    aplicar_formato_detalle(writer.sheets['Detalle'])
                    aplicar_formato_resumen(writer.sheets['Resumen'])
                    aplicar_formato_asiento(writer.sheets['Asiento'])
                
                st.success(f"✅ Archivo procesado. {len(df_filtrado)} registros en Detalle, {len(resumen)-1} tiendas en Resumen y Asiento.")
                
                st.download_button(
                    label="⬇️ Descargar Excel con Detalle, Resumen y Asiento",
                    data=excel_buffer.getvalue(),
                    file_name="rappi_asientos_contables.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"Error: {e}")
                import traceback
                st.code(traceback.format_exc())
