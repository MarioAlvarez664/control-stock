import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import io

st.set_page_config(page_title="Control de Stock", page_icon="📦", layout="centered")

# 1. Configuración de los sectores
CONFIG = {
    "Picatoste": {
        "archivo": "stock.xlsx",
        "productos": ["Frito", "Ajo", "Tostado", "Azeite", "Alho/Salsa"],
        "stock_inicial": [37, 48, 80, 2, 3],
        "demanda_diaria": {
            "Frito": 3.98,
            "Ajo": 3.22,
            "Tostado": 3.94,
            "Azeite": 0.23,
            "Alho/Salsa": 0.40
        },
        "ventas_anuales": {
            "Frito": 3679261,
            "Ajo": 2702671,
            "Tostado": 2769737,
            "Azeite": 175712,
            "Alho/Salsa": 276146
        },
        "factores_bolsas": {
            "Frito": 2880,
            "Ajo": 2880,
            "Tostado": 2400,
            "Azeite": 2304,
            "Alho/Salsa": 2304
        }
    },
    "Pan Rayado": {
        "archivo": "stock_pan_rayado.xlsx",
        "productos": ["Box", "Cajas", "A+P", "Casero", "Crujiente"],
        "stock_inicial": [121, 47, 72, 21, 31],
        "demanda_diaria": {
            "Box": 0,
            "Cajas": 0,
            "A+P": 11.39,
            "Casero": 7.80,
            "Crujiente": 2.91
        }
    }
}

# Función para calcular la fecha de fin de stock excluyendo los domingos
def calcular_fecha_agotamiento(stock_val, demanda):
    if demanda <= 0:
        return "Indefinido"
    if stock_val <= 0:
        return "Agotado"
    
    current_date = datetime.now()
    stock_restante = float(stock_val)
    
    while stock_restante > 0:
        current_date += timedelta(days=1)
        if current_date.weekday() != 6:  # Excluir domingos
            stock_restante -= demanda
            
    return current_date.strftime('%d-%m-%Y')

# 2. Selector de sector en la barra lateral
st.sidebar.title("🏢 Sector Activo")
sector = st.sidebar.selectbox("Elige la línea:", ["Picatoste", "Pan Rayado"])

st.sidebar.markdown("---")

# 3. Opciones de menú según el sector
if sector == "Picatoste":
    menu_opciones = ["Ver Stock", "Actualizar Stock (Masivo)", "Historial", "Objetivos Mensuales", "Resumen Global"]
else:
    menu_opciones = ["Ver Stock", "Actualizar Stock (Masivo)", "Historial", "Resumen Global"]

menu = st.sidebar.radio("Navegación", menu_opciones)

# Si NO estamos en Resumen Global, operamos sobre el sector activo
if menu != "Resumen Global":
    datos_sector = CONFIG[sector]
    ARCHIVO = datos_sector["archivo"]
    PRODUCTOS = datos_sector["productos"]
    STOCK_INICIAL = datos_sector["stock_inicial"]
    DEMANDA_DIARIA = datos_sector["demanda_diaria"]

    # Inicialización del archivo si no existe
    if not os.path.exists(ARCHIVO):
        stock_df = pd.DataFrame({
            "Producto": PRODUCTOS, 
            "Stock": STOCK_INICIAL,
            "Última Actualización": ["Sin registros"] * len(PRODUCTOS)
        })
        movimientos_df = pd.DataFrame(columns=["Fecha", "Tipo", "Producto", "Cantidad"])
        
        with pd.ExcelWriter(ARCHIVO) as writer:
            stock_df.to_excel(writer, sheet_name="Stock", index=False)
            movimientos_df.to_excel(writer, sheet_name="Historial", index=False)

    # Carga de datos del sector activo
    stock = pd.read_excel(ARCHIVO, sheet_name="Stock")
    historial = pd.read_excel(ARCHIVO, sheet_name="Historial")

    # Parche de seguridad para compatibilidad
    if "Última Actualización" not in stock.columns:
        stock["Última Actualización"] = "Sin registros"

    st.title(f"📦 Control de Stock - {sector}")

    if menu == "Ver Stock":
        st.subheader("Stock Actual, Días y Fecha Límite (Sin contar domingos)")
        
        cols = st.columns(3)
        for i, fila in stock.iterrows():
            prod = fila["Producto"]
            valor = fila["Stock"]
            ultima_fecha = fila["Última Actualización"]
            
            demanda = DEMANDA_DIARIA.get(prod, 0)
            
            if demanda > 0:
                dias_restantes = valor / demanda
                fecha_fin = calcular_fecha_agotamiento(valor, demanda)
                texto_dias = f"⏳ **{dias_restantes:.1f} días**\n📅 *Hasta el {fecha_fin}*"
                
                if dias_restantes < 3:
                    estado = "error"
                elif dias_restantes <= 7:
                    estado = "warning"
                else:
                    estado = "success"
            else:
                texto_dias = "⏳ **Sin consumo diario**"
                if valor <= 0:
                    estado = "error"
                else:
                    estado = "success"
            
            texto_tarjeta = (
                f"**{prod}**\n\n"
                f"📦 **{valor}** palets\n\n"
                f"{texto_dias}\n\n"
                f"🕒 *Act: {ultima_fecha}*"
            )
            
            with cols[i % 3]:
                if estado == "error":
                    st.error(texto_tarjeta)
                elif estado == "warning":
                    st.warning(texto_tarjeta)
                else:
                    st.success(texto_tarjeta)

    elif menu == "Actualizar Stock (Masivo)":
        st.subheader("Registrar Movimientos")
        st.write("Escribe las cantidades a **añadir** (Producción) o **quitar** (Pedido) en la tabla:")
        
        df_input = pd.DataFrame({
            "Producto": stock["Producto"],
            "Stock Actual": stock["Stock"],
            "➕ Añadir": [0] * len(stock),
            "➖ Quitar": [0] * len(stock)
        })
        
        editado = st.data_editor(
            df_input,
            hide_index=True,
            use_container_width=True,
            disabled=["Producto", "Stock Actual"],
            column_config={
                "➕ Añadir": st.column_config.NumberColumn(min_value=0, step=1),
                "➖ Quitar": st.column_config.NumberColumn(min_value=0, step=1)
            }
        )
        
        if st.button("Guardar todos los cambios", type="primary"):
            cambios_realizados = False
            errores = False
            nuevos_movimientos = []
            ahora_str = datetime.now().strftime('%d-%m-%Y %H:%M')
            ahora_dt = datetime.now()
            
            for index, row in editado.iterrows():
                prod = row["Producto"]
                add_qty = row["➕ Añadir"]
                sub_qty = row["➖ Quitar"]
                stock_actual = stock.loc[index, "Stock"]
                
                if sub_qty > (stock_actual + add_qty):
                    st.error(f"⚠️ No puedes quitar {sub_qty} palets de {prod}. Solo hay {stock_actual}.")
                    errores = True
                    continue
                
                if add_qty > 0 or sub_qty > 0:
                    cambios_realizados = True
                    stock.loc[index, "Stock"] = stock_actual + add_qty - sub_qty
                    stock.loc[index, "Última Actualización"] = ahora_str
                    
                    if add_qty > 0:
                        nuevos_movimientos.append({"Fecha": ahora_dt, "Tipo": "Producción", "Producto": prod, "Cantidad": add_qty})
                    if sub_qty > 0:
                        nuevos_movimientos.append({"Fecha": ahora_dt, "Tipo": "Pedido", "Producto": prod, "Cantidad": sub_qty})
            
            if cambios_realizados and not errores:
                df_nuevos = pd.DataFrame(nuevos_movimientos)
                historial = pd.concat([historial, df_nuevos], ignore_index=True)
                
                with pd.ExcelWriter(ARCHIVO) as writer:
                    stock.to_excel(writer, sheet_name="Stock", index=False)
                    historial.to_excel(writer, sheet_name="Historial", index=False)
                
                st.success("✅ Cambios guardados correctamente.")
                st.rerun()

    elif menu == "Historial":
        st.subheader(f"Historial de Movimientos - {sector}")
        
        historial_display = historial.copy()
        if not historial_display.empty:
            historial_display["Fecha"] = pd.to_datetime(historial_display["Fecha"]).dt.strftime('%d-%m-%Y %H:%M')
            
        st.dataframe(historial_display.sort_index(ascending=False), use_container_width=True)

        with open(ARCHIVO, "rb") as f:
            st.download_button(
                label=f"📥 Exportar Excel de {sector}",
                data=f,
                file_name=f"stock_{sector.lower().replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    elif menu == "Objetivos Mensuales":
        st.subheader("🎯 Objetivos Mensuales - Picatoste")
        st.write("Control de cumplimiento de la proyección mensual (Ventas Anuales + 5% / 12) frente a lo vendido (Pedidos) en el mes en curso.")
        
        # Calcular palets vendidos en el mes actual a partir del historial
        if not historial.empty and "Fecha" in historial.columns:
            historial['Fecha_dt'] = pd.to_datetime(historial['Fecha'], errors='coerce')
            now = datetime.now()
            current_month_hist = historial[
                (historial['Fecha_dt'].dt.month == now.month) & 
                (historial['Fecha_dt'].dt.year == now.year) & 
                (historial['Tipo'] == 'Pedido')
            ]
            sold_palets = current_month_hist.groupby('Producto')['Cantidad'].sum().to_dict()
        else:
            sold_palets = {}
            
        ventas_anuales = CONFIG["Picatoste"]["ventas_anuales"]
        factores = CONFIG["Picatoste"]["factores_bolsas"]
        productos = CONFIG["Picatoste"]["productos"]
        
        objetivos_data = []
        for prod in productos:
            v_anual = ventas_anuales[prod]
            v_anual_5 = v_anual * 1.05
            proj_mensual_bags = v_anual_5 / 12
            factor = factores[prod]
            proj_mensual_palets = proj_mensual_bags / factor
            
            palets_vendidos = sold_palets.get(prod, 0)
            bags_vendidas = palets_vendidos * factor
            
            pct = (bags_vendidas / proj_mensual_bags) * 100 if proj_mensual_bags > 0 else 0
            
            objetivos_data.append({
                "Producto": prod,
                "Proyección Mensual (Bolsas)": round(proj_mensual_bags, 1),
                "Proyección Mensual (Palets)": round(proj_mensual_palets, 1),
                "Vendido este Mes (Bolsas)": round(bags_vendidas, 1),
                "Vendido este Mes (Palets)": palets_vendidos,
                "% Completado": round(pct, 1)
            })
            
            st.markdown(f"### 📦 {prod}")
            col_m1, col_m2, col_m3 = st.columns(3)
            col_m1.metric("Proyección Mensual", f"{proj_mensual_bags:,.0f} bolsas", f"{proj_mensual_palets:.1f} palets")
            col_m2.metric("Vendido este Mes", f"{bags_vendidas:,.0f} bolsas", f"{palets_vendidos} palets")
            col_m3.metric("% Completado", f"{pct:.1f}%")
            
            progress_val = min(pct / 100.0, 1.0)
            st.progress(progress_val)
            st.markdown("---")
            
        df_obj = pd.DataFrame(objetivos_data)
        st.subheader("📋 Tabla Resumen de Objetivos")
        st.dataframe(df_obj, use_container_width=True, hide_index=True)

# Vista de Resumen Global (Ambos sectores juntos)
else:
    st.title("🌐 Resumen Global (Picatoste y Pan Rayado)")
    st.write("Vista consolidada con el stock actual, demanda diaria, días restantes y fecha estimada de fin de stock (excluyendo domingos).")
    
    consolidated_rows = []
    for sec, cfg in CONFIG.items():
        f_path = cfg["archivo"]
        if os.path.exists(f_path):
            df_sec = pd.read_excel(f_path, sheet_name="Stock")
        else:
            df_sec = pd.DataFrame({
                "Producto": cfg["productos"],
                "Stock": cfg["stock_inicial"],
                "Última Actualización": ["Sin registros"] * len(cfg["productos"])
            })
        
        demanda_dict = cfg["demanda_diaria"]
        for _, row in df_sec.iterrows():
            prod = row["Producto"]
            stock_val = row["Stock"]
            ultima = row["Última Actualización"]
            demanda = demanda_dict.get(prod, 0)
            
            if demanda > 0:
                dias_val = round(stock_val / demanda, 1)
                fecha_fin = calcular_fecha_agotamiento(stock_val, demanda)
                demanda_str = f"{demanda} palets/día"
            else:
                dias_val = "N/A"
                fecha_fin = "Indefinido"
                demanda_str = "Sin consumo diario"
            
            consolidated_rows.append({
                "Sector": sec,
                "Producto": prod,
                "Stock Actual (Palets)": stock_val,
                "Demanda Diaria": demanda_str,
                "Días de Stock": dias_val,
                "Fecha Fin de Stock (Sin Domingos)": fecha_fin,
                "Última Actualización": ultima
            })
    
    df_global = pd.DataFrame(consolidated_rows)
    
    st.dataframe(df_global, use_container_width=True, hide_index=True)
    
    # Generación del archivo Excel en memoria para descarga global
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_global.to_excel(writer, sheet_name="Resumen Global", index=False)
    excel_data = output.getvalue()
    
    st.markdown("---")
    st.download_button(
        label="📥 Descargar Excel Global Consolidado",
        data=excel_data,
        file_name=f"stock_global_consolidado_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
