import streamlit as st
import pandas as pd
from datetime import datetime
import os

st.set_page_config(page_title="Control de Stock", page_icon="📦", layout="centered")

PRODUCTOS = ["Frito", "Ajo", "Tostado", "Azeite", "Alho/Salsa"]

# Demanda media diaria por producto (en palets/día)
DEMANDA_DIARIA = {
    "Frito": 3.98,
    "Ajo": 3.22,
    "Tostado": 3.94,
    "Azeite": 0.23,
    "Alho/Salsa": 0.4
}

ARCHIVO = "stock.xlsx"

# 1. Inicialización del archivo si no existe
if not os.path.exists(ARCHIVO):
    stock = pd.DataFrame({
        "Producto": PRODUCTOS, 
        "Stock": [37, 48, 80, 2, 3],
        "Última Actualización": ["Sin registros"] * len(PRODUCTOS)
    })
    movimientos = pd.DataFrame(columns=["Fecha", "Tipo", "Producto", "Cantidad"])
    
    with pd.ExcelWriter(ARCHIVO) as writer:
        stock.to_excel(writer, sheet_name="Stock", index=False)
        movimientos.to_excel(writer, sheet_name="Historial", index=False)

# Carga de datos
stock = pd.read_excel(ARCHIVO, sheet_name="Stock")
historial = pd.read_excel(ARCHIVO, sheet_name="Historial")

# Parche de compatibilidad por si la columna no existe en el Excel previo
if "Última Actualización" not in stock.columns:
    stock["Última Actualización"] = "Sin registros"

st.title("📦 Control de Stock")

menu = st.sidebar.radio("Navegación", ["Ver Stock", "Actualizar Stock (Masivo)", "Historial"])

if menu == "Ver Stock":
    st.subheader("Stock Actual y Días de Cobertura")
    
    cols = st.columns(3)
    for i, fila in stock.iterrows():
        prod = fila["Producto"]
        valor = fila["Stock"]
        ultima_fecha = fila["Última Actualización"]
        
        # 2. Cálculo de días de stock restantes
        demanda = DEMANDA_DIARIA.get(prod, 0)
        dias_restantes = round(valor / demanda, 1) if demanda > 0 else 0
        
        # Formato de la tarjeta visual
        texto_tarjeta = (
            f"**{prod}**\n\n"
            f"📦 **{valor}** palets\n\n"
            f"⏳ **{dias_restantes} días** de stock\n\n"
            f"🕒 *{ultima_fecha}*"
        )
        
        with cols[i % 3]:
            # 3. Alertas basadas en días de stock restantes (no en palets brutos)
            if dias_restantes < 3:
                st.error(texto_tarjeta)
            elif dias_restantes <= 7:
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
            
            # Validación de stock suficiente
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
    st.subheader("Historial de Movimientos")
    
    historial_display = historial.copy()
    if not historial_display.empty:
        historial_display["Fecha"] = pd.to_datetime(historial_display["Fecha"]).dt.strftime('%d-%m-%Y %H:%M')
        
    st.dataframe(historial_display.sort_index(ascending=False), use_container_width=True)

    with open(ARCHIVO, "rb") as f:
        st.download_button(
            label="📥 Exportar a Excel",
            data=f,
            file_name=f"stock_exportado_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
