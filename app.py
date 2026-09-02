import streamlit as st
import pandas as pd
from datetime import datetime
import os

st.set_page_config(page_title="Control de Stock", page_icon="📦", layout="centered")

PRODUCTOS = ["Frito", "Ajo", "Tostado", "Azeite", "Alho/Salsa"]
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

# PARCHE: Si el Excel antiguo no tiene la columna de "Última Actualización", se la añadimos
if "Última Actualización" not in stock.columns:
    stock["Última Actualización"] = "Sin registros"

st.title("📦 Control de Stock")

menu = st.sidebar.radio("Navegación", ["Ver Stock", "Actualizar Stock (Masivo)", "Historial"])

if menu == "Ver Stock":
    st.subheader("Stock Actual")
    
    cols = st.columns(3)
    for i, fila in stock.iterrows():
        valor = fila["Stock"]
        ultima_fecha = fila["Última Actualización"]
        
        # 2. Formato del texto de la tarjeta, incluyendo la fecha
        texto_tarjeta = f"**{fila['Producto']}**\n\n📦 {valor} palets\n\n🕒 *{ultima_fecha}*"
        
        with cols[i % 3]:
            if valor < 5:
                st.error(texto_tarjeta)
            elif valor <= 10:
                st.warning(texto_tarjeta)
            else:
                st.success(texto_tarjeta)

elif menu == "Actualizar Stock (Masivo)":
    st.subheader("Registrar Movimientos")
    st.write("Escribe las cantidades a **añadir** (Producción) o **quitar** (Pedido) en la tabla:")
    
    # 3. Preparamos una tabla temporal para que el usuario edite
    df_input = pd.DataFrame({
        "Producto": stock["Producto"],
        "Stock Actual": stock["Stock"],
        "➕ Añadir": [0] * len(stock),
        "➖ Quitar": [0] * len(stock)
    })
    
    # 4. Tabla editable
    editado = st.data_editor(
        df_input,
        hide_index=True,
        use_container_width=True,
        disabled=["Producto", "Stock Actual"], # Bloqueamos para que no puedan cambiar nombres ni el stock base
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
            
            # Evitar números negativos en el stock
            if sub_qty > (stock_actual + add_qty):
                st.error(f"⚠️ No puedes quitar {sub_qty} palets de {prod}. Solo hay {stock_actual}.")
                errores = True
                continue
            
            # Si el usuario puso algún número en Añadir o Quitar
            if add_qty > 0 or sub_qty > 0:
                cambios_realizados = True
                
                # Actualizamos stock y la fecha de última actualización
                stock.loc[index, "Stock"] = stock_actual + add_qty - sub_qty
                stock.loc[index, "Última Actualización"] = ahora_str
                
                if add_qty > 0:
                    nuevos_movimientos.append({"Fecha": ahora_dt, "Tipo": "Producción", "Producto": prod, "Cantidad": add_qty})
                if sub_qty > 0:
                    nuevos_movimientos.append({"Fecha": ahora_dt, "Tipo": "Pedido", "Producto": prod, "Cantidad": sub_qty})
        
        # 5. Guardado global
        if cambios_realizados and not errores:
            df_nuevos = pd.DataFrame(nuevos_movimientos)
            historial = pd.concat([historial, df_nuevos], ignore_index=True)
            
            with pd.ExcelWriter(ARCHIVO) as writer:
                stock.to_excel(writer, sheet_name="Stock", index=False)
                historial.to_excel(writer, sheet_name="Historial", index=False)
            
            st.success("✅ Cambios guardados correctamente.")
            st.rerun() # Recarga la página para poner la tabla a 0 de nuevo

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
