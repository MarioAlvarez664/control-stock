import streamlit as st
import pandas as pd
from datetime import datetime
import os

# 1. Configuración de página (recomendado al inicio)
st.set_page_config(page_title="Control de Stock", page_icon="📦", layout="centered")

PRODUCTOS = ["Frito", "Ajo", "Tostado", "Azeite", "Alho/Salsa"]
ARCHIVO = "stock.xlsx"

# 2. Inicialización segura del archivo
if not os.path.exists(ARCHIVO):
    stock = pd.DataFrame({"Producto": PRODUCTOS, "Stock": [37, 48, 80, 2, 3]})
    movimientos = pd.DataFrame(columns=["Fecha", "Tipo", "Producto", "Cantidad"])
    
    with pd.ExcelWriter(ARCHIVO) as writer:
        stock.to_excel(writer, sheet_name="Stock", index=False)
        movimientos.to_excel(writer, sheet_name="Historial", index=False)

# Carga de datos en cada recarga de Streamlit
stock = pd.read_excel(ARCHIVO, sheet_name="Stock")
historial = pd.read_excel(ARCHIVO, sheet_name="Historial")

st.title("📦 Control de Stock")

menu = st.sidebar.radio("Navegación", ["Ver Stock", "Introducir Datos", "Historial"])

if menu == "Ver Stock":
    st.subheader("Stock Actual")
    
    # 3. Uso de columnas para un dashboard más limpio (3 columnas por fila)
    cols = st.columns(3)
    for i, fila in stock.iterrows():
        valor = fila["Stock"]
        
        # Renderizar cada producto en una columna
        with cols[i % 3]:
            if valor < 5:
                st.error(f"**{fila['Producto']}**\n\n{valor} palets")
            elif valor <= 10:
                st.warning(f"**{fila['Producto']}**\n\n{valor} palets")
            else:
                st.success(f"**{fila['Producto']}**\n\n{valor} palets")

elif menu == "Introducir Datos":
    st.subheader("Registrar Movimiento")
    
    producto = st.selectbox("Producto", PRODUCTOS)
    tipo = st.selectbox("Tipo", ["Pedido", "Producción"])
    cantidad = st.number_input("Palets", min_value=1, step=1)

    if st.button("Guardar Movimiento", type="primary"):
        indice = stock[stock["Producto"] == producto].index[0]
        stock_actual = stock.loc[indice, "Stock"]
        
        # 4. Prevención de stock negativo
        if tipo == "Pedido" and cantidad > stock_actual:
            st.error(f"⚠️ Operación denegada: Solo hay {stock_actual} palets de {producto} disponibles.")
        else:
            # Actualización de stock
            if tipo == "Pedido":
                stock.loc[indice, "Stock"] -= cantidad
            else:
                stock.loc[indice, "Stock"] += cantidad

            nuevo_movimiento = pd.DataFrame([{
                "Fecha": datetime.now(),
                "Tipo": tipo,
                "Producto": producto,
                "Cantidad": cantidad
            }])

            historial = pd.concat([historial, nuevo_movimiento], ignore_index=True)

            # Escritura en Excel
            with pd.ExcelWriter(ARCHIVO) as writer:
                stock.to_excel(writer, sheet_name="Stock", index=False)
                historial.to_excel(writer, sheet_name="Historial", index=False)

            st.success(f"✅ Se han registrado {cantidad} palets de {producto} ({tipo}).")

elif menu == "Historial":
    st.subheader("Historial de Movimientos")
    
    # 5. Formateo de fechas para que la tabla sea más legible
    historial_display = historial.copy()
    if not historial_display.empty:
        historial_display["Fecha"] = pd.to_datetime(historial_display["Fecha"]).dt.strftime('%d-%m-%Y %H:%M')
        
    st.dataframe(historial_display.sort_index(ascending=False), use_container_width=True)

    with open(ARCHIVO, "rb") as f:
        # 6. Añadido el MIME type correcto para Excel
        st.download_button(
            label="📥 Exportar a Excel",
            data=f,
            file_name=f"stock_exportado_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
