import streamlit as st
import pandas as pd
from datetime import datetime
import os

st.set_page_config(page_title="Control de Stock", page_icon="📦", layout="centered")

# 1. Configuración de los dos sectores
CONFIG = {
    "Picatoste": {
        "archivo": "stock.xlsx",  # Mantiene tu archivo actual para no perder datos
        "productos": ["Frito", "Ajo", "Tostado", "Azeite", "Alho/Salsa"],
        "stock_inicial": [37, 48, 80, 2, 3],
        "demanda_diaria": {
            "Frito": 3.98,
            "Ajo": 3.22,
            "Tostado": 3.94,
            "Azeite": 0.23,
            "Alho/Salsa": 0.40
        }
    },
    "Pan Rayado": {
        "archivo": "stock_pan_rayado.xlsx", # Crea un archivo nuevo independiente
        "productos": ["Box", "Cajas", "A+P", "Casero", "Crujiente"],
        "stock_inicial": [121, 47, 72, 21, 31],
        "demanda_diaria": {
            "Box": 0,         # 0 indica que no hay demanda diaria predecible
            "Cajas": 0,
            "A+P": 11.39,
            "Casero": 7.80,
            "Crujiente": 2.91
        }
    }
}

# 2. Selector de sector en la barra lateral
st.sidebar.title("🏢 Sector")
sector = st.sidebar.selectbox("Elige la línea de producción:", ["Picatoste", "Pan Rayado"])

st.sidebar.markdown("---")

# 3. Extraemos las variables según el sector elegido
datos_sector = CONFIG[sector]
ARCHIVO = datos_sector["archivo"]
PRODUCTOS = datos_sector["productos"]
STOCK_INICIAL = datos_sector["stock_inicial"]
DEMANDA_DIARIA = datos_sector["demanda_diaria"]

# 4. Inicialización del archivo si no existe
if not os.path.exists(ARCHIVO):
    stock = pd.DataFrame({
        "Producto": PRODUCTOS, 
        "Stock": STOCK_INICIAL,
        "Última Actualización": ["Sin registros"] * len(PRODUCTOS)
    })
    movimientos = pd.DataFrame(columns=["Fecha", "Tipo", "Producto", "Cantidad"])
    
    with pd.ExcelWriter(ARCHIVO) as writer:
        stock.to_excel(writer, sheet_name="Stock", index=False)
        movimientos.to_excel(writer, sheet_name="Historial", index=False)

# Carga de datos
stock = pd.read_excel(ARCHIVO, sheet_name="Stock")
historial = pd.read_excel(ARCHIVO, sheet_name="Historial")

# Parche de seguridad para compatibilidad con versiones anteriores
if "Última Actualización" not in stock.columns:
    stock["Última Actualización"] = "Sin registros"

st.title(f"📦 Control de Stock - {sector}")

# 5. Menú de navegación
menu = st.sidebar.radio("Navegación", ["Ver Stock", "Actualizar Stock (Masivo)", "Historial"])

if menu == "Ver Stock":
    st.subheader("Stock Actual y Días Restantes")
    
    cols = st.columns(3)
    for i, fila in stock.iterrows():
        prod = fila["Producto"]
        valor = fila["Stock"]
        ultima_fecha = fila["Última Actualización"]
        
        demanda = DEMANDA_DIARIA.get(prod, 0)
        
        # 6. Lógica inteligente para productos con y sin demanda diaria
        if demanda > 0:
            dias_restantes = valor / demanda
            texto_dias = f"⏳ **{dias_restantes:.1f} días** de stock"
            
            # Alertas visuales según días de stock
            if dias_restantes < 3:
                estado = "error"
            elif dias_restantes <= 7:
                estado = "warning"
            else:
                estado = "success"
        else:
            texto_dias = "⏳ **Sin consumo diario**"
            # Si no hay demanda, se marca en rojo solo si nos quedamos a cero
            if valor <= 0:
                estado = "error"
            else:
                estado = "success"
        
        texto_tarjeta = (
            f"**{prod}**\n\n"
            f"📦 **{valor}** palets\n\n"
            f"{texto_dias}\n\n"
            f"🕒 *{ultima_fecha}*"
        )
        
        # Renderizado de tarjetas según el estado
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
