# -*- coding: utf-8 -*-
"""
Created on Tue Jul 29 14:38:44 2025

@author: abell
"""


# -*- coding: utf-8 -*-
import pandas as pd

# Cargar el dataset
df = pd.read_csv(r"C:\Users\abell\Downloads\Opiniones_Comerciales_Extendidas.csv")

# ===========================
# MÉTRICAS AGREGADAS
# ===========================

# Crear columna auxiliar para cálculos
df['negativa'] = df['prediccion'].apply(lambda x: 1 if x == 'NEG' else 0)
df['positiva'] = df['prediccion'].apply(lambda x: 1 if x == 'POS' else 0)
df['neutra'] = df['prediccion'].apply(lambda x: 1 if x == 'NEU' else 0)

# ===========================
# RESUMEN POR PRODUCTO
# ===========================
resumen_producto = df.groupby('producto_id').agg(
    total_opiniones=('prediccion', 'count'),
    negativas=('negativa', 'sum'),
    positivas=('positiva', 'sum'),
    neutras=('neutra', 'sum')
)
resumen_producto['% NEG'] = resumen_producto['negativas'] / resumen_producto['total_opiniones']
resumen_producto['% POS'] = resumen_producto['positivas'] / resumen_producto['total_opiniones']
resumen_producto['recomendacion'] = resumen_producto['% NEG'].apply(
    lambda x: '⚠ REVISAR producto: alta negatividad' if x > 0.30 else '✅ OK'
)

# ===========================
# RESUMEN POR PAÍS
# ===========================
resumen_pais = df.groupby('pais').agg(
    total_opiniones=('prediccion', 'count'),
    negativas=('negativa', 'sum'),
    positivas=('positiva', 'sum'),
    neutras=('neutra', 'sum')
)
resumen_pais['% NEG'] = resumen_pais['negativas'] / resumen_pais['total_opiniones']
resumen_pais['recomendacion'] = resumen_pais['% NEG'].apply(
    lambda x: '⚠ REVISAR país: muchas quejas' if x > 0.30 else '✅ OK'
)

# ===========================
# RESUMEN POR CANAL
# ===========================
resumen_canal = df.groupby('canal').agg(
    total_opiniones=('prediccion', 'count'),
    negativas=('negativa', 'sum'),
    positivas=('positiva', 'sum'),
    neutras=('neutra', 'sum')
)
resumen_canal['% NEG'] = resumen_canal['negativas'] / resumen_canal['total_opiniones']
resumen_canal['recomendacion'] = resumen_canal['% NEG'].apply(
    lambda x: '⚠ MEJORAR canal: mal feedback' if x > 0.30 else '✅ OK'
)

# ===========================
# TOP PRODUCTOS CON MÁS QUEJAS
# ===========================
top_quejas = df[df['prediccion'] == 'NEG'].groupby('producto_id').size().sort_values(ascending=False).head(5)
top_quejas = top_quejas.reset_index().rename(columns={0: 'quejas_totales'})

# ===========================
# TENDENCIA POR MES
# ===========================
df['fecha'] = pd.to_datetime(df['fecha'])
df['mes'] = df['fecha'].dt.to_period('M')
tendencia_mensual = df.groupby(['mes', 'prediccion']).size().unstack().fillna(0)

# ===========================
# EXPORTAR A EXCEL
# ===========================
with pd.ExcelWriter(r"C:\Users\abell\Downloads\reporte_optimizacion_extendido.xlsx") as writer:
    resumen_producto.to_excel(writer, sheet_name='Resumen Producto')
    resumen_pais.to_excel(writer, sheet_name='Resumen País')
    resumen_canal.to_excel(writer, sheet_name='Resumen Canal')
    top_quejas.to_excel(writer, sheet_name='Top Quejas')
    tendencia_mensual.to_excel(writer, sheet_name='Tendencia Mensual')

print("✅ Reporte extendido generado: reporte_optimizacion_extendido.xlsx")
