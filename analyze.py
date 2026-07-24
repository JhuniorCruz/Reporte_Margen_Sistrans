import pandas as pd
import numpy as np

# Load dataset
df_raw = pd.read_excel('rptlistadoserviciosmargen.xls', header=None)

# Extract table
df = df_raw.iloc[11:159].copy()
cols = ['servicio_ida', 'fecha_salida', 'tipo_servicio', 'ruta', 'placa', 'operador', 
        'servicio_retorno', 'flete_ida', 'flete_retorno', 'gastos_ida', 'gastos_retorno', 'margen_bruto']
df = df.iloc[:, :12]
df.columns = cols

df['servicio_ida'] = pd.to_numeric(df['servicio_ida'], errors='coerce')
df['servicio_retorno'] = pd.to_numeric(df['servicio_retorno'], errors='coerce').fillna(0)
df['fecha_salida'] = pd.to_datetime(df['fecha_salida'], errors='coerce')

num_cols = ['flete_ida', 'flete_retorno', 'gastos_ida', 'gastos_retorno', 'margen_bruto']
for col in num_cols:
    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

df['flete_total'] = df['flete_ida'] + df['flete_retorno']
df['gastos_total'] = df['gastos_ida'] + df['gastos_retorno']
df['margen_calculado'] = df['flete_total'] - df['gastos_total']
df['pct_margen'] = np.where(df['flete_total'] > 0, (df['margen_calculado'] / df['flete_total']) * 100, 0)

print('=== 1. RESUMEN GENERAL DEL ARCHIVO ===')
print(f"Total registros de servicios: {len(df)}")
print(f"Fecha mínima: {df['fecha_salida'].min()}")
print(f"Fecha máxima: {df['fecha_salida'].max()}")
print(f"Flete Total (Ingresos): S/ {df['flete_total'].sum():,.2f}")
print(f"Gastos Totales: S/ {df['gastos_total'].sum():,.2f}")
print(f"Margen Bruto Acumulado: S/ {df['margen_calculado'].sum():,.2f}")
print(f"Porcentaje Margen Promedio Glob: {((df['margen_calculado'].sum()/df['flete_total'].sum())*100):.2f}%")

print('\n=== 2. EVOLUCIÓN MENSUAL ===')
df['mes_año'] = df['fecha_salida'].dt.strftime('%Y-%m')
mes_summary = df.groupby('mes_año').agg(
    servicios=('servicio_ida', 'count'),
    flete_ida=('flete_ida', 'sum'),
    flete_retorno=('flete_retorno', 'sum'),
    flete_total=('flete_total', 'sum'),
    gastos_total=('gastos_total', 'sum'),
    margen=('margen_calculado', 'sum')
)
mes_summary['%_margen'] = (mes_summary['margen'] / mes_summary['flete_total']) * 100
print(mes_summary.round(2))

print('\n=== 3. DETALLE SEMANAL ===')
df['semana'] = df['fecha_salida'].dt.isocalendar().week
df['año'] = df['fecha_salida'].dt.year
semana_summary = df.groupby(['año', 'semana']).agg(
    servicios=('servicio_ida', 'count'),
    flete_total=('flete_total', 'sum'),
    gastos_total=('gastos_total', 'sum'),
    margen=('margen_calculado', 'sum')
)
print(semana_summary.tail(8).round(2))

print('\n=== 4. TIPO DE VIAJE (SOLO IDA vs IDA Y VUELTA) ===')
df['tipo_viaje'] = np.where(df['servicio_retorno'] > 0, 'Ida y Vuelta', 'Solo Ida')
viaje_summary = df.groupby('tipo_viaje').agg(
    servicios=('servicio_ida', 'count'),
    flete_total=('flete_total', 'sum'),
    gastos_total=('gastos_total', 'sum'),
    margen=('margen_calculado', 'sum'),
    flete_promedio=('flete_total', 'mean'),
    margen_promedio=('margen_calculado', 'mean')
)
viaje_summary['%_margen'] = (viaje_summary['margen'] / viaje_summary['flete_total']) * 100
print(viaje_summary.round(2))

print('\n=== 5. TOP 10 PLACAS MÁS RENTABLES ===')
top_placas = df.groupby('placa').agg(
    servicios=('servicio_ida', 'count'),
    flete_total=('flete_total', 'sum'),
    gastos_total=('gastos_total', 'sum'),
    margen=('margen_calculado', 'sum'),
    pct_margen=('pct_margen', 'mean')
).sort_values('margen', ascending=False)
print(top_placas.head(10).round(2))

print('\n=== 6. TOP 10 RUTAS MÁS IMPORTANTES ===')
top_rutas = df.groupby('ruta').agg(
    servicios=('servicio_ida', 'count'),
    flete_total=('flete_total', 'sum'),
    gastos_total=('gastos_total', 'sum'),
    margen=('margen_calculado', 'sum')
).sort_values('servicios', ascending=False)
print(top_rutas.head(10).round(2))

print('\n=== 7. TIPO DE SERVICIO (EXCLUSIVO vs CONSOLIDADO vs REPARTO) ===')
tipo_summary = df.groupby('tipo_servicio').agg(
    servicios=('servicio_ida', 'count'),
    flete_total=('flete_total', 'sum'),
    gastos_total=('gastos_total', 'sum'),
    margen=('margen_calculado', 'sum')
).sort_values('servicios', ascending=False)
tipo_summary['%_margen'] = (tipo_summary['margen'] / tipo_summary['flete_total']) * 100
print(tipo_summary.round(2))

print('\n=== 8. ANÁLISIS DE OPERADORES / TERCERIZADOS vs PROPIOS ===')
oper_summary = df.groupby('operador').agg(
    servicios=('servicio_ida', 'count'),
    flete_total=('flete_total', 'sum'),
    gastos_total=('gastos_total', 'sum'),
    margen=('margen_calculado', 'sum')
)
print(oper_summary.round(2))
