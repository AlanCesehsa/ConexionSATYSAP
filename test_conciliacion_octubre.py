import openpyxl
import pandas as pd
from conexion_sap_db import obtener_facturas_sap_db
from motor_conciliacion import normalizar_sat_df, normalizar_sap_df

# 1. Obtener SAP de BD
df_sap = obtener_facturas_sap_db('YUKME', 2026, 10, 10)
print(f'Filas obtenidas de SAP: {len(df_sap)}')

# 2. Cargar SAT del archivo del contador para simular la descarga oficial completa
f_contador = r'c:\Users\Stock CESEHSA\Downloads\CONCILIACION DE FACTURAS YUKME OCTUBRE 2026.xlsx'
df_sat = pd.read_excel(f_contador, sheet_name='SAT')
print(f'Filas obtenidas de SAT: {len(df_sat)}')

# 3. Analizar comprobantes SAT
sat_norm = normalizar_sat_df(df_sat)
print(f'Filas normalizadas SAT: {len(sat_norm)}')
print(sat_norm['sat_tipo'].value_counts())
print(sat_norm['sat_estatus'].value_counts())

# 4. Cruzar
# Sumas SAP
sap_pesos = df_sap[df_sap['Moneda'].isin(['MXP', 'MXN'])]['TotalPesos'].sum()
sap_usd = df_sap[df_sap['Moneda'] == 'USD']['TotalUSD'].sum()

print(f'\nSAP Bruto Pesos: {sap_pesos:,.2f}')
print(f'SAP Bruto USD:   {sap_usd:,.2f}')

# Separar NC de SAP
nc_sap = df_sap[df_sap['Documento'].str.contains('Credito|Nota', case=False, na=False)]
print(f'\nTotal Notas de Crédito en SAP: {len(nc_sap)}')

sat_map = {row['uuid_clean']: row for _, row in sat_norm.iterrows()}

nc_fiscales = []
nc_no_fiscales = []

for _, r in nc_sap.iterrows():
    u = str(r.get('UUID', '')).strip().upper()
    if u in sat_map and sat_map[u]['sat_tipo'] == 'E' and sat_map[u]['sat_estatus'] == 'Vigente':
        nc_fiscales.append(r)
    else:
        nc_no_fiscales.append(r)

df_nc_no_fisc = pd.DataFrame(nc_no_fiscales)
print(f'NC Fiscales timbradas en SAT: {len(nc_fiscales)}')
print(f'NC No Fiscales (internas):    {len(df_nc_no_fisc)}')

for _, r in df_nc_no_fisc.iterrows():
    print(f"  Folio: {r['NoDocumento']}, Cliente: {r['Cliente']}, Moneda: {r['Moneda']}, TotalPesos: {r['TotalPesos']}, TotalUSD: {r['TotalUSD']}")

nc_no_fisc_pesos = df_nc_no_fisc[df_nc_no_fisc['Moneda'].isin(['MXP', 'MXN'])]['TotalPesos'].sum()
nc_no_fisc_usd = df_nc_no_fisc[df_nc_no_fisc['Moneda'] == 'USD']['TotalUSD'].sum()

print(f'\nTotal NC No Fiscales Pesos: {nc_no_fisc_pesos:,.2f}')
print(f'Total NC No Fiscales USD:   {nc_no_fisc_usd:,.2f}')

# Totales SAT
# Cruzar moneda con SAP para asignar Pesos o USD a los comprobantes SAT
sat_i_pesos = 0.0
sat_i_usd = 0.0
sat_e_pesos = 0.0
sat_e_usd = 0.0

sap_mon_map = {str(r.get('UUID', '')).strip().upper(): r['Moneda'] for _, r in df_sap.iterrows() if str(r.get('UUID', '')).strip()}

for _, r in df_sat.iterrows():
    u = str(r['Uuid']).strip().upper()
    est = str(r.get('Estatus', '')).strip()
    # Estatus 1 o Vigente
    if est not in ['1', 1, 'Vigente', 'VIGENTE']:
        continue
    monto = float(r['Monto'])
    tipo = str(r.get('EfectoComprobante', 'I')).strip().upper()
    mon = sap_mon_map.get(u, 'MXN')
    
    if tipo == 'I':
        if mon == 'USD':
            sat_i_usd += monto
        else:
            sat_i_pesos += monto
    elif tipo == 'E':
        if mon == 'USD':
            sat_e_usd += monto
        else:
            sat_e_pesos += monto

sat_neto_pesos = sat_i_pesos - sat_e_pesos
sat_neto_usd = sat_i_usd - sat_e_usd

print(f'\nSAT Ingreso Pesos: {sat_i_pesos:,.2f} | Egreso: {sat_e_pesos:,.2f} | Neto: {sat_neto_pesos:,.2f}')
print(f'SAT Ingreso USD:   {sat_i_usd:,.2f} | Egreso: {sat_e_usd:,.2f} | Neto: {sat_neto_usd:,.2f}')

# Conciliación ajustada
sap_ajust_pesos = sap_pesos - nc_no_fisc_pesos
sap_ajust_usd = sap_usd - nc_no_fisc_usd

diff_pesos = sap_ajust_pesos - sat_neto_pesos
diff_usd = sap_ajust_usd - sat_neto_usd

print(f'\n=== RESULTADO FINAL DE CONCILIACIÓN ===')
print(f'PESOS:   SAP Ajustado: {sap_ajust_pesos:,.2f} vs SAT: {sat_neto_pesos:,.2f} -> Diferencia: {diff_pesos:,.4f}')
print(f'DOLARES: SAP Ajustado: {sap_ajust_usd:,.2f} vs SAT: {sat_neto_usd:,.2f} -> Diferencia: {diff_usd:,.4f}')

