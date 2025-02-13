## Script que compara facturas cargadas contra Recibidos de AFIP
import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe

compras_mercovan = pd.read_excel('compras mercovan.xlsx', header=1)
compras_afip = pd.read_excel('compras afip.xlsx', header=1)

compras_afip = compras_afip.rename(columns={'Número Desde': 'Comprobante', 
                            'Nro. Doc. Emisor': 'CUIT'})

compras_mercovan = compras_mercovan.rename(columns={'Nrofac': 'Comprobante',
                                                    'Cuit': 'CUIT'})

compras_mercovan = compras_mercovan.iloc[1:]

compras_mercovan['Comprobante'] = compras_mercovan.loc[:, 'Comprobante'].astype(int).astype(str)
compras_mercovan.loc[:, 'CUIT'] = compras_mercovan['CUIT'].apply(lambda x: '{:.0f}'.format(x))

compras_afip.loc[:, 'Comprobante'] = compras_afip['Comprobante'].astype(str)
compras_afip.loc[:, 'CUIT'] = compras_afip['CUIT'].astype(str)

mismatches_mercovan = compras_mercovan[~compras_mercovan.set_index(['Comprobante', 'CUIT']).index.isin(compras_afip.set_index(['Comprobante', 'CUIT']).index)]
mismatches_afip = compras_afip[~compras_afip.set_index(['Comprobante', 'CUIT']).index.isin(compras_mercovan.set_index(['Comprobante', 'CUIT']).index)]

mismatches_mercovan = mismatches_mercovan[['Razon Social', 'N.Prov', 'Serie', 'P.Vtas', 'Comprobante', 'Fec.Fac', 'CUIT',  'Total Fac.', ]]
mismatches_afip = mismatches_afip[['Denominación Emisor', 'Fecha', 'Tipo', 'Comprobante', 'CUIT', 'Imp. Neto Gravado', 'IVA', 'Imp. Total']]
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

print("Facturas que faltan en AFIP:")
print(mismatches_mercovan)
print("\n\nFacturas que faltan en Mercovan:")
print(mismatches_afip)

gc = gspread.service_account(filename='credenciales_gsheets.json')

google_sheet = gc.create(f'Comparacion compras Mercovan-AFIP {pd.Timestamp.now().date()}')

google_sheet.share('marajadesantelmo@gmail.com', perm_type='user', role='writer')
#google_sheet.share('manuel@dassa.com.ar', perm_type='user', role='writer')
#822163google_sheet.share('jose@mercovan.com.ar', perm_type='user', role='writer')

faltan_afip = google_sheet.add_worksheet(title='Faltan en AFIP', rows="100", cols="20")
set_with_dataframe(faltan_afip, mismatches_mercovan)

faltan_mercovan = google_sheet.add_worksheet(title='Faltan en Mercovan', rows="100", cols="20")
set_with_dataframe(faltan_mercovan, mismatches_afip)

### Chequeo de montos diferentes

merged = compras_mercovan.merge(compras_afip, on=['Comprobante', 'CUIT'], suffixes=('_mercovan', '_afip'))
mismatched_totals = merged[merged['Total Fac.'] != merged['Imp. Total']]

# Crear nuevas hojas en Google Sheets para reportar discrepancias en los totales
mismatched_totals_mercovan = mismatched_totals[['Comprobante', 'CUIT', 'Total Fac.']]
mismatched_totals_afip = mismatched_totals[['Comprobante', 'CUIT', 'Imp. Total']]

diferencias_mercovan = google_sheet.add_worksheet(title='Diferencias en montos Mercovan', rows="100", cols="20")
set_with_dataframe(diferencias_mercovan, mismatched_totals_mercovan)

diferencias_afip = google_sheet.add_worksheet(title='Diferencias en montos AFIP', rows="100", cols="20")
set_with_dataframe(diferencias_afip, mismatched_totals_afip)