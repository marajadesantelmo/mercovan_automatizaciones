## Script que compara facturas cargadas contra Recibidos de AFIP
import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe
import os

if os.path.exists('C:/Users/Karelys/Desktop/mercovan_automatizaciones'):
    os.chdir('C:/Users/Karelys/Desktop/mercovan_automatizaciones')

print('Abriendo datos de mercovan...')
compras_mercovan = pd.read_excel('compras mercovan.xlsx', header=1)

pd.set_option('display.max_rows', 10)
pd.set_option('display.max_columns', 10)

### Procesamiento recibidos AFIP
def process_recibidos_xlsx(file_path):
    df_recibidos = pd.read_excel(file_path, header=1)
    df_recibidos = df_recibidos.fillna(0)
    #df_recibidos.loc[df_recibidos['Tipo'].str.contains('3'), ['Imp. Neto Gravado', 'Imp. Neto No Gravado', 'Imp. Op. Exentas', 'IVA', 'Imp. Total']] *= -1
    #df_recibidos.loc[df_recibidos['Tipo'].str.contains('8'), ['Imp. Neto Gravado', 'Imp. Neto No Gravado', 'Imp. Op. Exentas', 'IVA', 'Imp. Total']] *= -1
    numeric_columns = ['Tipo Cambio', 'Imp. Neto Gravado', 'Imp. Neto No Gravado', 'Imp. Op. Exentas', 'Otros Tributos', 'IVA', 'Imp. Total']
    for column in numeric_columns:
        if column != 'Tipo Cambio':  # Skip Tipo Cambio as it is the exchange rate
            df_recibidos.loc[df_recibidos['Moneda'] == 'USD', column] *= df_recibidos.loc[df_recibidos['Moneda'] == 'USD', 'Tipo Cambio']
    df_recibidos[numeric_columns] = df_recibidos[numeric_columns].apply(pd.to_numeric, errors='coerce').fillna(0) # .astype(int)
    return df_recibidos

print('Abriendo datos de AFIP...')
compras_afip = process_recibidos_xlsx('compras afip.xlsx')

compras_afip = compras_afip.rename(columns={'Número Desde': 'Comprobante', 
                            'Nro. Doc. Emisor': 'CUIT'})
compras_mercovan = compras_mercovan.rename(columns={'Nrofac': 'Comprobante',
                                                    'Cuit': 'CUIT'})
compras_mercovan = compras_mercovan.iloc[1:]
compras_mercovan['Comprobante'] = compras_mercovan.loc[:, 'Comprobante'].astype(int).astype(str)
compras_mercovan.loc[:, 'CUIT'] = compras_mercovan['CUIT'].apply(lambda x: '{:.0f}'.format(x))
compras_afip.loc[:, 'Comprobante'] = compras_afip['Comprobante'].astype(str)
compras_afip.loc[:, 'CUIT'] = compras_afip['CUIT'].astype(str)
print('Chequeando facturas faltantes...')
mismatches_mercovan = compras_mercovan[~compras_mercovan.set_index(['Comprobante', 'CUIT']).index.isin(compras_afip.set_index(['Comprobante', 'CUIT']).index)]
mismatches_afip = compras_afip[~compras_afip.set_index(['Comprobante', 'CUIT']).index.isin(compras_mercovan.set_index(['Comprobante', 'CUIT']).index)]

mismatches_mercovan = mismatches_mercovan.iloc[:, 1:]
mismatches_afip = mismatches_afip[['Denominación Emisor', 'Fecha', 'Tipo', 'Comprobante', 'CUIT', 'Imp. Neto Gravado', 'Imp. Neto No Gravado', 'Imp. Op. Exentas', 'Otros Tributos', 'IVA', 'Imp. Total']]

### Chequeo de montos diferentes
merged = compras_mercovan.merge(compras_afip, on=['Comprobante', 'CUIT'], suffixes=('_mercovan', '_afip'))
mismatched_totals = merged[(merged['Total Fac.'] != merged['Imp. Total']) | (merged['T.Iva'] != merged['IVA'])]
mismatched_totals['Dif. Importe'] = mismatched_totals['Total Fac.'] - mismatched_totals['Imp. Total']
mismatched_totals['Dif. IVA'] = mismatched_totals['T.Iva'] - mismatched_totals['IVA']
mismatched_totals_mercovan = mismatched_totals[['Razon Social', 'Fec.Fac', 'N.Prov', 'Serie', 'P.Vtas', 'Comprobante', 'CC', 'CUIT', 'T.Iva', 'Total Fac.', 'Dif. IVA', 'Dif. Importe']]
mismatched_totals_afip = mismatched_totals[['Denominación Emisor', 'Fecha', 'Tipo', 'Comprobante', 'CUIT', 'Imp. Neto Gravado', 'Imp. Neto No Gravado', 'Imp. Op. Exentas', 'Otros Tributos', 'IVA', 'Imp. Total', 'Dif. IVA', 'Dif. Importe']]


pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

print("Facturas que faltan en AFIP:")
print(mismatches_mercovan)
print("\n\nFacturas que faltan en Mercovan:")
print(mismatches_afip)
print("Subiendo a sheets")
gc = gspread.service_account(filename='credenciales_gsheets.json')
google_sheet = gc.create(f'Comparacion compras Mercovan-AFIP {pd.Timestamp.now().date()}')

faltan_afip = google_sheet.add_worksheet(title='Faltan en AFIP', rows="100", cols="20")
set_with_dataframe(faltan_afip, mismatches_mercovan)

faltan_mercovan = google_sheet.add_worksheet(title='Faltan en Mercovan', rows="100", cols="20")
set_with_dataframe(faltan_mercovan, mismatches_afip)

diferencias_mercovan = google_sheet.add_worksheet(title='Diferencias en montos Mercovan', rows="100", cols="20")
set_with_dataframe(diferencias_mercovan, mismatched_totals_mercovan)

diferencias_afip = google_sheet.add_worksheet(title='Diferencias en montos AFIP', rows="100", cols="20")
set_with_dataframe(diferencias_afip, mismatched_totals_afip)

# Delete "Sheet1" if exists
try:
    sheet1 = google_sheet.worksheet("Sheet1")
    google_sheet.del_worksheet(sheet1)
except gspread.exceptions.WorksheetNotFound:
    print("Sheet1 does not exist, no need to delete.")


google_sheet.share('marajadesantelmo@gmail.com', perm_type='user', role='writer')
google_sheet.share('manuel@dassa.com.ar', perm_type='user', role='writer')
google_sheet.share('jose@mercovan.com.ar', perm_type='user', role='writer')
google_sheet.share('karelys@mercovan.com.ar', perm_type='user', role='writer')