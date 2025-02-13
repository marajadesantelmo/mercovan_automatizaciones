# Script que compara movimientos bancarios de Credicoop con los de la empresa

# Importamos las librerias necesarias
import pandas as pd

# Cargamos los archivos de Excel
mayor = pd.read_excel('libro mayor credicoop.xlsx')
movimientos = pd.read_excel('movimientos credicoop.xls')

compras = pd.read_excel('compras mercovan.xlsx')