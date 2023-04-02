import pandas as pd
import xlwings as xw

# Abre los archivos de Excel utilizando xlwings
workbook1 = xw.Book('PRUEBA.xlsx')

# Selecciona las hojas en las que deseas buscar
sheet1 = workbook1.sheets['Hoja2']
sheet2 = workbook1.sheets['Hoja1']

# Crea un DataFrame de pandas a partir de los datos en la columna
df1 = pd.DataFrame(sheet1.range('E4').expand('down').value, columns=['Valor'])
df2 = pd.DataFrame(sheet2.range('N4').expand('down').value, columns=['Debitos'])
df3 = pd.DataFrame(sheet2.range('O4').expand('down').value, columns=['Creditos'])

print(df1.head())
print(df2.head())

# Crea una nueva columna en cada DataFrame para indicar si la celda ya ha sido coloreada
df1['Coloreado'] = False
df2['Coloreado'] = False
df3['Coloreado'] = False
# print(df1.head())
# print(df2.head())

# # Recorre cada fila en el DataFrame de la primera hoja y busca el valor en el DataFrame de la segunda hoja
for index, row in df1.iterrows():
    value = row['Valor']
    # print(value)
    currval = df2[(df2['Debitos'] == value) & (df2['Coloreado'] == False)]['Coloreado']
    if currval.empty:
        continue
    # print(currval)
    if currval.iloc[0] == False:
        # Si se encuentra el valor y la celda correspondiente no está coloreada, entonces colorear las celdas correspondientes
        sheet1.range('E4').offset(index, 0).color = (255, 255, 0)  # Color de celda roja
        sheet2.range('N4').offset(df2[(df2['Debitos'] == value) & (df2['Coloreado'] == False)].index[0], 0).color = (255, 255, 0)  # Color de celda roja
        df1.loc[index, 'Coloreado'] = True
        df2.loc[df2[(df2['Debitos'] == value) & (df2['Coloreado'] == False)].index[0], 'Coloreado'] = True
