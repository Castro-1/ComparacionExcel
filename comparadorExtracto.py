import PySimpleGUI as sg
import pandas as pd
import xlwings as xw

sg.theme('DarkTeal9')

layout = [
    [sg.Text("Por favor llene TODOS los campos:")],
    [sg.Text("Para las celdas utilice valores como 'N4', N = Columna, 4 = Fila.")],
    [sg.Text('Archivo'),sg.Input(key="Archivo"), sg.FileBrowse()],
    [sg.Text('Hoja 1', size=(15,1)),sg.InputText(key="Hoja1")],
    [sg.Text('Celda Valores', size=(15,1)),sg.InputText(key="CeldaV")],
    [sg.Text('Hoja 2', size=(15,1)),sg.InputText(key="Hoja2")],
    [sg.Text('Celda Debitos', size=(15,1)),sg.InputText(key="CeldaD")],
    [sg.Text('Celda Creditos', size=(15,1)),sg.InputText(key="CeldaC")],
    [sg.Submit(), sg.Exit()],
]

window = sg.Window("Comparador Extracto", layout)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    if event == "Submit":
        print(values["Archivo"])
        archivo = values["Archivo"]
        hoja1 = values["Hoja1"]
        hoja2 = values["Hoja2"]
        celdaValores = values["CeldaV"]
        celdaDebitos = values["CeldaD"]
        celdaCreditos = values["CeldaC"]
        # Abre los archivos de Excel utilizando xlwings
        workbook1 = xw.Book(archivo)

        # Selecciona las hojas en las que deseas buscar
        sheet1 = workbook1.sheets[hoja1]
        sheet2 = workbook1.sheets[hoja2]

        # Crea un DataFrame de pandas a partir de los datos en la columna
        df1 = pd.DataFrame(sheet1.range(celdaValores).expand('down').value, columns=['Valor'])
        df2 = pd.DataFrame(sheet2.range(celdaDebitos).expand('down').value, columns=['Debitos'])
        df3 = pd.DataFrame(sheet2.range(celdaCreditos).expand('down').value, columns=['Creditos'])

        # Crea una nueva columna en cada DataFrame para indicar si la celda ya ha sido coloreada
        df1['Coloreado'] = False
        df2['Coloreado'] = False
        df3['Coloreado'] = False

        # # Recorre cada fila en el DataFrame de la primera hoja y busca el valor en el DataFrame de la segunda hoja
        for index, row in df1.iterrows():
            value = row['Valor']
            if value > 0:
                # print(value)
                currval = df2[(df2['Debitos'] == value) & (df2['Coloreado'] == False)]['Coloreado']
            else:
                currval = df3[(df3['Creditos'] == value*-1) & (df3['Coloreado'] == False)]['Coloreado']

            if currval.empty:
                continue
            # print(currval)
            # Si se encuentra el valor y la celda correspondiente no está coloreada, entonces colorear las celdas correspondientes
            sheet1.range(celdaValores).offset(index, 0).color = (255, 255, 0)  # Color de celda roja
            df1.loc[index, 'Coloreado'] = True
            if value > 0:
                sheet2.range(celdaDebitos).offset(currval.index[0], 0).color = (255, 255, 0)  # Color de celda roja
                df2.loc[currval.index[0], 'Coloreado'] = True
                continue
            sheet2.range(celdaCreditos).offset(currval.index[0], 0).color = (255, 255, 0)  # Color de celda roja
            df3.loc[currval.index[0], 'Coloreado'] = True
        break



window.close()