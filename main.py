import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate

doc = DocxTemplate("plantilla.docx")

nombre = "Jhairo Yurivilca"
telefono = "994 332 142"
correo = "jhairoyuri@gmail.com"
fecha = datetime.today().strftime("%d/%m/%Y")

constantes = {'nombre':nombre, 'telefono':telefono, 'correo':correo,'fecha':fecha}

df = pd.read_excel("alumnos.xlsx")
print(df)
# Itera sobre las filas del df, obteniendo el indice una panda series de las filas
for indice, fila in df.iterrows():
    contenido = {"nombre_alumno":fila["Nombre"],"nota_mat":fila["Mat"],"nota_fis":fila["Fis"],"nota_qui":fila["Qui"]}
    # Actualiza el diccionario
    contenido.update(constantes)
    # Cambia las palabras dadas por {{}} de la plantilla, con un diccionario
    doc.render(contenido)
    doc.save(f"prueba de {fila['Nombre']}.docx")    
