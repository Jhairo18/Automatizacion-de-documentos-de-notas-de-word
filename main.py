import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate


# El presente programa consiste que a partir de un documento plantilla (plantilla.docx) tenemos una 
# carta que dara las notas de los alumnos, pero tenemos diferentes alumnos y diferentes notas
# Tenemos estos datos en un archivo excel. El programa consiste en llenar los datos cambiantes de
# la plantilla las cuales estaran referenciadas como {{}} con los datos excel.


# La libreria docxtpl sirve para poder trabajar los archivos de word, además que con el render podemos
# reemplazar algunas palabras escritas en el siguiente formato {{palabra}}, y con el save pues guardamos 
# los archivos.
# En pandas usamos read_excel para leer los archivos excel de python, además  de usar el metodo iterrows 
# que me permite como el mismo nombre dice iterar las filas, y asi obtener por una parte los indices
# y también las filas las cuales estaran representadas por una de pdSeries.
# Con datetime obtenemos la fecha actualizada, poniendo nuestro formato de tiempo usando el metodo
# strftime 


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
