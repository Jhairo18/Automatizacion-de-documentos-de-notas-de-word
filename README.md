# Automatizacion de documento de notas usando word y excel
El presente programa consiste que a partir de un documento plantilla (plantilla.docx) tenemos una  carta que dara las notas de los alumnos, pero tenemos diferentes alumnos y diferentes notas.
Tenemos los nombres y notas en un archivo excel. El programa consiste que a partir de la hoja en excel crear documentos de word para cada alumno de manera personalizada.

La libreria docxtpl sirve para poder trabajar los archivos de word, además que con el render podemos reemplazar algunas palabras escritas en el siguiente formato {{palabra}}, y con el save pues guardamos  los archivos.
En pandas usamos read_excel para leer los archivos excel de python, además  de usar el metodo iterrows  que me permite como el mismo nombre dice iterar las filas, y asi obtener por una parte los indices y también las filas las cuales estaran representadas por una de pdSeries.
Con datetime obtenemos la fecha actualizada, poniendo nuestro formato de tiempo usando el metodo strftime 
