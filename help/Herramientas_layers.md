# Área

Esta herramienta muestra en consola el área de las capas seleccionadas en metros cuadrados y en hectáreas.

<p align="center">
  <img src="..\Images\Consola_area.PNG">
</p>

En caso de no tener ninguna capa seleccionada se mostrará el siguiente mensaje.

<p align="center">
  <img src="..\Images\Consola_area_error.PNG">
</p>

# Ruta

Esta herramienta muestra en consola la ruta de las capas seleccionadas.
<p align="center">
  <img src="..\Images\Consola_rutas.PNG">
</p>

En caso de no tener ninguna capa seleccionada se mostrará el siguiente mensaje.

<p align="center">
  <img src="..\Images\Consola_area_error.PNG">
</p>

# Coordenadas

Esta herramienta muestra el nombre del sistema de coordenadas de las capas seleccionadas.
<p align="center">
  <img src="..\Images\Consola_cordenadas.PNG">
</p>

En caso de no tener ninguna capa seleccionada se mostrará el siguiente mensaje.

<p align="center">
  <img src="..\Images\Consola_area_error.PNG">
</p>

# Estadísticas

Esta herramienta crea una tabla en memoria con las estadísticas de una capa, teniendo en cuenta un campo definido. Para definir el campo de estadísticas basta con: **editar el nombre de la capa en el mxd colocando dos guiones seguidos del nombre del campo de la capa sobre el cual se van a realizar las estadísticas**

A continuación una imagen que ilustra la configuración y los resultados.


<p align="center">
  <img src="..\Images\Resultados_estadisticas.png">
</p>


En caso de no tener ninguna capa seleccionada se mostrará el siguiente mensaje.

<p align="center">
  <img src="..\Images\Consola_area_error.PNG">
</p>

# Tabla a Excel

Esta herramienta convierte las tablas seleccionadas en archivos de Excel.
Al hacer clic en la herramienta se despliega una ventana solicitando al usuario la ruta de almacenamiento de los archivos.

<p align="center">
  <img src="..\Images\ruta_de_salida.png">
</p>

Una vez seleccionada la ruta, se desplegará una nueva ventana listando todas las tablas disponibles para la conversión. El usuario deberá seleccionarlas y hacer clic en el botón **convertir**.

<p align="center">
  <img src="..\Images\tablas_a_excel.PNG">
</p>

En caso de no tener ninguna capa seleccionada se mostrará el siguiente mensaje.

<p align="center">
  <img src="..\Images\consola_estadisticas_error.PNG">
</p>

# Multi Query

Esta herramienta permite definir una expresión SQL en la propiedad definition query para todas las capas seleccionadas.

Al hacer clic en la herramienta se despliega una ventana en la cual se debe escribir la expresión SQL. Una vez escrita se debe hacer clic en el botón **OK**.

<p align="center">
  <img src="..\Images\ventana_query.png">
</p>

# Multi Export

Esta herramienta permite exportar en formato tabla o Feature Class las capas seleccionadas. Al activar la herramienta aparecerá una ventana solicitando el formato de salida de los archivos.

<p align="center">
  <img src="..\Images\menu_multiexport.png">
</p>

Posteriormente la herramienta aparecerá una ventana solicitando la ruta de salida.

<p align="center">
  <img src="..\Images\ruta_salida_multiexport.png">
</p>

El usuario deberá dar la ruta de una File Geodatabase en caso de quiera exportar tablas y Feature class a una geodatabase, o la ruta de una carpeta en caso de que quiera convertir las capas a formato Esri Shapefile.

En caso de no tener ninguna capa seleccionada se mostrará el siguiente mensaje.

<p align="center">
  <img src="..\Images\Consola_area_error.PNG">
</p>
