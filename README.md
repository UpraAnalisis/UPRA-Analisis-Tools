# UPRA Análisis Tools

<p align="center">
  <img src="/Images/Logo_Upra_Analisis_Tools.png">
</p>


Es una funcionalidad del tipo Add-in (o extensión) que se incorpora a ArcGIS y permite realizar diversas operaciones orientadas a agilizar el proceso de generación de resultados de los geoprocesos adelantados por los profesionales del grupo de análisis, para dar cumplimiento al objeto misional de la [UPRA].

## Requisitos de instalación

Antes de instalar el Add-in Upra_Analisis_Tools verifique lo siguiente:

+ Que tenga permisos de escritura sobre la carpeta de Python en su pc. Para habilitar los permisos de escritura diríjase a la unidad C: de su pc,  haga clic derecho sobre la carpeta Python27 y seleccione propiedades. Una ventana como la mostrada a continuación aparecerá. Cerciórese de que el botón de solo lectura este no esté seleccionado.

<p align="center">
  <img src="Images\Carpeta_python.png">
</p>


## Instalación

El proceso de descarga e instalación es el siguiente:

### Descarga del Add-In:

Para descargar el Add-in haga clic en el link de **Download ZIP** tal como lo muestra la siguiente imagen.

<p align="center">
  <img src="..\Images\Descarga.PNG">
</p>

La instalación puede realizarse de dos formas:
Manual y automatizada.

### Instalación manual

La forma manual consiste en copiar los archivos: **fillable_box.py** y **UPRA-Analisis-Tools.esriaddin** en el directorio de Add-ins de ArcGis tal y como se muestra en la siguiente imagen.

<p align="center">
  <img src="Images\Copia_manual.png">
</p>

En el paso anterior se copiaron los archivos necesarios que componen el Add-in, sin embargo, es necesario instalar librerías de python auxiliares para su correcto funcionamiento. Para llevar a cabo esta tarea se deben ejecutar los archivos **get-pip.py**e **instalador_auxiliar.py**, en el respectivo orden empleando la versión de Python de 32 bits.

Para realizar la ejecución a 32 bits copie la ruta del ejecutable de 32 bits de Python de ArcGis en un bloc de notas seguido de un espacio, luego escriba la ruta del archivo **get-pip.py**, tal como lo indica la imagen.

<p align="center">
  <img src="Images\ejecucion_getpip.png">
</p>

Posteriormente, copie el texto del bloc de notas, presione las teclas **Windows** **+** **E** y pegue el texto en la ventana de ejecución de la siguiente forma.


<p align="center">
  <img src="Images\ejecutar_get_pip.png">
</p>


Para ejecutar el archivo de  **instalador_auxiliar.py** repita el procedimiento anterior.
AL finalizar la ejecución del archivo de instalador_auxiliar deberá aparecer una ventana como la mostrada a continuación, indicando que el proceso ha finalizado con éxito.

<p align="center">
  <img src="Images\auxiliar.png">
</p>

### Instalación automatizada

La forma automatizada consiste en ejecutar el instalador que viene por defecto en el Add-in. Para ello se debe ejecutar el archivo **Instalador_UPRA-Analisis-Tools.py**.

<p align="center">
  <img src="Images\ejecutable.png">
</p>

 Al hacerlo una serie de ventanas aparecerán instalando las librerías de Python necesarias y copiando los archivos del Add-in en su directorio local de ArcGis, al finalizar la instalación, se mostrará la siguiente imagen.

  <p align="center">
   <img src="Images\Ventana_Ejecucion.png">
 </p>

### Configuración

 Si al iniciar ArcMap no encuentra el Add-In, debe configurar su visualización dentro de la barra de herramientas. Para ello, haga clic derecho sobre la barra de herramientas y seleccione la opción **customize**. Al hacerlo se activará el menú mostrado en la siguiente imagen en donde se debe seleccionar la barra de herramientas Upra_Análisis_Tools.

<p align="center">
 <img src="/img/activar_menu.png">
</p>

## Uso

La barra de herramientas del Add-in Upra Análisis Tools se compone de dos grupos de funcionalidades: **Herramientas para presentaciones** y **Herramientas para manipular layers**.

<p align="center">
 <img src="Images\barra_de_herramientas.png">
</p>

### Herramientas para presentaciones
Herramientas que preparan de forma masiva los insumos para construir presentaciones con la información resultante de los geoprocesos.


+ #### Multicortes [(Ver ayuda)](https://github.com/UpraAnalisis/UPRA-Analisis-Tools/blob/master/help/Multicortes.md#multicortes-)
+ #### ToPowerpoint [(Ver ayuda)](https://github.com/UpraAnalisis/UPRA-Analisis-Tools/blob/master/help/to_powerpoint.md#to_powerpoint-)


### Herramientas para manipular layers
Herramientas que permiten consultar, filtrar y exportar de forma masiva las capas y tablas cargadas en un mxd.

+ #### Área [(Ver ayuda)](https://github.com/UpraAnalisis/UPRA-Analisis-Tools/blob/master/help/Herramientas_layers.md#%C3%81rea)
+ #### Ruta [(Ver ayuda)](https://github.com/UpraAnalisis/UPRA-Analisis-Tools/blob/master/help/Herramientas_layers.md#ruta)
+ #### Coordenadas [(Ver ayuda)](https://github.com/UpraAnalisis/UPRA-Analisis-Tools/blob/master/help/Herramientas_layers.md#coordenadas)
+ #### Estadísticas [(Ver ayuda)](https://github.com/UpraAnalisis/UPRA-Analisis-Tools/blob/master/help/Herramientas_layers.md#estad%C3%ADsticas)
+ #### Tabla a Excel [(Ver ayuda)](https://github.com/UpraAnalisis/UPRA-Analisis-Tools/blob/master/help/Herramientas_layers.md#tabla-a-excel)
+ #### Multi Query [(Ver ayuda)](https://github.com/UpraAnalisis/UPRA-Analisis-Tools/blob/master/help/Herramientas_layers.md#multi-query)
+ #### Multi Export [(Ver ayuda)](https://github.com/UpraAnalisis/UPRA-Analisis-Tools/blob/master/help/Herramientas_layers.md#multi-export)
