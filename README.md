# MultiCortes AddIn

Es una herramienta del tipo AddIn que se incorpora a ArcGIS y permite realizar el corte de múltiples capas con respecto a una o varias zonas de interés o moldes y obtener las estadísticas de área de acuerdo con un campo definido. Esta herramienta se desarrolló con el fin de agilizar el proceso de generación de resultados de los geoprocesos adelantados por los profesionales del grupo de análisis, para dar cumplimiento al objeto misisonal de la  [UPRA](http://www.upra.gov.co).


## Instalación

El proceso de instalación es el siguiente:

1. Descargar el `AddIn`:

    Existen dos opciones de descarga, se puede descargar todo el repositorio o solo el ultimo Release del AddIn [Aquí](https://github.com/UpraAnalisis/MultiCortes/releases/latest)

2. Luego hay dos opciones:

3. Si descargó el código, descomprimir y ejecutar el script `makeaddin.py`

4. Si descargó el release, descomprimir. 

5. finalmente común a ambos copiar el archivo con extensión `.esriaddin` en el directorio:

    ```directorio Arcgis
    %USERPROFILE%\Documents\ArcGIS\AddIns\
    ``` 

Ingresar a la versión de arcgis en la que se quiera instalar el AddIn y pegar el archivo.

## Configuración

Si al iniciar ArcMap no encuentra el addin debe configurar su visualización dentro de la barra de herramientas. Para ello haga clic derecho sobre la barra de herramientas y seleccione la opción customize. AL hacerlo se mostrará el menú mostrado en la siguiente imágen en donde debe seleccionar la barra de herramientas Upra_Análisis_Tools.
(/img/activar_menu.png)

## Uso
### Configurar el mxd
1. Construir tres dataframes con la estructura mostrada en la siguiente imágen:
 (/img/dataframes.PNG)
Recuerde mantener pausada la visualización para mejorar el rendimiento y la estabiidad del addin.

### Cargar Capas Necesarias

1. Las capas necesarias para la operación de este addin son de dos (2) tipos y deben estar separadas en diferentes dataframes de la siguiente forma: 
moldes: Este dataframe contiene las capas con las cuales se realizarán los cortes.
capas: Este dataframe contiene las capas que serán cortadas y sobre las cuales se realizarán las estadísticas.

2. Si se va a realizar un corte con estadísticas o una extracción por campo con estadísticas, deben configurarse para cada capa los campos sobre los cuales se van a sacar las estadísticas. Para ello se escribirán dos guiones al final del nombre de cada capa, seguidos del nombre del campo con el cual se van a hacer las estadísticas. A continuación una imagen que representa lo anteriormente mencionado.
(/img/confi_nombres.png


### Selección criterio

1. Contruir tres dataframes con la estrutura mostrada en la siguiente imágen:

  

2. .

    ![Select Herramienta](/img/selherr.PNG)

3. Dar Click en el punto de interés que se desea consultar 

    ![Dataframes](/img/dataframes.PNG)

### Selección Variables

Una vez realizado el proceso anterior el AddIn consulta en la gdb con las variables aquellas que corresponden a el criterio de interés, luego de eso carga los nombre en el menú desplegable y estan listas para ser adicionadas a la visualización del mapa. 

1. Seleccione variable a cargar en el en el menú desplegable

    ![Select Variable](/img/seleccvar.png)

2. Mensaje advierte que la variable se está cargando, pulse OK

    ![Cargar Variable](/img/carvar.PNG)

### Generación Reporte

1. Se carga la variable y mensaje Advierte si quiere generar reporte. (Reporte es Opcional)

    ![Generar Report](/img/genrep.PNG)

2. Si decide generar reporte seleccione donde lo desea guardar.

    ![Salvar Report](/img/savrep.PNG)

3. De no generar reporte los datos quedan en un layer temporal llamado data y ahí los puede consultar. 

    ![Capa Cargada](/img/capcar.PNG)
    
4. El reporte aparece en excel de la siguiente forma: 

    ![report](/img/rep.PNG)
