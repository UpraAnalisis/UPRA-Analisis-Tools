# Multicortes
<p align="center">
 <img style width="40%" src="..\Images\logo_multicortes.gif" >
</p>

La funcionalidad de Multicortes permite realizar el corte y las estadísticas de un conjunto de capas teniendo en cuenta unos moldes definidos.

## ¿Cómo funciona?

<br></br>
<p align="center">
 <img style src="..\Images\descripcion_multicortes.png" >
<br></br>
<br></br>
<p align="center">
 <img style src="..\Images\molde1.png" >
<br></br>
<br></br>
<p align="center">
 <img style src="..\Images\molde2.png" >
 <br></br></p>


## Configurar el mxd

Para emplear esta funcionalidad es necesario crear tres dataframes, dos deben llevar de forma obligatoria los nombres : **moldes** y **capas** respectivamente y el tercero puede tener cualquier nombre. A continuación una imagen que ilustra lo descrito.

<p align="center">
 <img src="/img/dataframes.PNG">
</p>

Tenga en cuenta:

Mantenga abierta la consola de python para visualizar los mensajes de error que se muestren en la ejecución de la herramienta.

<p align="center">
 <img src="..\Images\icono_consola_python.png">
</p>

Recuerde mantener pausada la visualización para mejorar el rendimiento y la estabilidad del Add-in.

<p align="center">
 <img src="..\Images\pausar_dibujado.png">
</p>

## Cargar Capas Necesarias

Las capas necesarias para la operación de este Add-in son de dos (2) tipos y deben estar separadas en diferentes dataframes de la siguiente forma:

**moldes**: Este dataframe contiene las capas con las cuales se realizarán los cortes.

**capas**: Este dataframe contiene las capas que serán cortadas y sobre las cuales se realizarán las estadísticas.

Si se desea obtener estadísticas a partir de las capas, deben configurarse para cada capa los campos sobre los cuales se van a procesar. Para ello se escribirán dos guiones al final del nombre de cada capa, seguidos del nombre del campo con el cual se van a hacer las estadísticas. **En caso de que muchas capas tengan el mismo nombre de campo, este no debe escribir en el mxd si no, que cuando se vaya a realizar el corte se escribirá como campo opcional**. A continuación una imagen que representa lo anteriormente mencionado.


<p align="center">
 <img src="/img/confi_nombres.png">
</p>


Tal y como se describe la capa **Z_PINA_GENERAL_FINAL_25ha** no tiene asociado ningún campo de estadísticas y esto se debe a que este será suministrado cuando las ventanas emergentes lo soliciten.

<p align="center">
 <img src="..\Images\campo_estadisticas.png">
</p>

**Tenga en cuenta**:
Al finalizar la configuración de los campos de estadísticas y las capas, es necesario que el usuario active el tercer dataframe donde se alojarán los resultados.

El mxd totalmente configurado y dispuesto para su ejecución debe lucir de la siguiente forma:

<p align="center">
 <img src="..\Images\mxd_configurado.png">
</p>


## Realizar los cortes

Una vez se construyan los dataframes y se adicionen las capas necesarias con su respectivo campo de estadísticas se procederá a realizar el corte. Para hacerlo haga clic sobre el ícono de multicortes localizado en la barra de herramientas en el menú de presentaciones.

Al hacer clic aparecerá un ventana preguntando la ruta de salida de los resultados. Tal y como se muestra en la siguiente imagen.

<p align="center">
 <img src="..\Images\ruta_de_salida.png">
</p>

Seleccione la carpeta y haga clic en **add**. Una vez seleccionada la ruta de salida aparecerá el menú principal de la herramienta.

<p align="center">
 <img src="..\Images\tipo_analisis.png">
</p>

Si se selecciona la opción nada, la herramienta finaliza la ejecución y muestra un mensaje de despedida.

<p align="center">
 <img src="..\Images\ventana_salida.png">
</p>


### Corte y estadísticas

Si se selecciona esta opción automáticamente se despliega una nueva ventana preguntando qué tipo de operación de corte se efectuará.

<p align="center">
 <img src="..\Images\tipo_de_corte.png">
</p>

Al seleccionar la operación de corte, aparecerá la ventana solicitando el campo opcional que será usado para hacer las estadísticas sobre las capas que no tengan definido el campo de estadísticas en el mxd.


<p align="center">
 <img src="..\Images\campo_estadisticas.png">
</p>

Por último, aparecerá una ventana preguntando si el usuario desea continuar con la ejecución. En caso de ser afirmativo la ejecución continua, en caso contrario, aparece el mensaje de despedida antes mostrado.

Al continuar la ejecución una ventana aparece mostrando una animación de procesamiento. Le recomendamos al usuario que tenga paciencia con la ejecución del proceso. Dependiendo de la complejidad y número de capas involucradas el proceso puede demorarse. La ventana de ArcMap quedará inactiva hasta que finalice todos los cortes.  

<p align="center">
 <img src="..\Images\ventana_proceso.png">
</p>

Una vez finalizado el proceso, los resultados de las capas de corte se cargarán en el dataframe de resultados tal y como se aprecia en la siguientes imágenes.

<p align="center">
 <img src="..\Images\resultados_corte.png">
</p>

<p align="center">
 <img src="..\Images\resultados_corte1.png">
</p>

Tal y como lo muestra la última imagen, las capas cortadas aparecerán cargadas con el nombre de la capa seguidas de un de un guion bajo y el nombre del molde. Este nombre es solo el nombre del layer, en la geodatabase de salida los cortes conservan el nombre original de la capa.

A continuación se detallan los resultados almacenados dentro del folder de salida:


<p align="center">
 <img src="..\Images\resultados_multicortes.png">
</p>


### Solo Corte

Si selecciona esta opción solo se realizaran los cortes de las capas, por lo tanto solo se obtendrán Feature Class.

<p align="center">
 <img src="..\Images\resultado_solo_corte.png">
</p>

### Extracción por Campo con Estadísticas

Esta opción permite obtener los mismos resultados de corte con estadísticas, la única diferencia es que no se realizan los cortes, si no, que de las capas a cortar se extraen aquellos elementos que tengan el mismo valor de atributo de las capas de molde.

<p align="center">
 <img src="..\Images\campo_de_extraccion.png">
</p>

Justo después de seleccionar la opción extracción por campo con estadísticas, aparecerá una ventana solicitando el campo de extracción. Este campo debe ser idéntico en nombre y en tipo al campo que tienen las capas a cortar. A continuación se aprecian las tablas y el campo de extracción para hacer un corte por municipios.


<p align="center">
 <img src="..\Images\extraccion_por_campo.png">
</p>

Los resultados se aprecian a continuación.

<p align="center">
 <img src="..\Images\resultados_multicortes.png">
</p>
