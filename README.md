# ProcesarCartolaCuentaRUT2-2024
Programa en python que recibe las cartolas de CuentaRUT enviadas por correo a los usuarios de BancoEstado. Procesa los movimientos durante el 2do semestre de 2024. Exporta un archivo excel con la información sobre todas las entradas de dinero en la cuenta en ese plazo. De esta forma se puede monitorear la cantidad de ingresos desde distintas personas, calculando y mostrando la cantidad de personas distintas que transfieren o depositan en plazos mensuales o semestrales. 
Sólo considerará los movimientos hechos entre el 01 de Julio de 2024 y el 31 de Diciembre de 2024. 

Creado en función de las nuevas normas del proyecto de ley publicadas por el SII el 04 de octubre de 2024, específicamente la norma N°12 sobre la obligatoriedad de los bancos a entregar información sobre las cuentas que reciban 50 o más transeferencias al mes de personas distintas, o 100 al semestre, considerando el segundo semestre del año 2024 como el primer rango de tiempo a ser informado (léase en la [página oficial del SII](https://www.sii.cl/noticias/2024/031024noti03srm.htm)).

Sólo para Windows.

ESTE PROGRAMA QUEDARÁ OBSOLETO EL DÍA 01/01/2025.

# Instrucciones de uso:
#### El programa puede ejecutarse de dos maneras:
#### 1- Forma fácil y rápida: 
#### Descargar el .exe y ejecutar el programa sólo haciéndole click (enfocado a usuarios comunes).
 
2- Descargando el código fuente en .py, para correrlo en consola y/o modificar el código (enfocado a desarrolladores o gente curiosa). 
# Para ejecutar en sólo un click (Forma fácil y rápida):
### Paso 1: 
Descarga las cartolas de la CuentaRUT que envía BancoEstado a tu e-mail, para encontrarlas fácilmente escribe lo siguiente en el buscador de tu correo:

`cartola de cuentarut - le adjuntamos `

Ingresando esto se mostrarán todos los archivos de cartola que has recibido. 

El programa sólo considerará los registros desde el 01 de julio en adelante, así que no servirá descargar información de meses anteriores.

### Paso 2:
Descarga el archivo ejecutable "ProcesadorCartolas.exe" ([click aquí para descargarlo](https://github.com/sebsatian/ProcesarCartolaCuentaRUT2-2024/raw/refs/heads/main/ProcesadorCartolas.exe))

### Paso 3:
Crea una carpeta e ingresa todas las cartolas descargadas, junto al archivo "ProcesadorCartolas.exe" que acabas de descargar.

### Paso 4:
Haz doble click en el archivo ejecutable "ProcesadorCartolas.exe" (si te sale un anuncio de Windows, sólo presiona "Más información" y "Ejecutar de todas formas"), se abrirá el programa y te pedirá la contraseña de los archivos. Por defecto, la contraseña que BancoEstado establece para estos archivos serán los últimos 4 números de tu rut antes del guión (-).

`Si tu RUT es 12.345.678-9 ----> Tu contraseña será 5678 `

Y listo! dentro de la misma carpeta en que ubicaste todos los archivos se creará un archivo de Excel con el nombre "InfoTransferencias2-2024.xlsx", donde se encontrará toda la información sobre los ingresos ordenada por fecha y separada por meses y semestre.

# Descarga y ejecución para desarrolladores:

Teniendo python instalado, instala las siguientes dependencias:

`pip install pdfplumber pandas openpyxl `

y léete el código para saber más lol


