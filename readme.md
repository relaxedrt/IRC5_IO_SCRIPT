Antes de comenzar necesitaremos tener instalado python, lo podriamos hacer directamente desde la windows store o desde la propia
pagina web de python. Tambien deberemos instalar a través del terminal las librerias os y openpyxl utilizando el comando pip install.
Para el optimo funcionamiento de este script deberas rellenar los datos de forma correcta en el archivo MapExample.xlsx,
en la columna nombre escribiremos la columna nombre del excel oficial, en la columna canal escribiremos el numero que tenga
en la columna x del original (Si es una GO o una GI deberemos poner de que numero a que numero va), en tipo escribiremos DI, DO, GI o GO
dependiendo que tipo de variable sea y en la columna unidad escribiremos el nombre de la tarjeta fisica de entradas salidas o
la tarjeta de profinet, que sería (PN_Internal_Device).
El script te hará unas preguntas antes de crear el archivo.
Como quieres que se llame el archivo? Con esta pregunta tomaremos el nombre del archivo que crearemos.
Como se llama el excel? Recuerda que tiene que estar en la carpeta. Normalmente se llama MapExample.xlsx a no ser que lo hayas tocado
Cuantas entradas vamos a convertir? El numero de entradas que vamos a convertir
Cuantas salidas vamos a convertir? El numero de salidas que vamos a conovertir