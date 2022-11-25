# email_masivo
Aplicacion para enviar correo masivo personalizado con plantilla HTML Version 1.0

Introduccion
Esta pequeña aplicacion fue desarrollada en base a los pocos conocimientos en Python, es una aplicacion simple y facil, solo cuenta con una funcion que realiza 
la authenticacion y envia un correo electronico a todos los correos que se encuentran en la lista datos.csv la aplicacion solo cuenta con 3 simples pasos para su proceso de ejecucion. Quizas pueda trabajar en perfeccionar el codigo mas adelante y tirar una version 1.1 con mas cosas implementadas como un panel para gestionar desde una web, pero lo puede hacer cualquiera


Funcionamiento
1°- Se debe llenar la lista datos.csv (En caso de no existir se debe crear dicho archivo dentro de la ruta) la lista debe tener un encabezado ya que la primera linea no se toma en cuenta al momento de leer el archivo, el encabezado es el siguiente dato1,dato2,dato3,dato4
la cantidad de datos va a depender de la cantidad de datos que nesecite rescatar del archivo para usar en la aplicacion Ejemplo:
yo del archivo rescato el dato1 como el nombre para usarlo en la plantilla HTML como tambien el usuario y la contraseña que vendrian siendo dato3 y dato4 y el dato2 lo uso en la funcion enviar como el correo receptor

2°- Una vez lleno el listado datos.csv se debe ejecutar llenar el archivo conn con los datos del servidor de correo y cuenta para la authenticacion

3°- Ejecutar el archivo main.py y esperar que el proceso termine por completo.


Si desea clonar el codigo no tengo problemas siempre se me den los creditos correspondientes.

By:Luis'Guillermo
