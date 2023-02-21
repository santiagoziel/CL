Archivo Excel:
	Se deben copiar la hoja "SheetNames" y la hoja "backend" del archvio Sheets al archivo objetivo

Macros:
	Hay dos tipos de macros necesesarios, macros de hoja y macors de modulos, se encuentran separados por carpetas y la diferencia es en donde debes copiar el texto.

	Macros de hoja:
		Los Macros de hojas son pertinentes a unicamente a una hoja del archivo.
		Para pegarlos abre Visual basic desde el menu de Developer (Desarollador), en la parte izquierda encontraras el menu de objetos de excel, te debe aparecer como una carpeta con las hojas que contiene el archvio, el nombre de la hoja aparece entre parentecis.
		Cada archvio denotro de la carpeta macros_hojas debe corresponder en nombre a una hoja de tu archivo, haz doble click en la hoja pertinenete y en la ventana que salga pega el contenido del archivo de texto.

	Macros de modulo
		Los Macros de modulo son pertinentes a todo el arcvhio por lo tanto deben ir en un modulo.
		La carpeta macors_modulos contiene 2 archvios, ejecutables y otro llamado funciones.
		Debemos crear dos modulos para recibir su contenido.

		PAra crear un modulo, Click Derecho en el menu de objetos de Excel (Micorsoft Excel Objects) es la carpeta donde se encuentran las hojas del archivo.
		Despues de click Derecho, clcik en incertar y finalmente en modulo.
		Una Vez creado, en la aprte izquierda pero ams abajo debe salirte un pequeño menu que diga algo parecido a "(Name) Module1" no es necesario pero recomiendo cambiar los nombres de acurdo a cual de los archvios pegaste en cada uno

		toma el contenido de cada archivo y pegalo en su modulo ocrrespondiente.

		El archvio de Ejecutables contine todos los macros que deberas llamar directamente, asi como una pequeña explicacion de que hacen cuando es necesario

Una vez echo esto tu Archivo esta Casi listo, solo hace falta una cosa.

El primer MAcro en el modulo de Ejecutables tiene el nombre de setup, este se asegurara de preparar cirtas cosas, pore favor correlo, deberas ver un mensaje que diga "SET UP Done" cuando termine.

Despeus de esto tu arvho esta listo, recuerda guardarlo como un archivo excel compatible con macros.

En la Hoja de "Cuentas" rango R4:S5, encontraras un pequeño menu de configurar para ciertos comportamientos del macro "formatear_archivo"
		


 