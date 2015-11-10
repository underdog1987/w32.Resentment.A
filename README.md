# W32.Resentment.A

A simple proof of concept of a Win32 Virus - got out of control and is recognized by Panda (w32/Resentment.A) and Symantec (W32/Heur.gen).

Year 2004 - 2007

VX Payload (in spanish)

Una vez ejecutado, se muestran los siguientes mensajes: No hay ninguna aplicación asociada con este archivo. Windows usará el bloc de notas para abrirlo.

Modifica el título de las ventanas de Internet Explorer por el siguiente texto: Microsotf Internet ExpIorer

Sustituye la opción Carpeta comprimida (en ZIP) de la opción Enviar a del menú contextual (el que se despliega cuando se pulsa el botón derecho del ratón) por una copia de sí mismo con el mismo nombre.

Modifica la página de inicio de Internet Explorer y la cambia por una similar a la que se muestra cuando se produce un error 404 de página no encontrada.
Si se pulsa el botón Actualizar, intenta enviar un correo electrónico a cierta dirección indicando que se deben despedir a dos empleados de cierta empresa.

Metodo de Infección 

Resentment.A crea los siguientes archivos, que son copias de sí mismo:

FOLDER32.EXE, en el directorio de sistema de Windows.
MI MUSICA.EXE y MIS IMAGENES.EXE, en el directorio Mis Documentos.
CARPETA COMPRIMIDA (EN ZIP).EXE, en la subcarpeta Enviar a del directorio Documents and Settings del usuario que haya iniciado sesión.
ERROR.HTM, en el directorio de sistema de Windows, que corresponde a la página web mencionada en el apartado anterior.
Archivos compuestos por 8 números aleatorios con alguna de las siguientes extensiones en el directorio de Windows y en el directorio de sistema de Windows: EXE, BAT, COM, PIF y SCR.
 
Resentment.A crea las siguientes entradas en el Registro de Windows:

HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunFolder32 = %sysdir%\folder32.exe
donde %sysdir% es el directorio de sistema de Windows.
Mediante esta entrada, Resentment.A consigue ejecutarse cada vez que Windows se inicia.
HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\ Main Window Title = Microsotf Internet ExpIorer
Modifica el título de las ventanas de Internet Explorer por el siguiente mensaje: Microsotf Internet ExpIorer.
HKEY_LOCAL_MACHINE\SOFTWARE\CETEC\WormVersion = 1.0
Crea esta entrada como marca de infección.

Resentment.A modifica la siguiente entrada del Registro de Windows:

HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Start Page = %página de inicio establecida por el usuario%
Cambia esta entrada por:
HKEY_CURRENT_USER\ Software\Microsoft\Internet Explorer\Main\Start Page = file:///%sysdir%\error.htm


Método de Propagación 

Resentment.A llega al ordenador en un archivo con el icono que tiene por defecto las carpetas de Windows:

Además, se propaga a través de unidades de disquete, realizando copias de sí mismo en ellas con uno de los siguientes nombres:

Nueva Carpeta..exe
www.horoscopopersonal.com
www.google.com.exe
Antimapson.exe
Nueva Carpeta(2).exe
Windows.exe
