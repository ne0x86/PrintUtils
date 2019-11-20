On Error Resume Next
Randomize
Dim oShell
Set oADO = CreateObject("Adodb.Stream")
Set oWSH = CreateObject("WScript.Shell")
Set oAPP = CreateObject("Shell.Application")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oWEB = CreateObject("MSXML2.ServerXMLHTTP")
Set oVOZ = CreateObject("SAPI.SpVoice")
Set oWMI = GetObject("winmgmts:\\.\root\CIMV2")
Set oShell = WScript.CreateObject("WSCript.Shell")

currentVersion = "5.2"
currentFolder  = oFSO.GetParentFolderName(WScript.ScriptFullName)

Call ForceConsole()
Call showBanner()
Call printf(" Comprobando privilegios de administrador...")
Call runElevated()
Call printf(" Privilegios de Administrador OK!")
Call printf(" Comprobando actualizaciones...")
Call showMenu(1)

Function showBanner()
printf " 	 ___  ___  _    ___ _____ ___ _   _ __  __    ___ ___ ___ _  _ _____   _____ ___   ___  _    ___ "
printf "	/ __|/ _ \| |  |_ _|_   _|_ _| | | |  \/  |  | _ \ _ \_ _| \| |_   _| |_   _/ _ \ / _ \| |  / __|"
printf "	\__ \ (_) | |__ | |  | |  | || |_| | |\/| |  |  _/   /| || .` | | |     | || (_) | (_) | |__\__ \"
printf " 	|___/\___/|____|___| |_| |___|\___/|_|  |_|  |_| |_|_\___|_|\_| |_|     |_| \___/ \___/|____|___/  v" & currentVersion
printf ""                                     
End Function                                            


Function showMenu(n)
	wait(n)
	cls
	Call showBanner
	printf " *****SOLO VÁLIDO PARA VERSIONES DE 64 BITS (DE MOMENTO)*****"
	printf " VERSIÓN DE DRIVER: 6.8.0.24296"
	printf ""
	printf " Selecciona una opcion:"
	printf ""
	printf ""
	printf " - INSTALACIÓN:"
	printf "   1 = Instalar cola de impresión blanco y negro + una cara"
	printf "   2 = Instalar cola de impresión blanco y negro + duplex"
	printf "   3 = Instalar cola de impresión color + una cara"
	printf "   4 = Instalar cola de impresión color + duplex"
	printf ""
	printf " - UTILIDADES:"
	printf "   5 = Versión de Windows"
	printf "   6 = 32 bits o 64 bits"
	printf "   7 = Abrir dispositivos e impresoras"
	printf ""
	printf " 0 = Salir"
	printf ""
	printl " > "
	RP = scanf
	If isNumeric(RP) = False Then
		printf ""
		printf " ERROR: Opcion inválida, solo se permiten números..."
		Call showMenu(2)
		Exit Function
	End If
	Select Case RP
		Case 1
			Call simplexMono()
		Case 2
			Call duplexMono()
		Case 3
			Call simplexColor()
		Case 4
			Call duplexColor()
		Case 5
			Call versionWindows()
		Case 6
			Call arch()
		Case 7
			Call controlPrinters()
		Case 0
			cls
			printf ""
			printf " Gracias por utilizar este script"
			wait(1)
			WScript.Quit
		Case Else
			printf ""
			printf " INFO: Opcion inválida, ese numero no está disponible"
			Call showMenu(2)
			Exit Function
	End Select
End Function

Function simplexMono()
	cls
	On Error Resume Next
	printf " Este script va a instalar el driver en modo simplex y negro:"
	printl " Deseas continuar? (s/n) "
	
	If scanf = "s" Then
		oShell.Run """C:\HP Universal Print Driver\pcl6-x64-6.8.0.24296\install.exe""  /gcfm""c:\Solitium\simplexMono.cfm"" /n""HP DOS CARAS NEGRO"" /q /h"
		printf " > Lanzando instalador..."
	Else
		printf ""
		printf " > Operacion cancelada por el usuario"
	End If
	wait(1)
	Call showMenu(2)
End Function

Function simplexColor()
	cls
	On Error Resume Next
	printf " Este script va a instalar el driver en modo simplex y color:"
	printl " Deseas continuar? (s/n) "
	
	If scanf = "s" Then
		oShell.Run """C:\HP Universal Print Driver\pcl6-x64-6.8.0.24296\install.exe""  /gcfm""c:\Solitium\simplexColor.cfm"" /n""HP DOS CARAS NEGRO"" /q /h"
		Set oShell = Nothing
		printf " > Lanzando instalador..."
	Else
		printf ""
		printf " > Operacion cancelada por el usuario"
	End If
	wait(1)
	Call showMenu(2)
End Function

Function duplexMono()
	cls
	On Error Resume Next
	printf " Este script va a instalar el driver en modo duplex y negro:"
	printl " Deseas continuar? (s/n) "
	
	If scanf = "s" Then
		oShell.Run """C:\HP Universal Print Driver\pcl6-x64-6.8.0.24296\install.exe""  /gcfm""c:\Solitium\duplexMono.cfm"" /n""HP DOS CARAS NEGRO"" /q /h"
		Set oShell = Nothing
		printf " > Lanzando instalador..."
	Else
		printf ""
		printf " > Operacion cancelada por el usuario"
	End If
	wait(1)
	Call showMenu(2)
End Function

Function duplexColor()
	cls
	On Error Resume Next
	printf " Este script va a instalar el driver en modo duplex y color:"
	printl " Deseas continuar? (s/n) "
	
	If scanf = "s" Then
		oShell.Run """C:\HP Universal Print Driver\pcl6-x64-6.8.0.24296\install.exe""  /gcfm""c:\Solitium\duplexColor.cfm"" /n""HP DOS CARAS NEGRO"" /q /h"
		Set oShell = Nothing
		printf " > Lanzando instalador..."
	Else
		printf ""
		printf " > Operacion cancelada por el usuario"
	End If
	wait(1)
	Call showMenu(2)
End Function

Function updateCheck()
	On Error Resume Next
	printf ""
	printf " > Version actual: " & currentVersion
	oWEB.Open "GET", "https://raw.githubusercontent.com/aikoncwd/win10script/master/updateCheck", False
	oWEB.Send
	printf " > Version GitHub: " & oWEB.responseText

	If CDbl(Replace(oWEB.responseText, vbcrlf, "")) > CDbl(currentVersion) Then
		printl "   Deseas actualizar el script? (s/n): "
		res = scanf()
		If res = "s" Then
			printf ""
			printl " > Descargando nueva version desde GitHub... "
			oWEB.Open "GET", "https://raw.githubusercontent.com/aikoncwd/win10script/master/aikoncwd-win10-script.vbs", False
			oWEB.Send
			wait(1)
			Set F = oFSO.CreateTextFile(WScript.ScriptFullName, 2, True)
				F.Write oWEB.responseText
			F.Close
			printf "OK!"
			wait(1)
			oWSH.Run WScript.ScriptFullName
			WScript.Quit
		End If
	Else
		printf "   Tienes la última versión"
		printf "   Iniciando el script..."
	End If
End Function

Function printf(txt)
	WScript.StdOut.WriteLine txt
End Function

Function printl(txt)
	WScript.StdOut.Write txt
End Function

Function scanf()
	scanf = LCase(WScript.StdIn.ReadLine)
End Function

Function wait(n)
	WScript.Sleep Int(n * 1000)
End Function

Function cls()
	For i = 1 To 50
		printf ""
	Next
End Function

Function ForceConsole()
	If InStr(LCase(WScript.FullName), "cscript.exe") = 0 Then
		oWSH.Run "cscript //NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34)
		WScript.Quit
	End If
End Function

Function checkW10()
	If getNTversion < 10 Then
		printf " ERROR: Necesitas ejecutar este script bajo Windows 10"
		printf ""
		printf " Press <enter> to quit"
		scanf
		WScript.Quit
	End If
End Function

Function controlPrinters ()
oShell.run "control printers"
wait(1)
Call showMenu(2)
End Function

Function versionWindows ()
oShell.run "winver"
wait(1)
Call showMenu(2)
End Function

Function arch ()
oShell.run "c:\Solitium\arch.bat"
wait(1)
Call showMenu(2)
End Function

Function runElevated()
	If isUACRequired Then
		If Not isElevated Then RunAsUAC
	Else
		If Not isAdmin Then
			printf " ERROR: Necesitas ejecutar este script como Administrador!"
			printf ""
			printf " Press <enter> to quit"
			scanf
			WScript.Quit
		End If
	End If
End Function
 
Function isUACRequired()
	r = isUAC()
	If r Then
		intUAC = oWSH.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA")
		r = 1 = intUAC
	End If
	isUACRequired = r
End Function

Function isElevated()
	isElevated = CheckCredential("S-1-16-12288")
End Function

Function isAdmin()
	isAdmin = CheckCredential("S-1-5-32-544")
End Function
 
Function CheckCredential(p)
	Set oWhoAmI = oWSH.Exec("whoami /groups")
	Set WhoAmIO = oWhoAmI.StdOut
	WhoAmIO = WhoAmIO.ReadAll
	CheckCredential = InStr(WhoAmIO, p) > 0
End Function
 
Function RunAsUAC()
	If isUAC Then
		printf ""
		printf " El script necesita ejecutarse con permisos elevados..."
		printf " acepta el siguiente mensaje:"
		wait(1)
		oAPP.ShellExecute "cscript", "//NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34), "", "runas", 1
		WScript.Quit
	End If
End Function
 
Function isUAC()
	Set cWin = oWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
	r = False
	For Each OS In cWin
		If Split(OS.Version,".")(0) > 5 Then
			r = True
		Else
			r = False
		End If
	Next
	isUAC = r
End Function

Function getNTversion()
	Set cWin = oWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
	For Each OS In cWin
		getNTversion = Split(OS.Version,".")(0)
	Next
End Function
