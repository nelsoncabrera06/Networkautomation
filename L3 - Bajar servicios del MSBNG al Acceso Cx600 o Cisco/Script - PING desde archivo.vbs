# $language = "VBScript"
# $interface = "1.0"

' Toma como input el TXT "Datos de entrada.txt"
' Los resultados obtenidos los guarda en "Resultados\OUTPUT - Prueba de ping.txt"
' El formato a pegar en el TXT para que funcione bien el SCRIPT debe ser como el siguiente ejemplo:
' ping   181.111.250.193
' ping -vpn-instance pami-vpn     181.15.22.65    (HUAWEI)
' ping vrf pami-vpn  172.17.1.102 (CISCO)
' por ahora funciona solo para HUAWEI lo voy a modificar tmb para CISCO

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const BUTTON_YESNO = 4		' Yes and No buttons
Const IDYES = 6		' Yes button clicked
Const IDNO = 7		' No button clicked
NombreArchivo_Entrada = "Input.txt"
NombreArchivo_Salida = "OUTPUT - Prueba de ping.txt"


Sub Main

  crt.Screen.Synchronous = True
  Dim waitStrs, comandostring
  Dim row, screenrow, readline, items


  Dim fso, file
  Set fso = CreateObject("Scripting.FileSystemObject")

  Dim fsodatos, filedatos, str
  Set fsodatos = CreateObject("Scripting.FileSystemObject")

  Set filedatos = fsodatos.OpenTextFile(NombreArchivo_Entrada, ForReading, False)
 
  Set file = fso.OpenTextFile("OUTPUT - Prueba de ping.txt", ForWriting, True)

  row = 1 
  waitStrs = Array( Chr(10), "% packet loss" , "Success rate")		'Chr(10) = retorno de carro
  siguientecomando = Array( ">", "#" )
	ResultadoPING = "Sin Analizar"
	
    crt.Screen.Synchronous = True
	
	SEGUIR = funcion_AbrirINPUT(file, filedatos, index, CVLAN, SVLAN, ESTADO)
	if SEGUIR = 1 then
		MsgBox("Salida del script") 
		exit sub
	end if
	
	crt.Screen.Send Chr(13)

    Do While filedatos.AtEndOfStream <> True ' hacer mientras no haya terminado de leer el archivo
	  str = filedatos.Readline					' lee la linea ejemplo: ping   181.111.250.193
	  crt.Screen.Send str & Chr(13)				' envia la linea leida como un comando y lo ejecuta
	  file.Write str & VBTab 					' escribe el comando leido en el TXT 
	  
	  Do While True								' este sería un while infinito
          result = crt.Screen.WaitForStrings( waitStrs )   ' cuando lee alguno de los casos en waitStrs lo guarda en result
	  			
		Select Case result
			Case 2 ' % packet loss HUAWEI
				screenrow = crt.screen.CurrentRow
				readline = crt.Screen.Get(screenrow, 5, screenrow, 50 )		' guarda el resultado leido en readline
				readline = Trim(readline)    ' la funcion trim le saca los espacios al string al principio y al final 
				
				if readline = "100.00% packet loss" then 
					ResultadoPING = "No llega"
				else
					ResultadoPING = "ok"
				end if
				'MsgBox readline 
				
				file.Write readline & VBTab & ResultadoPING & vbCrLf			' escribe el resultado leido en el archivo de salida
				
				'file.Write readline & vbCrLf			' escribe el resultado leido en el archivo de salida
				Exit Do
			Case 3  ' Success rate CISCO
				crt.Screen.WaitForString ")"
				screenrow = crt.screen.CurrentRow
				readline = crt.Screen.Get(screenrow, 1, screenrow, 50 )		' guarda el resultado leido en readline
				readline = Trim(readline)    ' la funcion trim le saca los espacios al string al principio y al final 
				if readline = "Success rate is 0 percent (0/5)" then 
					ResultadoPING = "No llega"
				else
					ResultadoPING = "ok"
				end if
				'MsgBox readline 
				file.Write readline & VBTab & ResultadoPING & vbCrLf			' escribe el resultado leido en el archivo de salida
				'file.Write readline & vbCrLf			' escribe el resultado leido en el archivo de salida
				Exit Do				
		End Select

       'Wend									' fin del while
	  Loop										' fin del do while 
	  'crt.Screen.WaitForStrings (">", "#")							' espero hasta poder volver a tirar el siguiente comando
	  crt.Screen.WaitForStrings (siguientecomando) '"#"	 			' espero hasta poder volver a tirar el siguiente comando
    Loop															' repetir

  crt.screen.synchronous = false

	MsgBox("Fin del script") 
	Set g_shell = CreateObject("WScript.Shell")
	g_shell.Run chr(34) & NombreArchivo_Salida & chr(34)
	
End Sub

Function funcion_AbrirINPUT(file, filedatos, index, CVLAN, SVLAN, ESTADO)
	MsgBox "FORMATO DE EJEMPLO: ping 181.88.56.93"
	
	Set g_shell = CreateObject("WScript.Shell")
	g_shell.Run chr(34) & NombreArchivo_Entrada & chr(34) 'abro el Input.txt cargo los datos, guardo y cierro
	
	crt.Screen.Send "" & chr(13)
	verific = crt.Dialog.MessageBox("Continuar?", cap, BUTTON_YESNO)
 

	if verific = IDYES then 
		funcion_AbrirINPUT = 0
	else 
		funcion_AbrirINPUT = 1 'considero 1 para salir
	end if
	
End Function