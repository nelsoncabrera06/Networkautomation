# $language = "VBScript"
# $interface = "1.0"

' este script entra a equipos CISCO MSBNG, y tambien a los Huawei CX600, entra a la subinterface
' la analiza
' guarda los datos en variables 
' imprime todas las variables necesarias para armar el servicio

' Agregar en casos futuros: 
'	_ redes que esta compartiendo el neighbor en HUAWEI
' 	  ver los ping en el caso de VRF HUAWEI

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const BUTTON_YESNO = 4		' Yes and No buttons
Const IDYES = 6				' Yes button clicked
Const IDNO = 7				' No button clicked
NombreArchivo_Entrada = "Input.txt"
NombreArchivo_Salida = "OUTPUT - Analisis de Subinterfaces a borrar.txt"
NombreArchivo_TABLADATOS = "OUTPUT - comandos para BORRAR subint.txt"
Dim Estaticas_array
Dim BGP_array
Dim Subinterfaces_array

Sub Main

  crt.Screen.Synchronous = True
  Dim waitStrs, comandostring
  Dim row, screenrow, readline, items
	Dim pass, USUARIO

  Dim fso, file
  Set fso = CreateObject("Scripting.FileSystemObject")

  Dim fsodatos, filedatos, str
  Set fsodatos = CreateObject("Scripting.FileSystemObject")

  Set filedatos = fsodatos.OpenTextFile(NombreArchivo_Entrada, ForReading, False)
  Set file = fso.OpenTextFile(NombreArchivo_Salida, ForWriting, True)
  Set fileoutTabladeDatos = fso.OpenTextFile(NombreArchivo_TABLADATOS, ForWriting, True)
  
  row = 1 
  ESTADO = 0
  RESULTADO = 0
  MARCA = "sin datos"
  Salir = 0
  waitStrs = Array( Chr(10), "...")
  index = 0
  
    crt.Screen.Synchronous = True
	crt.Screen.Send Chr(13)
	
	SEGUIR = funcion_AbrirINPUT(file, filedatos, index, CVLAN, SVLAN, ESTADO)
	if SEGUIR = 1 then
		MsgBox("Salida del script") 
		exit sub
	end if
	
	crt.Screen.Send Chr(13)
	ESTADO = crt.screen.WaitForStrings("Username:", ">", "#", "$", "...")
	Select Case ESTADO
		Case 2 'Ya estoy dentro del equipo caso HUAWEI
			ESTADO = 2 'Huawei
			crt.Screen.Send Chr(13)
			crt.screen.WaitForString ">"
			columnaActual = crt.screen.CurrentColumn
			screenrow = crt.screen.CurrentRow
			EQUIPO = crt.Screen.Get(screenrow, 2, screenrow, columnaActual-2)
			'MsgBox EQUIPO
		Case 3 'estoy dentro de un CISCO
			ESTADO = 3 'Cisco
			crt.Screen.Send Chr(13)
			crt.screen.WaitForString "#"
			columnaActual = crt.screen.CurrentColumn
			screenrow = crt.screen.CurrentRow
			EQUIPO = crt.Screen.Get(screenrow, 1, screenrow, columnaActual-2)
			lngPos = Instr(EQUIPO, ":")
			EQUIPO = Mid(EQUIPO, lngPos+1, Len(EQUIPO))
			EQUIPO = trim(EQUIPO)
			'MsgBox EQUIPO
		Case 4 'no estoy en ningun equipo
			ESTADO = 0
			USUARIO = crt.Dialog.Prompt("Ingrese su usuario")
			pass = crt.Dialog.Prompt("Ingrese su password:", "Logon Script", "", True)
	end select
	
	
    Do While filedatos.AtEndOfStream <> True
	      
		  do while ESTADO <> 7
		  Select Case ESTADO
		    Case 0 'inicio
				'EQUIPO = filedatos.Read(8) ' leo el equipo CASO CMUNECAS
				EQUIPO = crt.Dialog.Prompt("Ingrese el nombre del equipo")
				verific = crt.Dialog.MessageBox("Desea ingresar a " &EQUIPO , cap, BUTTON_YESNO)
				'MsgBox "Equipo " & str
				if verific = IDYES then
					crt.Screen.Send "ttelnet " & EQUIPO & Chr(13)
					ESTADO = crt.screen.WaitForStrings("Username:", ">", "#")', "$", "...")
				else 
					crt.Dialog.MessageBox("VUELVA A INTENTAR")
					ESTADO = 7
				end if
				'MsgBox "ESTADO " & ESTADO
			Case 1 'ingreso mi usuario
				crt.Screen.Send USUARIO & chr(13)
				crt.Screen.WaitForString "Password:"
				crt.Screen.Send pass & chr(13)
				ESTADO = crt.screen.WaitForStrings("Username:", ">", "#", "$")
			Case 2 'HUAWEI
				MARCA = "HUAWEI"
				RESULTADO = funcion_HUAWEI(file, filedatos, index, MARCA, EQUIPO, ESTADO, fileoutTabladeDatos)
				ESTADO = RESULTADO				
			Case 3 'CISCO
				MARCA = "CISCO"
			    RESULTADO = funcion_CISCO(file, filedatos, index, MARCA, EQUIPO, ESTADO, fileoutTabladeDatos)
				ESTADO = RESULTADO
			'Case 4 'no ingreso al equipo
			'	crt.Screen.Send chr(13)
			'	ESTADO = 0
			'Case 5 'escribe la IP del equipo
			'	screenrow = crt.screen.CurrentRow
		    '   readline = crt.Screen.Get(screenrow, 8, screenrow, crt.screen.CurrentColumn-4)
				'file.Write readline & vbCrLf
			'	ESTADO = crt.screen.WaitForStrings("Username:", ">", "#", "$", "...")
		   End Select
		Loop   	   
   Loop


  crt.screen.synchronous = false

	MsgBox("       Fin del script       " & vbCrLf & "Autor: Nelson Cabrera") 
	Set g_shell = CreateObject("WScript.Shell")
	g_shell.Run chr(34) & NombreArchivo_Salida & chr(34)
	
	    Set g_shell = CreateObject("WScript.Shell")
	g_shell.Run chr(34) & NombreArchivo_TABLADATOS & chr(34)
	
End Sub


Function funcion_AbrirINPUT(file, filedatos, index, CVLAN, SVLAN, ESTADO)
	MsgBox "FORMATO DE EJEMPLO: Te0/2/0/22.22010022"
	
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

Function funcion_HUAWEI(file, filedatos, index, MARCA, EQUIPO, ESTADO, fileoutTabladeDatos)
	Dim TABLA(30,10) 'sin tamaño
	Dim SVLAN(30) 'sin tamaño
	Dim Estatica(50) 
	Dim pingEstatica(50)
	Dim SalidaINTERNET(50)
	Dim MASCARA_estatica(50) 
	Dim ipPING(50)
	Dim Resultado_pingEstatica(50)
	Dim Redes_aprendidas(50)
	Cant_estaticas = 0
	Cant_redes = 0
	aux = 0
	indice = 0
	indice_SVLAN = 0
	indice_max = 0
	i_redes = 0
	Estatica(0) = "NO TIENE"
	SalidaINTERNET(0) = "NO TIENE"
	PUERTO = "NO TIENE"
	Subint = "NO TIENE"
	SUBINTERFACE = "NO TIENE"
	Estado_Subint = "NO TIENE"
	DESCRIPCION = "NO TIENE"
	BANDWIDTH = "NO TIENE"
	policy_input = "NO TIENE"
	policy_output = "NO TIENE"
	VRF = "NO TIENE"
	ip_WAN = "NO TIENE"
	mascara_WAN = "NO TIENE"
	NEXTHOP = "NO TIENE"
	CVLAN = "NO TIENE"
	RD = "NO TIENE"
	AS_actual = "NO TIENE"
	descripcion_BGP = "NO TIENE"
	BGP = "NO TIENE"
	PEAKFLOW = "NO TIENE"
	PalabrasCLAVE = array("arg", "ARG", "GMT")
	Equipo_gestor = "<" & EQUIPO & ">"
			
    
	MsgBox "esta version de script no contempla el caso para HUAWEI" 
	exit function
	
	str = filedatos.Readline 'aca lee el puerto o subinterface
	PUERTO = str
	SUBINTERFACE = str
	file.Write EQUIPO & " " & PUERTO & " Descripcion: disp int " & PUERTO & vbCrLf
	crt.Screen.Send "disp int " & str & chr(13) 'tira el comando para mostrar la descripcion
	crt.screen.WaitForString "current state :"
	columnaActual = crt.screen.CurrentColumn
	Salir = crt.screen.WaitForStrings (chr(13), Equipo_gestor, "Line protocol", "Description:")
	do while Salir <> 2	' mientras no aparezca # hacer
		screenrow = crt.screen.CurrentRow 'fila actual
		readline = crt.Screen.Get(screenrow, 1, screenrow, 200)
		readline = rtrim(readline)
		file.Write EQUIPO & " " & PUERTO & " " & readline & vbCrLf
		Salir = crt.screen.WaitForStrings (chr(13), Equipo_gestor, "Line protocol", "Description:") 'esperar por los siguiente caracteres
		Select Case Salir
			Case 2 'encontró # termino de mostrar la descripcion
				Salir = 2
				'MsgBox "aux = 2" 
			Case 3 'encontró Line protocol
				'<REC3MU>disp int Eth-Trunk6.26430002
				'Eth-Trunk6.26430002 current state : UP
				'Line protocol current state : UP 
				screenrow = crt.screen.CurrentRow
				crt.screen.WaitForString chr(13)
				Estado_Subint_phy = crt.Screen.Get(screenrow-1, columnaActual, screenrow-1, 200)
				Estado_Subint_phy = trim(Estado_Subint_phy)
				Estado_Subint = crt.Screen.Get(screenrow, 30, screenrow, 200)
				Estado_Subint = trim(Estado_Subint)
				'SUBINTERFACE = Mid(Estado_Subint, 1, lngPos-1)
				'MsgBox Estado_Subint_phy & " " & Estado_Subint
				Estado_Subint = Estado_Subint_phy & " " & Estado_Subint
			Case 4 'encontró Description:
				columnaActual = crt.screen.CurrentColumn
				screenrow = crt.screen.CurrentRow	
				crt.screen.WaitForString chr(13)
				DESCRIPCION = crt.Screen.Get(screenrow, columnaActual, screenrow, 200)
				DESCRIPCION = trim(DESCRIPCION)
				'MsgBox DESCRIPCION  ----- esto esta ok
		End Select
	Loop
	'MsgBox "termine de mostrar la descripcion"
	
	file.Write EQUIPO & " " & SUBINTERFACE & " Configuracion: disp curr int " & PUERTO & vbCrLf
	crt.Screen.Send "disp curr int " & PUERTO & chr(13) 'muestro la configuracion del puerto 
	'aca deberia contemplar el caso de Eth-Trunk y puerto Gi	
	crt.screen.WaitForString "#"	
	Salir = crt.screen.WaitForStrings (chr(13), Equipo_gestor, "#")
	do while Salir <> 2
		'MsgBox "estoy en el do while"
		screenrow = crt.screen.CurrentRow
		readline = crt.Screen.Get(screenrow, 1, screenrow, 200)
		readline = rtrim(readline)
		'MsgBox readline
		file.Write EQUIPO & " " & SUBINTERFACE & " " & readline & vbCrLf
		Salir = crt.screen.WaitForStrings(chr(13), Equipo_gestor, "bandwidth ", "inbound", "outbound", "vpn-instance ", "ip address ", "qinq termination ", "interface ", "description ", "encapsulation dot1q", "flow ipv4 monitor")
		Select Case Salir
		
		' interface Eth-Trunk6.295211	
		' description CONEXION-M A VPNSPREC2 Gi1/1.11 - VOZ - CVLAN11 (5626953) - AVD1MS 6/0/45
		' set flow-stat interval 10
		' control-vid 12 qinq-termination
		' qinq termination pe-vid 2952 ce-vid 11	' VLAN y CVLAN  ---------->ok
		' ip binding vpn-instance vpn-personal-voz					---------->ok
		' ip address 10.100.6.89 255.255.255.252
		' arp broadcast enable
		' trust upstream principal
		' trust 8021p
		' statistic enable
		
		Case 3 ' bandwidth
			columnaActual = crt.screen.CurrentColumn
			screenrow = crt.screen.CurrentRow
			crt.screen.WaitForString chr(13)
			BANDWIDTH = crt.Screen.Get(screenrow, columnaActual, screenrow, 200)
			BANDWIDTH = trim(BANDWIDTH)
			'MsgBox BANDWIDTH
		Case 4 ' inbound
			columnaActual = crt.screen.CurrentColumn
			screenrow = crt.screen.CurrentRow
			policy_input = crt.Screen.Get(screenrow, 13, screenrow, columnaActual-9)
			policy_input = trim(policy_input)
			'MsgBox "in " & policy_input
			crt.screen.WaitForString chr(13)
		Case 5 ' outbound
			columnaActual = crt.screen.CurrentColumn
			screenrow = crt.screen.CurrentRow
			policy_output = crt.Screen.Get(screenrow, 13, screenrow, columnaActual-9)
			policy_output = trim(policy_output)
			'MsgBox "out " & policy_input
			crt.screen.WaitForString chr(13)
		case 6 ' vpn-instance
			columnaActual = crt.screen.CurrentColumn
			screenrow = crt.screen.CurrentRow
			crt.screen.WaitForString chr(13)
			VRF = crt.Screen.Get(screenrow, columnaActual, screenrow, 100)
			VRF = trim(VRF)
			'MsgBox "VRF " & VRF
		Case 7 ' ip address 10.100.6.89 255.255.255.252
			columnaActual = crt.screen.CurrentColumn
			screenrow = crt.screen.CurrentRow
			crt.screen.WaitForString " 255."
			columnaFinal = crt.screen.CurrentColumn - 5
			ip_WAN = crt.Screen.Get(screenrow, columnaActual, screenrow, columnaFinal)
			crt.screen.WaitForString chr(13)
			columnaActual = columnaFinal
			mascara_WAN = crt.Screen.Get(screenrow, columnaActual, screenrow, 200)
			mascara_WAN = trim(mascara_WAN)
			'MsgBox ip_WAN & " mascara: " & mascara_WAN
		Case 8 ' qinq termination
			' qinq termination pe-vid 2952 ce-vid 11	' VLAN y CVLAN
			columnaActual = crt.screen.CurrentColumn
			screenrow = crt.screen.CurrentRow
			crt.screen.WaitForString chr(13)
			CVLAN = crt.Screen.Get(screenrow, columnaActual, screenrow, 200)
			CVLAN = trim(CVLAN)
			lngPos = Instr(CVLAN, "ce-vid")
			if lngPos > 0 then
				CVLAN = Mid(CVLAN, lngPos+6, Len(CVLAN))
				CVLAN = trim(CVLAN)
			end if
			'MsgBox "CVLAN " & CVLAN
		Case 9 ' Interface
			columnaActual = crt.screen.CurrentColumn
			screenrow = crt.screen.CurrentRow
			crt.screen.WaitForString "."
			columnaFinal = crt.screen.CurrentColumn - 2
			PUERTO = crt.Screen.Get(screenrow, columnaActual, screenrow, columnaFinal)
			crt.screen.WaitForString chr(13)
			columnaActual = columnaFinal + 2
			Subint = crt.Screen.Get(screenrow, columnaActual, screenrow, 50)
		Case 10 ' Description
			columnaActual = crt.screen.CurrentColumn
			screenrow = crt.screen.CurrentRow
			crt.screen.WaitForString chr(13)
			columnaFinal = crt.screen.CurrentColumn
			DESCRIPCION = crt.Screen.Get(screenrow, columnaActual, screenrow, 300)
			DESCRIPCION = trim(DESCRIPCION)
			'MsgBox DESCRIPCION
		Case 11 ' encapsulation dot1q -> para caso MSBNG ejemplo RSC1MB
			' encapsulation dot1q 294 second-dot1q 2
			lngPos = 0
			columnaActual = crt.screen.CurrentColumn
			screenrow = crt.screen.CurrentRow
			crt.screen.WaitForString chr(13)
			CVLAN = crt.Screen.Get(screenrow, columnaActual, screenrow, 200)
			CVLAN = trim(CVLAN)
			lngPos = Instr(CVLAN, "second-dot1q")
			if lngPos > 0 then
				CVLAN = Mid(CVLAN, lngPos+13, Len(CVLAN))
				CVLAN = trim(CVLAN)
			end if
			'MsgBox CVLAN
		Case 12 'flow ipv4 monitor FMMAP sampler FNF_SAMPLER_MAP ingress
			PEAKFLOW = "Si tiene"
			End Select
	loop
	'MsgBox "termine de leer la configuracion de la subinterface" 
		
	' ahora tengo que ver la configuracion de la VRF
	if VRF <> "NO TIENE" then
	file.Write EQUIPO & " " & SUBINTERFACE & " Config Actual VRF: disp curr conf vpn-instance " & VRF & vbCrLf
	crt.Screen.Send "disp curr conf vpn-instance " & VRF & chr(13) 'muestro la configuracion de la VRF
	'crt.screen.WaitForStrings PalabrasCLAVE  'espera la palabra arg			
	Salir = crt.screen.WaitForStrings (chr(13), Equipo_gestor)
	do while Salir <> 2
		screenrow = crt.screen.CurrentRow
		readline = crt.Screen.Get(screenrow, 1, screenrow, 200)
		readline = rtrim(readline)
		'MsgBox readline
		file.Write EQUIPO & " " & SUBINTERFACE & " " & readline & vbCrLf
		Salir = crt.screen.WaitForStrings(chr(13), Equipo_gestor, "route-distinguisher")
		Select Case Salir
			Case 3 ' 7303 encontré el RD
				columnaActual = crt.screen.CurrentColumn
				screenrow = crt.screen.CurrentRow
				crt.screen.WaitForString chr(13)
				RD = crt.Screen.Get(screenrow, columnaActual, screenrow, 100)
				RD = trim(RD)
				'MsgBox "RD " & RD		
			End Select
	loop
	end if
	'MsgBox "termine de leer la configuracion de la VRF"
		
	
	' IP de WAN+1
	'Calculo de Siguiente IP 
			
	if ip_WAN <> "NO TIENE" then
		NEXTHOP = SiguienteIP(ip_WAN)
		'MsgBox "NEXTHOP " & NEXTHOP
	end if
	
	' ahora tengo que ver la configuracion del BGP
	' MsgBox ip_WAN
	if ip_WAN <> "NO TIENE" then
	if VRF <> "NO TIENE" then
	BGP = "NO TIENE"
	file.Write EQUIPO & " " & SUBINTERFACE & " Config Actual BGP: disp curr conf bgp | begin " & VRF & vbCrLf
	crt.Screen.Send "disp curr conf bgp | begin " & VRF & chr(13) 'muestro la configuracion del BGP
	Salir = crt.screen.WaitForStrings (chr(13), "#", "as-number ", "description ", "import-route ")
	do while Salir <> 2
		screenrow = crt.screen.CurrentRow
		readline = crt.Screen.Get(screenrow, 1, screenrow, 200)
		readline = rtrim(readline)
		'MsgBox readline
		file.Write EQUIPO & " " & SUBINTERFACE & " " & readline & vbCrLf
		Salir = crt.screen.WaitForStrings (chr(13), "#", "as-number ", "description ", "import-route ")
		Select Case Salir
			Case 3 ' peer 10.100.6.90 as-number 65001
				columnaActual = crt.screen.CurrentColumn 
				screenrow = crt.screen.CurrentRow
				crt.screen.WaitForString chr(13)
				linealeida = crt.Screen.Get(screenrow, 1, screenrow, 150)
				'MsgBox "linealeida " & linealeida
				lngPos = Instr(linealeida, NEXTHOP)
				if lngPos > 0 then
					AS_actual = Mid(linealeida, columnaActual, Len(linealeida))
					AS_actual = trim(AS_actual)
					'MsgBox AS_actual
				end if
				BGP = "Si tiene"
			Case 4 ' description BGP 
				columnaActual = crt.screen.CurrentColumn
				screenrow = crt.screen.CurrentRow
				crt.screen.WaitForString chr(13)
				linealeida = crt.Screen.Get(screenrow, 1, screenrow, 150)
				lngPos = Instr(linealeida, NEXTHOP)
				if lngPos > 0 then
					descripcion_BGP = Mid(linealeida, columnaActual, Len(linealeida))
					descripcion_BGP = trim(descripcion_BGP)
					'MsgBox descripcion_BGP
				end if
				BGP = "Si tiene"
			Case 5 'import-route
				BGP = "Si tiene"
			End Select
	loop
	else 'si no tiene VRF
		' habria que probar este caso!!!!!!!!!!!!!!!!!!!
		'<REC3MU>disp curr conf bgp | begin 181.15.6.134
		'peer 181.15.6.134 as-number 64512
		'peer 181.15.6.134 description INTEGRA - BOLZAN JUAN ALBERTO - Ref 881408 - Lin 2859391 - Acc 55MLB01 3/4/1
	file.Write EQUIPO & " " & SUBINTERFACE & " Config Actual BGP: disp curr conf bgp | begin " & NEXTHOP & vbCrLf
	crt.Screen.Send "disp curr conf bgp | begin " & NEXTHOP & chr(13) 'muestro la configuracion del BGP
	Salir = crt.screen.WaitForStrings (chr(13), "#", NEXTHOP)
	do while Salir <> 2
		screenrow = crt.screen.CurrentRow
		readline = crt.Screen.Get(screenrow, 1, screenrow, 200)
		readline = rtrim(readline)
		'MsgBox readline
		file.Write EQUIPO & " " & SUBINTERFACE & " " & readline & vbCrLf
		Salir = crt.screen.WaitForStrings (chr(13), "#", NEXTHOP)
		Select Case Salir
			Case 3 ' peer 10.100.6.90 as-number 65001
				columnaActual = crt.screen.CurrentColumn 
				screenrow = crt.screen.CurrentRow
				crt.screen.WaitForString chr(13)
				linealeida = crt.Screen.Get(screenrow, 1, screenrow, 150)
				'MsgBox "linealeida " & linealeida
				lngPos = Instr(linealeida, "as-number")
				if lngPos > 0 then
					AS_actual = Mid(linealeida, columnaActual+10, Len(linealeida))
					AS_actual = trim(AS_actual)
					'MsgBox "AS_actual " & AS_actual
				else 
					lngPos = Instr(linealeida, "description")
					if lngPos > 0 then
						descripcion_BGP = Mid(linealeida, columnaActual+12, Len(linealeida))
						descripcion_BGP = trim(descripcion_BGP)
						'MsgBox "descripcion_BGP " & descripcion_BGP
					end if
				end if
				BGP = "Si tiene"
			End Select
	loop
	crt.screen.WaitForString Equipo_gestor
	end if
	end if
	'MsgBox "termine de leer la configuracion de la BGP"
	
	
	' ahora tengo que ver las estaticas
	' <REC3MU>disp curr | inc 181.15.38.98 
	
	if ip_WAN <> "NO TIENE" then
		file.Write EQUIPO & " " & SUBINTERFACE & " Config Actual Estaticas: disp curr | inc " & NEXTHOP & vbCrLf
		crt.Screen.Send "disp curr | inc " & NEXTHOP & chr(13) 'muestro las estaticas			
		Salir = crt.screen.WaitForStrings (chr(13), Equipo_gestor, "ip route-static ")
		do while Salir <> 2
			'if Salir = 4 or Salir = 3 then crt.screen.WaitForString chr(13)
			screenrow = crt.screen.CurrentRow
			readline = crt.Screen.Get(screenrow, 1, screenrow, 200)
			readline = rtrim(readline)
			'MsgBox readline
			file.Write EQUIPO & " " & SUBINTERFACE & " " & readline & vbCrLf
			Salir = crt.screen.WaitForStrings (chr(13), Equipo_gestor, "ip route-static ")
			Select Case Salir
				Case 3 ' "ip route-static "
				screenrow = crt.screen.CurrentRow
				columnaActual = crt.screen.CurrentColumn
				crt.screen.WaitForString " "
				columnaFinal = crt.screen.CurrentColumn
				'Provisorio = crt.Screen.Get(screenrow, columnaActual, screenrow, columnaFinal)
				Estatica(indice) = crt.Screen.Get(screenrow, columnaActual, screenrow, columnaFinal)
				ipPING(indice) = Estatica(indice)
				columnaActual = columnaFinal
				crt.screen.WaitForString " "
				columnaFinal = crt.screen.CurrentColumn	
				MASCARA_estatica(indice) = crt.Screen.Get(screenrow, columnaActual, screenrow, columnaFinal)
				MASCARA_estatica(indice) = trim(MASCARA_estatica(indice))
				'MsgBox "Estatica " & indice & " " & Estatica(indice) & " mascara " & MASCARA_estatica(indice) & " " & Len(MASCARA_estatica(indice))
				indice = indice + 1
				End Select
		loop
		Cant_estaticas = indice
		'MsgBox "Cant_estaticas " & Cant_estaticas
		indice = 0
	end if
	'MsgBox "termine de ver las estaticas" ----------------->>>>>>>>>>>>> HASTA ACA FUNCIONA REE PIOLAAAAA 20/8/19
			
	' ahora tengo que ver las estaticas que tienen salida a internet
	' sh run | inc 181.10.29.16
	' sh run | inc IP_de_LAN
	do while indice < Cant_estaticas
		if MASCARA_estatica(indice) <> "255.255.255.255" then
			file.Write EQUIPO & " " & SUBINTERFACE & " Estaticas con salida a internet: disp curr | inc " & Estatica(indice) & vbCrLf
			crt.Screen.Send "disp curr | inc " & Estatica(indice) & chr(13) 
			SalidaINTERNET(indice) = "NO TIENE"
			Salir = crt.screen.WaitForStrings (chr(13), Equipo_gestor, "RED-INTEGRA")
			if Salir = 2 then 
				SalidaINTERNET(indice) = "NO TIENE"
				indice = indice + 1
			end if
			do while Salir <> 2
				Select Case Salir
				Case 3 ' RED-INTEGRA
				SalidaINTERNET(indice) = Estatica(indice)
				'MsgBox "SalidaINTERNET " & indice & " " & SalidaINTERNET(indice)
				crt.screen.WaitForString chr(13)
				indice = indice + 1
				End Select
				screenrow = crt.screen.CurrentRow
				readline = crt.Screen.Get(screenrow, 1, screenrow, 200)
				readline = rtrim(readline)
				file.Write EQUIPO & " " & SUBINTERFACE & " " & readline & vbCrLf
				Salir = crt.screen.WaitForStrings (chr(13), Equipo_gestor, "RED-INTEGRA")
				if Salir = 2 then
					SalidaINTERNET(indice) = "NO TIENE"
					indice = indice + 1
				end if
			loop
			'MsgBox "Cant_estaticas " & Cant_estaticas
		else 
		SalidaINTERNET(indice) = "NO TIENE"
		indice = indice + 1
		end if
	loop
	indice = 0
	'MsgBox "termine de ver las estaticas con salida a internet"
						
	' ahora tengo que ver los policy
	if policy_input <> "NO TIENE" then
	'disp curr configuration qos-profile GESTION-USE-64k-TOS0
		file.Write EQUIPO & " " & SUBINTERFACE & " Config Actual Policy input: disp curr configuration qos-profile " & policy_input & vbCrLf
		crt.Screen.Send "disp curr configuration qos-profile " & policy_input & chr(13) 'muestro el policy input
		'crt.screen.WaitForStrings PalabrasCLAVE 'espera la palabra arg			
		Salir = crt.screen.WaitForStrings (chr(13), Equipo_gestor)
		do while Salir <> 2
			screenrow = crt.screen.CurrentRow
			readline = crt.Screen.Get(screenrow, 1, screenrow, 200)
			readline = rtrim(readline)
			file.Write EQUIPO & " " & SUBINTERFACE & " " & readline & vbCrLf
			Salir = crt.screen.WaitForStrings (chr(13), Equipo_gestor)
		loop
	end if
	' MsgBox "termine de ver el policy input"
	' ahora tengo que ver el policy output
	if policy_output <> "NO TIENE" then
		file.Write EQUIPO & " " & SUBINTERFACE & " Config Actual Policy output: disp curr configuration qos-profile " & policy_output & vbCrLf
		crt.Screen.Send "disp curr configuration qos-profile " & policy_output & chr(13) 'muestro el policy output
		'crt.screen.WaitForStrings PalabrasCLAVE  'espera la palabra arg				
		Salir = crt.screen.WaitForStrings (chr(13), Equipo_gestor)
		do while Salir <> 2
			screenrow = crt.screen.CurrentRow
			readline = crt.Screen.Get(screenrow, 1, screenrow, 200)
			readline = rtrim(readline)
			file.Write EQUIPO & " " & SUBINTERFACE & " " & readline & vbCrLf
			Salir = crt.screen.WaitForStrings(chr(13), Equipo_gestor)
		loop
	end if
	'MsgBox "termine de ver el policy output"
			
	' ahora tengo que tirar ping a la WAN+1
	'MsgBox "ip_WAN " & ip_WAN
	if ip_WAN <> "NO TIENE" then
	if VRF <> "NO TIENE" then
		file.Write EQUIPO & " " & SUBINTERFACE & " Config Actual Prueba de PING WAN+1: ping vrf " & VRF & " " & NEXTHOP & vbCrLf
		crt.Screen.Send "ping -vpn-instance " & VRF & " " & NEXTHOP & chr(13) 'tiro ping a la WAN+1
	else
		file.Write EQUIPO & " " & SUBINTERFACE & " Config Actual Prueba de PING WAN+1: ping " & " " & NEXTHOP & vbCrLf
		crt.Screen.Send "ping " & " " & NEXTHOP & chr(13) 'tiro ping a la WAN+1
	end if
		crt.screen.WaitForStrings "ping statistics"  'espera la palabra arg 
		Salir = crt.screen.WaitForStrings (chr(13), Equipo_gestor, "packet loss")
		do while Salir <> 2
			screenrow = crt.screen.CurrentRow
			readline = crt.Screen.Get(screenrow, 1, screenrow, 200)
			readline = rtrim(readline)
			file.Write EQUIPO & " " & SUBINTERFACE & " " & readline & vbCrLf
			Salir = crt.screen.WaitForStrings (chr(13), Equipo_gestor, "packet loss")
			Select Case Salir
			Case 3 ' packet loss
				screenrow = crt.screen.CurrentRow
				crt.screen.WaitForString chr(13)
				columnaFinal = crt.screen.CurrentColumn 
				Resultado_pingWAN = crt.Screen.Get(screenrow, 1, screenrow, 100)
				Resultado_pingWAN = trim(Resultado_pingWAN)	
				'crt.screen.WaitForString chr(13)
				'MsgBox "Resultado_pingWAN " & Resultado_pingWAN
			End Select
		loop
		file.Write EQUIPO & " " & SUBINTERFACE & vbCrLf 'solo un enter en el archivo
	end if
	'MsgBox "termino el ping a la WAN+1"
			
	' ahora tengo que tirar ping a las estaticas
	' habria que ver el caso que tiene VRF ese me falta 
	if Estatica(0) <> "NO TIENE" then
	if ip_WAN <> "NO TIENE" then
		file.Write EQUIPO & " " & SUBINTERFACE & " Config Actual Prueba de PING estaticas: " & vbCrLf
		indice = 0
		Salir = 0
		'MsgBox "MASCARA_estatica " & indice & ": " & MASCARA_estatica(indice)
		if VRF <> "NO TIENE" then
			do while Salir <> 4
				if MASCARA_estatica(indice) = "255.255.255.255" then
					crt.Screen.Send "ping -vpn-instante " & VRF & " " & ipPING(indice) & chr(13) 'tiro ping a la /32
					file.Write EQUIPO & " " & SUBINTERFACE & " " & "ping -vpn-instante " & VRF & " " & ipPING(indice) & vbCrLf
					pingEstatica(indice) = "Ping Estatica:	ping vrf " & VRF & " " & ipPING(indice) & VBTab & "ping -vpn-instance " & VRF & " " & ipPING(indice) & vbCrLf
					'crt.screen.WaitForStrings PalabrasCLAVE  'espera la palabra arg
					Salir = crt.screen.WaitForStrings (chr(13), Equipo_gestor, "packet loss")
				else 
					'Calculo de Siguiente IP 
					'MsgBox ipPING(indice)
					PING_PRUEBA = SiguienteIP(ipPING(indice))
					crt.Screen.Send "ping -vpn-instante " & VRF & " " & PING_PRUEBA & chr(13) 'tiro ping a la LAN
					file.Write EQUIPO & " " & SUBINTERFACE & " " & "ping vrf " & VRF & " " & PING_PRUEBA & vbCrLf
					pingEstatica(indice) = "Ping Estatica:	ping vrf " & VRF & " " & PING_PRUEBA & VBTab & "ping -vpn-instance " & VRF & " " & PING_PRUEBA & vbCrLf
					crt.screen.WaitForStrings PalabrasCLAVE  'espera la palabra arg					
					Salir = crt.screen.WaitForStrings (chr(13), Equipo_gestor, "packet loss")
				end if
					
					do while Salir <> 2
						screenrow = crt.screen.CurrentRow
						readline = crt.Screen.Get(screenrow, 1, screenrow, 200)
						readline = rtrim(readline)
						file.Write EQUIPO & " " & SUBINTERFACE & " " & readline & vbCrLf
						Salir = crt.screen.WaitForStrings (chr(13), Equipo_gestor, "packet loss")
						
						'Encontrado = Instr(readline, "Success")
						'if Encontrado <> 0 then
						'	Resultado_pingEstatica(indice) = Mid(readline, 1, 33)
						'	MsgBox Resultado_pingEstatica(indice)
						'end if
						
						Select Case Salir
						Case 3 ' Success
							screenrow = crt.screen.CurrentRow
							crt.screen.WaitForString chr(13)
							columnaFinal = crt.screen.CurrentColumn 
							Resultado_pingEstatica(indice) = crt.Screen.Get(screenrow, 1, screenrow, 100)
							Resultado_pingEstatica(indice) = trim(Resultado_pingEstatica(indice))	
						End Select
					loop
					

					
					indice = indice + 1
					if indice = Cant_estaticas then
						Salir = 4
					end if
			loop
		indice = 0
		else ' en caso de no tener VRF
			do while Salir <> 4
				if MASCARA_estatica(indice) = "255.255.255.255" then
					file.Write EQUIPO & " " & SUBINTERFACE & " " & "ping " & ipPING(indice) & vbCrLf
					pingEstatica(indice) = "Ping Estatica:	ping " & ipPING(indice) & VBTab & "ping " & ipPING(indice) & vbCrLf
					'MsgBox "ping a la loopback - MASCARA_estatica(indice) " & MASCARA_estatica(indice)
					crt.Screen.Send "ping " & ipPING(indice) & chr(13) 'tiro ping a la /32
					crt.screen.WaitForStrings "ping statistics"  'espera la palabra arg 
					Salir = crt.screen.WaitForStrings (chr(13), Equipo_gestor, "packet loss")
				else 
					'Calculo de Siguiente IP
					'MsgBox ipPING(indice)
					PING_PRUEBA = SiguienteIP(ipPING(indice))
					file.Write EQUIPO & " " & SUBINTERFACE & " " & "ping " & PING_PRUEBA & vbCrLf
					pingEstatica(indice) = "Ping Estatica:	ping " & PING_PRUEBA & VBTab & "ping " & PING_PRUEBA & vbCrLf
					crt.Screen.Send "ping " & PING_PRUEBA & chr(13) 'tiro ping a la LAN
					crt.screen.WaitForStrings "ping statistics"  'espera la palabra arg 				
					Salir = crt.screen.WaitForStrings (chr(13), Equipo_gestor, "packet loss")
				end if
				
				do while Salir <> 2
					screenrow = crt.screen.CurrentRow
					readline = crt.Screen.Get(screenrow, 1, screenrow, 200)
					readline = rtrim(readline)
					file.Write EQUIPO & " " & SUBINTERFACE & " " & readline & vbCrLf
					Salir = crt.screen.WaitForStrings (chr(13), Equipo_gestor, "packet loss")
					
					Select Case Salir
						Case 3 '"packet loss"
							screenrow = crt.screen.CurrentRow
							crt.screen.WaitForString chr(13)
							columnaFinal = crt.screen.CurrentColumn
							Resultado_pingEstatica(indice) = crt.Screen.Get(screenrow, 1, screenrow, 100)
							Resultado_pingEstatica(indice) = trim(Resultado_pingEstatica(indice))	
							'MsgBox Resultado_pingEstatica(indice)
							crt.screen.WaitForString chr(13)
					End Select
				loop
					

					
					indice = indice + 1
					if indice = Cant_estaticas then
						Salir = 4
					end if
				loop
				indice = 0
			end if
			end if
			end if
			' ahora tengo que imprimir las variables

			file.Write   "---------------------------------------" & vbCrLf
			file.Write   "DATOS Configuracion Actual: " & vbCrLf
			file.Write   "EQUIPO:	" & EQUIPO & vbCrLf
			file.Write   "PORT:	" & PUERTO & vbCrLf
			file.Write   "Subinterface:	" & Subint & vbCrLf
			file.Write   "Subinterface completa:	" & SUBINTERFACE & vbCrLf
			file.Write   "Estado Subinterface:	" & Estado_Subint & vbCrLf
			file.Write   "Descripcion:	" & DESCRIPCION & vbCrLf
			file.Write   "Bandwidth:	" & BANDWIDTH & vbCrLf
			file.Write   "Policy input:	" & policy_input & vbCrLf
			file.Write   "Policy output:	" & policy_output & vbCrLf
			file.Write   "VRF:	" & VRF & vbCrLf
			file.Write   "Ip de WAN:	" & ip_WAN & vbCrLf
			file.Write   "Mascara WAN:	" & mascara_WAN & vbCrLf
			file.Write   "WAN+1:	" & NEXTHOP & vbCrLf
			file.Write   "CVlan:	" & CVLAN & vbCrLf
			file.Write   "RD:	" & RD & vbCrLf
			file.Write   "BGP:	" & BGP & vbCrLf
			file.Write   "AS:	" & AS_actual & vbCrLf
			file.Write   "Peakflow:	" & PEAKFLOW & vbCrLf
			file.Write   "Descripcion BGP:	" & descripcion_BGP & vbCrLf
			do while indice < Cant_estaticas
				file.Write   "Estatica " & indice & " :	" & Estatica(indice) & vbCrLf
				if SalidaINTERNET(indice) <> "NO TIENE" then
					file.Write   "Salida a INTERNET: " & SalidaINTERNET(indice) & vbCrLf
				end if
				indice = indice + 1
			loop
			file.Write   "Resultado ping WAN+1:	" & NEXTHOP & " " & Resultado_pingWAN & vbCrLf
			indice = 0
			do while indice < Cant_estaticas
				file.Write   "Resultado Estatica " & indice & ":	" & Estatica(indice) & " " & Resultado_pingEstatica(indice) & vbCrLf
				indice = indice + 1
			loop
			i_redes = 0
			do while i_redes < Cant_redes
				file.Write Redes_aprendidas(i_redes) & vbCrLf
				i_redes = i_redes + 1
			loop
			
			fileoutTabladeDatos.Write   "---------------------------------------" & vbCrLf
			fileoutTabladeDatos.Write   "DATOS Configuracion Actual: " & vbCrLf
			fileoutTabladeDatos.Write   "EQUIPO: " & EQUIPO & vbCrLf
			fileoutTabladeDatos.Write   "PORT: " & PUERTO & vbCrLf
			fileoutTabladeDatos.Write   "Subinterface: " & Subint & vbCrLf
			fileoutTabladeDatos.Write   "Subinterface completa: " & SUBINTERFACE & vbCrLf
			fileoutTabladeDatos.Write   "Estado Subinterface: " & Estado_Subint & vbCrLf
			fileoutTabladeDatos.Write   "Descripcion: " & DESCRIPCION & vbCrLf
			fileoutTabladeDatos.Write   "Bandwidth: " & BANDWIDTH & vbCrLf
			fileoutTabladeDatos.Write   "Policy input: " & policy_input & vbCrLf
			fileoutTabladeDatos.Write   "Policy output: " & policy_output & vbCrLf
			fileoutTabladeDatos.Write   "VRF: " & VRF & vbCrLf
			fileoutTabladeDatos.Write   "Ip de WAN: " & ip_WAN & vbCrLf
			fileoutTabladeDatos.Write   "Mascara WAN: " & mascara_WAN & vbCrLf
			fileoutTabladeDatos.Write   "WAN+1: " & NEXTHOP & vbCrLf
			fileoutTabladeDatos.Write   "CVlan: " & CVLAN & vbCrLf
			fileoutTabladeDatos.Write   "RD: " & RD & vbCrLf
			fileoutTabladeDatos.Write   "BGP: " & BGP & vbCrLf 
			fileoutTabladeDatos.Write   "AS: " & AS_actual & vbCrLf
			fileoutTabladeDatos.Write   "Peakflow: " & PEAKFLOW & vbCrLf
			fileoutTabladeDatos.Write   "Descripcion BGP: " & descripcion_BGP & vbCrLf
			indice = 0
			do while indice < Cant_estaticas
				fileoutTabladeDatos.Write   "Estatica: " & Estatica(indice) & vbCrLf
				if SalidaINTERNET(indice) <> "NO TIENE" then
					fileoutTabladeDatos.Write   "Salida a INTERNET: " & SalidaINTERNET(indice) & vbCrLf
					'MsgBox "Salida a INTERNET: " & SalidaINTERNET(indice)
				else 
					fileoutTabladeDatos.Write "Salida a INTERNET: " & "NO TIENE" & vbCrLf
				end if
				indice = indice + 1
			loop
			fileoutTabladeDatos.Write   "Resultado ping WAN+1: " & NEXTHOP & " " & Resultado_pingWAN & vbCrLf
			do while indice < Cant_estaticas
				fileoutTabladeDatos.Write   "Resultado Estatica " & indice & ": " & Estatica(indice) & " " & Resultado_pingEstatica(indice) & vbCrLf
				indice = indice + 1
			loop
			
			fileoutTabladeDatos.Write "Tabla de PING:	MSBNG	CX600" & vbCrLf
			if VRF = "NO TIENE" then
				fileoutTabladeDatos.Write "Ping WAN+1:	ping " & NEXTHOP & VBTab & "ping " & NEXTHOP & vbCrLf
			else
				fileoutTabladeDatos.Write "Ping WAN+1:	ping vrf " & VRF & " " & NEXTHOP & VBTab & "ping -vpn-instance " & VRF & " " & NEXTHOP & vbCrLf
			end if
			indice = 0
			do while indice < Cant_estaticas
				fileoutTabladeDatos.Write pingEstatica(indice)
				indice = indice + 1
			loop
			i_redes = 0
			do while i_redes < Cant_redes
				fileoutTabladeDatos.Write Redes_aprendidas(i_redes) & vbCrLf
				i_redes = i_redes + 1
			loop
			
			
			if filedatos.AtEndOfStream = True then
				crt.Screen.Send "quit" & chr(13)
				crt.screen.WaitForString("$")
				ESTADO = 7
				funcion_HUAWEI = ESTADO
				Exit Function
			end if
			'CMUNECAS
			'Siguiente_EQUIPO = filedatos.Read(8) ' leo el equipo
			'RSC1MB
			'Siguiente_EQUIPO = filedatos.Read(6) ' leo el equipo
			'MsgBox "Siguiente_EQUIPO " & Siguiente_EQUIPO
			
			'if (StrComp(Siguiente_EQUIPO,EQUIPO) = 0) then 
				ESTADO = 2 'Huawei
			'else
			'	crt.Screen.Send "exit" & chr(13)
			'	crt.screen.WaitForString("$")
			'	'EQUIPO = filedatos.Read(6) ' leo el equipo
			'	'MsgBox "Equipo " & str
			'	EQUIPO = Siguiente_EQUIPO
			'    crt.Screen.Send "ttelnet " & EQUIPO & Chr(13)
	        '    'file.Write EQUIPO & VBTab
			'	ESTADO = crt.screen.WaitForStrings("Username:", ">", "#", "$", "...")
				'ESTADO = 0
			'end if
			
			
			file.Write "--------------------------------------------------------" & vbCrLf
			file.Write "--------------------------------------------------------" & vbCrLf
			'ESTADO = 0
			if filedatos.AtEndOfStream = True then ESTADO = 7
			funcion_HUAWEI = ESTADO
End Function

Function funcion_CISCO(file, filedatos, index, MARCA, EQUIPO, ESTADO, fileoutTabladeDatos)
			Dim TABLA(30,10) 'sin tamaño
			Dim SVLAN(30) 'sin tamaño
			Dim Estatica(50) 
			Dim pingEstatica(50)
			Dim SalidaINTERNET(50)
			Dim MASCARA_estatica(50) 
			Dim ipPING(50)
			Dim Resultado_pingEstatica(50)
			Dim Redes_aprendidas(50)
			Cant_estaticas = 0
			Cant_redes = 0
			aux = 0
			indice = 0
			indice_SVLAN = 0
			indice_max = 0
			i_redes = 0
			Estatica(0) = "NO TIENE"
			SalidaINTERNET(0) = "NO TIENE"
			PUERTO = "NO TIENE"
			Subint = "NO TIENE"
			SUBINTERFACE = "NO TIENE"
			Estado_Subint = "NO TIENE"
			DESCRIPCION = "NO TIENE"
			BANDWIDTH = "NO TIENE"
			policy_input = "NO TIENE"
			policy_output = "NO TIENE"
			VRF = "NO TIENE"
			ip_WAN = "NO TIENE"
			mascara_WAN = "NO TIENE"
			NEXTHOP = "NO TIENE"
			CVLAN = "NO TIENE"
			RD = "NO TIENE"
			AS_actual = "NO TIENE"
			descripcion_BGP = "NO TIENE"
			BGP = "NO TIENE"
			PEAKFLOW = "NO TIENE"
			PalabrasCLAVE = array("arg", "ARG", "GMT")
					
			'CALVEAR#sh run int Te0/4/0/0.21903653 
			'CALVEAR#sh run router bgp 7303 vrf vpn-ldip-tasa-1083483143 neighbor 186.108.42.38	
			'CALVEAR#sh run | inc 186.108.42.38
			'CALVEAR#sh run router static vrf vpn-ldip-tasa-1083483143 
			
			'file.Write "MARCA = " & MARCA & vbCrLf
			
			'RP/0/8/CPU0:CALVEAR#sh int Te0/4/0/0.21903622
			'Fri Sep  6 11:35:36.971 ARG			
			'Interface not found (TenGigE0/4/0/0.21903622)
			
			str = filedatos.Readline 'aca lee el puerto o subinterface
			PUERTO = str
			file.Write EQUIPO & " " & PUERTO & " Descripcion: sh int " & PUERTO & vbCrLf
			crt.Screen.Send "sh int " & str & chr(13) 'tira el comando para mostrar la descripcion
			crt.screen.WaitForStrings PalabrasCLAVE  'espera la palabra arg
			Salir = crt.screen.WaitForStrings (chr(13), "#", "line protocol", "Description:", "Interface not found")
			do while Salir <> 2						' mientras no aparezca # hacer
					screenrow = crt.screen.CurrentRow 'fila actual
					readline = crt.Screen.Get(screenrow, 1, screenrow, 200)
					readline = rtrim(readline)
					file.Write EQUIPO & " " & PUERTO & " " & readline & vbCrLf
					Salir = crt.screen.WaitForStrings (chr(13), "#", "line protocol", "Description:", "Interface not found") 'esperar por los siguiente caracteres
				Select Case Salir
					Case 2 'encontró # termino de mostrar la descripcion
						Salir = 2
						'MsgBox "aux = 2" 
					Case 3 'encontró line protocol
						columnaActual = crt.screen.CurrentColumn - 15
						screenrow = crt.screen.CurrentRow
						crt.screen.WaitForString chr(13)
						Estado_Subint = crt.Screen.Get(screenrow, 1, screenrow, 200)
						lngPos = Instr(Estado_Subint, " ")
						SUBINTERFACE = Mid(Estado_Subint, 1, lngPos-1)
						Estado_Subint = Mid(Estado_Subint, lngPos+4, Len(Estado_Subint))
						Estado_Subint = trim(Estado_Subint)
						'SUBINTERFACE = crt.Screen.Get(screenrow, 1, screenrow, columnaActual) 'toma la subinterface 
						'MsgBox "Estado_Subint " & Estado_Subint
						'MsgBox "SUBINTERFACE " & SUBINTERFACE
					Case 4 'encontró Description:
						columnaActual = crt.screen.CurrentColumn
						screenrow = crt.screen.CurrentRow	
						crt.screen.WaitForString chr(13)
						DESCRIPCION = crt.Screen.Get(screenrow, columnaActual, screenrow, 200)
						DESCRIPCION = trim(DESCRIPCION)
						'MsgBox "DESCRIPCION " & DESCRIPCION
						'MsgBox "SUBINTERFACE " & SUBINTERFACE
					Case 5 'Interface not found
						fileoutTabladeDatos.Write   "!No encontré la subinterface " & PUERTO & vbCrLf
						file.Write   "!No encontré la subinterface " & PUERTO & vbCrLf
						crt.screen.WaitForString chr(13)
						ESTADO = 3
						funcion_CISCO = ESTADO
						exit function
						'MsgBox "DESCRIPCION " & DESCRIPCION
						'MsgBox "SUBINTERFACE " & SUBINTERFACE
				End Select
			Loop
			'MsgBox "termine de mostrar la descripcion"
			
			
			file.Write EQUIPO & " " & SUBINTERFACE & " Configuracion: sh run int " & PUERTO & vbCrLf
			crt.Screen.Send "sh run int " & PUERTO & chr(13) 'muestro la configuracion del puerto	
			crt.screen.WaitForStrings PalabrasCLAVE 'espera la palabra arg			
			Salir = crt.screen.WaitForStrings (chr(13), "#")
			do while Salir <> 2
				screenrow = crt.screen.CurrentRow
				readline = crt.Screen.Get(screenrow, 1, screenrow, 200)
				readline = rtrim(readline)
				file.Write EQUIPO & " " & SUBINTERFACE & " " & readline & vbCrLf
				Salir = crt.screen.WaitForStrings(chr(13), "#", "bandwidth ", "policy input ", "policy output ", "vrf ", "ipv4 address ", "dot1q vlan ", "interface ", "description ", "encapsulation dot1q", "flow ipv4 monitor", "shutdown")
				Select Case Salir
				Case 3 ' bandwidth
					columnaActual = crt.screen.CurrentColumn
					screenrow = crt.screen.CurrentRow
					crt.screen.WaitForString chr(13)
					BANDWIDTH = crt.Screen.Get(screenrow, columnaActual, screenrow, 200)
					BANDWIDTH = trim(BANDWIDTH)
					'MsgBox BANDWIDTH
				Case 4 ' policy input
					columnaActual = crt.screen.CurrentColumn
					screenrow = crt.screen.CurrentRow
					crt.screen.WaitForString chr(13)
					policy_input = crt.Screen.Get(screenrow, columnaActual, screenrow, 50)
					policy_input = trim(policy_input)
					'MsgBox policy_input	
				Case 5 ' policy output
					columnaActual = crt.screen.CurrentColumn
					screenrow = crt.screen.CurrentRow
					crt.screen.WaitForString chr(13)
					policy_output = crt.Screen.Get(screenrow, columnaActual, screenrow, 50)
					policy_output = trim(policy_output)
					'MsgBox policy_output
				Case 6 ' vrf
					columnaActual = crt.screen.CurrentColumn
					screenrow = crt.screen.CurrentRow
					crt.screen.WaitForString chr(13)
					VRF = crt.Screen.Get(screenrow, columnaActual, screenrow, 100)
					VRF = trim(VRF)
					'MsgBox VRF
				Case 7 ' ipv4 address
					columnaActual = crt.screen.CurrentColumn
					screenrow = crt.screen.CurrentRow
					crt.screen.WaitForString " 255."
					columnaFinal = crt.screen.CurrentColumn - 5
					ip_WAN = crt.Screen.Get(screenrow, columnaActual, screenrow, columnaFinal)
					crt.screen.WaitForString chr(13)
					columnaActual = columnaFinal
					mascara_WAN = crt.Screen.Get(screenrow, columnaActual, screenrow, 200)
					mascara_WAN = trim(mascara_WAN)
					'MsgBox ip_WAN & " mascara: " & mascara_WAN
				Case 8 ' dot1q vlan -> caso C2BELGRANO y esos
					'  dot1q vlan 2009 27
					columnaActual = crt.screen.CurrentColumn
					screenrow = crt.screen.CurrentRow
					crt.screen.WaitForString chr(13)
					CVLAN = crt.Screen.Get(screenrow, columnaActual, screenrow, 200)
					CVLAN = trim(CVLAN)
					lngPos = Instr(CVLAN, " ")
					if lngPos > 0 then
						CVLAN = Mid(CVLAN, lngPos, Len(CVLAN))
						CVLAN = trim(CVLAN)
					end if
					'MsgBox CVLAN
				Case 9 ' Interface
					columnaActual = crt.screen.CurrentColumn
					screenrow = crt.screen.CurrentRow
					crt.screen.WaitForString "."
					columnaFinal = crt.screen.CurrentColumn - 2
					PUERTO = crt.Screen.Get(screenrow, columnaActual, screenrow, columnaFinal)
					crt.screen.WaitForString chr(13)
					columnaActual = columnaFinal + 2
					Subint = crt.Screen.Get(screenrow, columnaActual, screenrow, 50)
				Case 10 ' Description
					columnaActual = crt.screen.CurrentColumn
					screenrow = crt.screen.CurrentRow
					crt.screen.WaitForString chr(13)
					columnaFinal = crt.screen.CurrentColumn
					DESCRIPCION = crt.Screen.Get(screenrow, columnaActual, screenrow, 300)
					DESCRIPCION = trim(DESCRIPCION)
					'MsgBox DESCRIPCION
				Case 11 ' encapsulation dot1q -> para caso MSBNG ejemplo RSC1MB
					' encapsulation dot1q 294 second-dot1q 2
					lngPos = 0
					columnaActual = crt.screen.CurrentColumn
					screenrow = crt.screen.CurrentRow
					crt.screen.WaitForString chr(13)
					CVLAN = crt.Screen.Get(screenrow, columnaActual, screenrow, 200)
					CVLAN = trim(CVLAN)
					lngPos = Instr(CVLAN, "second-dot1q")
					if lngPos > 0 then
						CVLAN = Mid(CVLAN, lngPos+13, Len(CVLAN))
						CVLAN = trim(CVLAN)
					end if
					'MsgBox CVLAN
				Case 12 'flow ipv4 monitor FMMAP sampler FNF_SAMPLER_MAP ingress
					PEAKFLOW = "Si tiene"
				Case 13 'shutdown
					Estado_Subint = "shutdown"
				End Select
			loop
			if Estado_Subint<>"shutdown" then	
				MsgBox "ERROR: la subinterface no esta shuteada"
				ESTADO = 7
				funcion_CISCO = ESTADO
				exit function
			end if
			'MsgBox "termine de leer la configuracion de la subinterface" 
			
			'ahora tengo que ver la configuracion del BGP
			'IP de WAN+1
			'Calculo de Siguiente IP 
			'MsgBox "Esta entrando en el else"
			
			if ip_WAN <> "NO TIENE" then
				NEXTHOP = SiguienteIP(ip_WAN)
				'MsgBox "NEXTHOP " & NEXTHOP
			end if
		
		'CALVEAR#sh run router bgp 7303 vrf vpn-ldip-tasa-1083483143 neighbor 186.108.42.38	
		if ip_WAN <> "NO TIENE" then 'si tiene IP de WAN
			if VRF <> "NO TIENE" then 'y tiene VRF
			file.Write EQUIPO & " " & SUBINTERFACE & " Config Actual BGP: sh run router bgp 7303 vrf " & VRF & " neighbor " & NEXTHOP & vbCrLf
			crt.Screen.Send "sh run router bgp 7303 vrf " & VRF & " neighbor " & NEXTHOP & chr(13) 'muestro la configuracion del BGP
			crt.screen.WaitForStrings PalabrasCLAVE   'espera la palabra arg			
			Salir = crt.screen.WaitForStrings (chr(13), "#", "remote-as ", "description ", "% No such configuration item(s)")
			do while Salir <> 2
				screenrow = crt.screen.CurrentRow
				readline = crt.Screen.Get(screenrow, 1, screenrow, 200)
				readline = rtrim(readline)
				file.Write EQUIPO & " " & SUBINTERFACE & " " & readline & vbCrLf
				Salir = crt.screen.WaitForStrings(chr(13), "#", "remote-as ", "description ", "% No such configuration item(s)")
				Select Case Salir
				Case 3 ' remote-as 
					columnaActual = crt.screen.CurrentColumn 
					screenrow = crt.screen.CurrentRow
					crt.screen.WaitForString chr(13)
					AS_actual = crt.Screen.Get(screenrow, columnaActual, screenrow, 30)
					AS_actual = trim(AS_actual)
					BGP = "Si tiene"
				Case 4 ' description BGP 
					columnaActual = crt.screen.CurrentColumn 
					screenrow = crt.screen.CurrentRow
					crt.screen.WaitForString chr(13)
					descripcion_BGP = crt.Screen.Get(screenrow, columnaActual, screenrow, 100)
					descripcion_BGP = trim(descripcion_BGP)
					BGP = "Si tiene"
				Case 5
					BGP = "NO TIENE"
				End Select
			loop
			else 'si no tiene VRF
				'sh run router bgp 7303 neighbor 181.96.26.250
				file.Write EQUIPO & " " & SUBINTERFACE & " Config Actual BGP: sh run router bgp 7303 " & " neighbor " & NEXTHOP & vbCrLf
				crt.Screen.Send "sh run router bgp 7303 " & " neighbor " & NEXTHOP & chr(13) 'muestro la configuracion del BGP
				crt.screen.WaitForStrings PalabrasCLAVE   'espera la palabra arg			
				Salir = crt.screen.WaitForStrings (chr(13), "#", "remote-as ", "description ", "% No such configuration item(s)")
				do while Salir <> 2
					screenrow = crt.screen.CurrentRow
					readline = crt.Screen.Get(screenrow, 1, screenrow, 200)
					readline = rtrim(readline)
					file.Write EQUIPO & " " & SUBINTERFACE & " " & readline & vbCrLf
					Salir = crt.screen.WaitForStrings(chr(13), "#", "remote-as ", "description ", "% No such configuration item(s)")
					Select Case Salir
					Case 3 ' remote-as 
						columnaActual = crt.screen.CurrentColumn 
						screenrow = crt.screen.CurrentRow
						crt.screen.WaitForString chr(13)
						AS_actual = crt.Screen.Get(screenrow, columnaActual, screenrow, 30)
						AS_actual = trim(AS_actual)
						BGP = "Si tiene"
					Case 4 ' description BGP 
						columnaActual = crt.screen.CurrentColumn 
						screenrow = crt.screen.CurrentRow
						crt.screen.WaitForString chr(13)
						descripcion_BGP = crt.Screen.Get(screenrow, columnaActual, screenrow, 100)
						BGP = "Si tiene"
					Case 5
						BGP = "NO TIENE"
					End Select
				loop
			end if
		end if
		'MsgBox "termine de leer la configuracion de la BGP"
			
		' ahora tengo que ver las estaticas
		if ip_WAN <> "NO TIENE" then
			if VRF <> "NO TIENE" then
				file.Write EQUIPO & " " & SUBINTERFACE & " Config Actual Estaticas: sh run router static vrf " & VRF & " | inc " & PUERTO & "." & Subint & vbCrLf
				crt.Screen.Send "sh run router static vrf " & VRF & " | inc " & PUERTO & "." & Subint & chr(13) 'muestro las estaticas			
				Salir = crt.screen.WaitForStrings (chr(13), "#", NEXTHOP, "arg", NEXTHOP)
				do while Salir <> 2
					'if Salir = 4 or Salir = 3 then crt.screen.WaitForString chr(13)
					screenrow = crt.screen.CurrentRow
					readline = crt.Screen.Get(screenrow, 1, screenrow, 200)
					readline = rtrim(readline)
					file.Write EQUIPO & " " & SUBINTERFACE & " " & readline & vbCrLf
					Salir = crt.screen.WaitForStrings(chr(13), "#", NEXTHOP, "arg")
					Select Case Salir
					Case 3 ' NEXTHOP 
						screenrow = crt.screen.CurrentRow
						crt.screen.WaitForString chr(13)
						'columnaActual = crt.screen.CurrentColumn - 49
						Provisirio = crt.Screen.Get(screenrow, 1, screenrow, 200) 'tomo toda la linea
						lngPos = Instr(Provisirio, "/")		'busco la posicion donde este la barra /
						Estatica(indice) = crt.Screen.Get(screenrow, 1, screenrow, lngPos+2) 'aca tomo la estatica entera con barra y todo
						ipPING(indice) = Mid(Estatica(indice), 1, lngPos-1)		'aca tomo solo la IP
						MASCARA_estatica(indice) = Mid(Estatica(indice), lngPos+1, Len(Estatica(indice))) 'aca tomo solo la Mascara
						Estatica(indice) = trim(Estatica(indice))
						'MsgBox "MASCARA_estatica " & indice & ": " & MASCARA_estatica(indice)
						'MsgBox "Estatica " & indice & " " & Estatica(indice) & " mascara " & MASCARA_estatica(indice)
						'MsgBox "ipPING " & indice & ": " & ipPING(indice)
						indice = indice + 1
					End Select
				loop
				Cant_estaticas = indice
				'MsgBox "Cant_estaticas " & Cant_estaticas
				indice = 0
			else
				file.Write EQUIPO & " " & SUBINTERFACE & " Config Actual Estaticas: sh run router static " & " | inc " & PUERTO & "." & Subint & vbCrLf
				crt.Screen.Send "sh run router static " & " | inc " & PUERTO & "." & Subint & chr(13) 'muestro las estaticas
				'crt.screen.WaitForString "arg"  'espera la palabra arg				
				Salir = crt.screen.WaitForStrings (chr(13), "#", NEXTHOP, "arg", NEXTHOP)
				do while Salir <> 2
					'if Salir = 4 or Salir = 3 then crt.screen.WaitForString chr(13)
					screenrow = crt.screen.CurrentRow
					readline = crt.Screen.Get(screenrow, 1, screenrow, 200)
					readline = rtrim(readline)
					file.Write EQUIPO & " " & SUBINTERFACE & " " & readline & vbCrLf
					Salir = crt.screen.WaitForStrings(chr(13), "#", NEXTHOP, "arg")
					Select Case Salir
					Case 3 ' NEXTHOP 
						screenrow = crt.screen.CurrentRow
						crt.screen.WaitForString chr(13)
						'columnaActual = crt.screen.CurrentColumn - 49
						Provisirio = crt.Screen.Get(screenrow, 1, screenrow, 200)
						lngPos = Instr(Provisirio, "/")
						Estatica(indice) = crt.Screen.Get(screenrow, 1, screenrow, lngPos+2)
						ipPING(indice) = Mid(Estatica(indice), 1, lngPos-1)
						MASCARA_estatica(indice) = Mid(Estatica(indice), lngPos+1, Len(Estatica(indice)))
						Estatica(indice) = trim(Estatica(indice))
						'MsgBox "MASCARA_estatica " & indice & ": " & MASCARA_estatica(indice)
						'MsgBox "Estatica " & indice & " " & Estatica(indice) & " mascara " & MASCARA_estatica(indice)
						'MsgBox "ipPING " & indice & ": " & ipPING(indice)
						indice = indice + 1
					End Select
				loop
				Cant_estaticas = indice
				'MsgBox "Cant_estaticas " & Cant_estaticas
				indice = 0
			end if
			end if
			'MsgBox "termine de ver las estaticas"
			'------------------------------------------------------- estoy aca
			
			' ahora tengo que ver las estaticas que tienen salida a internet
			' sh run | inc 181.10.29.16
			' sh run | inc IP_de_LAN
			do while indice < Cant_estaticas
				if MASCARA_estatica(indice) <> "32" then
					file.Write EQUIPO & " " & SUBINTERFACE & " Estaticas con salida a internet: sh run | inc " & Estatica(indice) & vbCrLf
					crt.Screen.Send "sh run | inc " & Estatica(indice) & chr(13) 
					SalidaINTERNET(indice) = "NO TIENE"
					Salir = crt.screen.WaitForStrings (chr(13), "#", "red-integra", "arg")
					if Salir = 2 then 
						SalidaINTERNET(indice) = "NO TIENE"
						indice = indice + 1
					end if
					do while Salir <> 2
						Select Case Salir
						Case 3 ' red-integra 
							SalidaINTERNET(indice) = Estatica(indice)
							'MsgBox "SalidaINTERNET " & indice & " " & SalidaINTERNET(indice)
							crt.screen.WaitForString chr(13)
							indice = indice + 1
						End Select
						screenrow = crt.screen.CurrentRow
						readline = crt.Screen.Get(screenrow, 1, screenrow, 200)
						readline = rtrim(readline)
						file.Write EQUIPO & " " & SUBINTERFACE & " " & readline & vbCrLf
						Salir = crt.screen.WaitForStrings (chr(13), "#", "red-integra", "arg")
						if Salir = 2 then
							SalidaINTERNET(indice) = "NO TIENE"
							indice = indice + 1
						end if
					loop
					'MsgBox "Cant_estaticas " & Cant_estaticas
				else 
					SalidaINTERNET(indice) = "NO TIENE"
					indice = indice + 1
				end if
			loop
			indice = 0
			'MsgBox "termine de ver las estaticas con salida a internet"
		
			' ahora tengo que imprimir las variables

			file.Write   "---------------------------------------" & vbCrLf
			file.Write   "DATOS Configuracion Actual: " & vbCrLf
			file.Write   "EQUIPO:	" & EQUIPO & vbCrLf
			file.Write   "PORT:	" & PUERTO & vbCrLf
			file.Write   "Subinterface:	" & Subint & vbCrLf
			file.Write   "Subinterface completa:	" & SUBINTERFACE & vbCrLf
			file.Write   "Estado Subinterface:	" & Estado_Subint & vbCrLf
			file.Write   "Descripcion:	" & DESCRIPCION & vbCrLf
			file.Write   "Bandwidth:	" & BANDWIDTH & vbCrLf
			file.Write   "Policy input:	" & policy_input & vbCrLf
			file.Write   "Policy output:	" & policy_output & vbCrLf
			file.Write   "VRF:	" & VRF & vbCrLf
			file.Write   "Ip de WAN:	" & ip_WAN & vbCrLf
			file.Write   "Mascara WAN:	" & mascara_WAN & vbCrLf
			file.Write   "WAN+1:	" & NEXTHOP & vbCrLf
			file.Write   "CVlan:	" & CVLAN & vbCrLf
			file.Write   "RD:	" & RD & vbCrLf
			file.Write   "BGP:	" & BGP & vbCrLf
			file.Write   "AS:	" & AS_actual & vbCrLf
			file.Write   "Peakflow:	" & PEAKFLOW & vbCrLf
			file.Write   "Descripcion BGP:	" & descripcion_BGP & vbCrLf
			do while indice < Cant_estaticas
				file.Write   "Estatica " & indice & " :	" & Estatica(indice) & vbCrLf
				if SalidaINTERNET(indice) <> "NO TIENE" then
					file.Write   "Salida a INTERNET: " & SalidaINTERNET(indice) & vbCrLf
				end if
				indice = indice + 1
			loop
			file.Write   "Resultado ping WAN+1:	" & NEXTHOP & " " & Resultado_pingWAN & vbCrLf
			indice = 0
			do while indice < Cant_estaticas
				file.Write   "Resultado Estatica " & indice & ":	" & Estatica(indice) & " " & Resultado_pingEstatica(indice) & vbCrLf
				indice = indice + 1
			loop
			i_redes = 0
			do while i_redes < Cant_redes
				file.Write Redes_aprendidas(i_redes) & vbCrLf
				i_redes = i_redes + 1
			loop
			
			fileoutTabladeDatos.Write   "---------------------------------------" & vbCrLf
			' borrar la salida a internet
			indice = 0
			do while indice < Cant_estaticas
				if SalidaINTERNET(indice) <> "NO TIENE" then
				'MsgBox SalidaINTERNET(indice)
				fileoutTabladeDatos.Write   "BORRAR SALIDA A INTERNET" & vbCrLf
				fileoutTabladeDatos.Write   "router bgp 7303" & vbCrLf
				fileoutTabladeDatos.Write   " address-family ipv4 unicast" & vbCrLf
				fileoutTabladeDatos.Write   "  no network " & SalidaINTERNET(indice) & " route-policy red-integra" & vbCrLf
				indice = indice + 1
				do while indice < Cant_estaticas
					if SalidaINTERNET(indice) <> "NO TIENE" then
					fileoutTabladeDatos.Write   "  no network " & SalidaINTERNET(indice) & " route-policy red-integra" & vbCrLf
					end if
					indice = indice + 1
				loop
				else
					indice = indice + 1
				end if
			loop
			indice = 0		
			' fin borrar la salida a internet
			
			' borrar las estaticas
			indice = 0
			if Cant_estaticas <> 0 then
				'fileoutTabladeDatos.Write   "BORRAR ESTATICAS" & vbCrLf
				fileoutTabladeDatos.Write   "router static" & vbCrLf
				if VRF <> "NO TIENE" then fileoutTabladeDatos.Write   " vrf " & VRF & vbCrLf 'para el caso de VRF
				fileoutTabladeDatos.Write   "  address-family ipv4 unicast" & vbCrLf
			end if
			do while indice < Cant_estaticas
				'fileoutTabladeDatos.Write   "Estatica: " & Estatica(indice) & vbCrLf
				'no 181.88.28.210/32 TenGigE0/4/0/0.21903653 186.108.42.38
				fileoutTabladeDatos.Write   "   no " & Estatica(indice) & " " & SUBINTERFACE & " " & NEXTHOP & vbCrLf
				indice = indice + 1
			loop
			indice = 0
			' fin borrar estaticas
			
			'fileoutTabladeDatos.Write   "BGP: " & BGP & vbCrLf
			'borrar neighbor BGP
			if BGP <> "NO TIENE" then 
				'fileoutTabladeDatos.Write   "BORRAR neighbor BGP" & vbCrLf
				fileoutTabladeDatos.Write   "router bgp 7303" & vbCrLf
				if VRF <> "NO TIENE" then fileoutTabladeDatos.Write   " vrf " & VRF & vbCrLf
				fileoutTabladeDatos.Write   "  no neighbor " & NEXTHOP & vbCrLf
			end if
			'fin borrar neighbor BGP
			'borrar subinterface
			'fileoutTabladeDatos.Write   "BORRAR SUBINTERFACE" & vbCrLf
			fileoutTabladeDatos.Write   "no int " & SUBINTERFACE & vbCrLf
			'fin borrar subinterface

			if filedatos.AtEndOfStream = True then
				crt.Screen.Send "exit" & chr(13)
				crt.screen.WaitForString("$")
				ESTADO = 7
				funcion_CISCO = ESTADO
				Exit Function
			end if
			
			ESTADO = 3
			
			file.Write "!--------------------------------------------------------" & vbCrLf
			file.Write "!--------------------------------------------------------" & vbCrLf
			if filedatos.AtEndOfStream = True then ESTADO = 7
			funcion_CISCO = ESTADO
End Function


Function SiguienteIP(IP)
	
	lngPos = Instr(IP, ".")
	lngPosAcumulado = lngPos
	PING_PRUEBA = Mid(IP, lngPos+1, Len(IP))
	lngPos = Instr(PING_PRUEBA, ".")
	lngPosAcumulado = lngPosAcumulado + lngPos
	PING_PRUEBA = Mid(PING_PRUEBA, lngPos+1, Len(PING_PRUEBA))
	lngPos = Instr(PING_PRUEBA, ".")	
	lngPosAcumulado = lngPosAcumulado + lngPos
	ultimoCARACTER = Mid(PING_PRUEBA, lngPos+1, Len(PING_PRUEBA))
	PING_PRUEBA = Mid(IP, 1, lngPosAcumulado) & (ultimoCARACTER+1)
	'MsgBox "Original " & IP & vbCrLf & "PING_PRUEBA " & PING_PRUEBA
	
	SiguienteIP = PING_PRUEBA
End Function
