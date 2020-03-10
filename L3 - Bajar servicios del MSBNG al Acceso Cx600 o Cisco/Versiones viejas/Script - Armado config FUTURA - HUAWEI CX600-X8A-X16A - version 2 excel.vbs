'# $language = "VBScript"
'# $interface = "1.0"

' este script entra a equipos HUAWEI, entra al puerto
' lee la configuracion del puerto/subinterface y la imprime en un archivo txt
' guarda los datos en variables 
' muestra la VRF
' muestra el BGP
' imprime todas las variables necesarias para armar el servicio

' consideraciones esta armado para CISCO ejemplo CMUNECAS (en los MB son menos caracteres)
' y tome como ejemplo una VPN habria que ver como funciona para Integras

'MsgBox("Hola MUNDO") 

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

ESTADO = 0
Estado_Anterior = 0
linealeida = "sin datos"
TIPOdeDATO = "sin datos"
EQUIPO = "sin datos"
PUERTOANTERIOR = "sin datos"
Subinterface_anterior = "sin datos"
Subinterface_completa_anterior = "sin datos"
Estado_Subinterface = "sin datos"
Descripcion = "sin datos"                                 
Bandwidth = "sin datos"                                                                                                                    
Policy_input = "sin datos"  
Policy_output = "sin datos"
VRF = "sin datos"
IpWAN = "sin datos"
Mascara_WAN = "sin datos"                                                                                                
NEXTHOP = "sin datos"
CVLAN = "sin datos"
RD = "sin datos"
BGP = "sin datos"
BGP_neighbor = "sin datos"
AS_sistemaautonomo = "sin datos"
Descripcion_BGP = "sin datos"
Dim Estatica(50) 
Estatica(0) = "sin datos"
Dim SalidaINTERNET(50) 
SalidaINTERNET(0) = "sin datos"
indice = 0
indice_salidaINTERNET = 0
Cant_estaticas = 0
Cant_IPSalidaInternet = 0
ResultadoNexthop = "sin datos"

ESTADO_FINAL = 9

NombreArchivo_Entrada = "OUTPUT - TabladeDatos.txt"
NombreArchivo_Salida = "OUTPUT - Config FUTURA - excel.txt"


'Sub Main
  Dim waitStrs, comandostring
  Dim row, screenrow, readline, items


  Dim fso, file
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set fsodatos = CreateObject("Scripting.FileSystemObject")
  
  Dim fsodatos, fileinput, str
  Set fileinput = fsodatos.OpenTextFile(NombreArchivo_Entrada, ForReading, False)
  Set fileoutput = fso.OpenTextFile(NombreArchivo_Salida, ForWriting, True)
  
  row = 1 

  RESULTADO = 0
  MARCA = "sin datos"
  Salir = 0
  waitStrs = Array( Chr(10), "...")
  index = 0

	PUERTO = InputBox("Ingrese el puerto futuro","CONFIGURACION FUTURA")
	
	
	do while ESTADO <> ESTADO_FINAL ' mi max estado hasta ahora es 9
		Select Case ESTADO
			Case 0 'recien arranca el script
				ESTADO = funcion_GuardarDatos(TIPOdeDATO, linealeida, fileinput)
				Subinterface_completa = PUERTO & "." & CVLAN
			Case 1 'verificar hay que armado en el equipo
				ESTADO = funcion_verificar(fileoutput)
				ESTADO = 2
			Case 2 'armar los qos-profile
				'ESTADO = funcion_qos-profile(fileoutput) ' esta funcion todavia no esta creada
				ESTADO = 3
			Case 3 'armar la VRF
				ESTADO = funcion_configVRF(fileoutput)
				ESTADO = 4
			Case 4 'armar el BGP
				ESTADO = funcion_configBGP(fileoutput)
				ESTADO = 5
			Case 5 'armar la subinterface
				ESTADO = funcion_configSUBINTERFACE(fileoutput)
				ESTADO = 6
			Case 6 'configurar las estaticas
				ESTADO = funcion_estaticas(fileoutput)
				ESTADO = 7
			Case 7 'configurar las estaticas con salida a internet
				ESTADO = funcion_salidaINTENET(fileoutput)
				ESTADO = 8
			Case 8 'ahora deberia repetir el proceso para el siguiente caso
				ESTADO = 0	'vuelvo a recopilar los datos
				'MsgBox "ESTADO " & ESTADO
		End Select
	loop
	
  fileoutput.Close
  fileinput.Close
  'Set objTextStream = Nothing
  'Set objFSO = Nothing

	
	MsgBox("Fin del script") 
	Set g_shell = CreateObject("WScript.Shell")
	g_shell.Run chr(34) & NombreArchivo_Salida & chr(34)

'End Sub                                                                                                                                                                                                                                                                                             

Function funcion_verificar(fileoutput) 'Modelo de ejemplo HUAWEI CX600-X8 - SPE2MU
	
	fileoutput.Write "verificar	" & "#  Verificaciones Previas " & vbCrLf
	fileoutput.Write "verificar	" & "disp curr configuration qos-profile " & policy_input & vbCrLf
	fileoutput.Write "verificar	" & "disp curr configuration qos-profile " & policy_output & vbCrLf
	if VRF <> "NO TIENE" then fileoutput.Write "verificar	" & "disp curr conf vpn-instance " & VRF & vbCrLf
	if BGP <> "NO TIENE" then fileoutput.Write "verificar	" & "disp curr conf bgp | begin " & VRF & vbCrLf
	fileoutput.Write "verificar	" & "disp int desc | inc " & PUERTO & vbCrLf
	fileoutput.Write "verificar	" & "disp int desc | inc " & Subinterface_completa & vbCrLf
	fileoutput.Write "verificar	" & "disp curr | inc " & NEXTHOP & vbCrLf
	fileoutput.Write "verificar	" & "#  " & vbCrLf
	
	funcion_verificar = ESTADO
End Function

Function funcion_configVRF(fileoutput) 'Modelo de ejemplo HUAWEI CX600-X8A-16A - LOP4MU
	
	'MsgBox "empezando a escribir VRF! ESTADO " & ESTADO
	
	if VRF <> "NO TIENE" then  
		fileoutput.Write "VRF	" & "#  VRF a configurar en Cx600" & vbCrLf
		fileoutput.Write "VRF	" &"#" & vbCrLf
		fileoutput.Write "VRF	" &"ip vpn-instance " & VRF & vbCrLf
		fileoutput.Write "VRF	" &"ipv4-family " & vbCrLf
		fileoutput.Write "VRF	" &"route-distinguisher " & RD & vbCrLf
		fileoutput.Write "VRF	" &"export route-policy loopback-vpn " & vbCrLf
		fileoutput.Write "VRF	" &"vpn-target " & RD & " export-extcommunity" & vbCrLf 	
		fileoutput.Write "VRF	" &"vpn-target " & RD & " 7303:2 import-extcommunity" & vbCrLf
		fileoutput.Write "VRF	" &"vpn-target 7303:2 import-extcommunity" & vbCrLf
		fileoutput.Write "VRF	" &"#" & vbCrLf
	else
		fileoutput.Write "VRF	" & "#  NO TIENE VRF" & vbCrLf
	end if
	
	funcion_configVRF = ESTADO
End Function

Function funcion_configBGP(fileoutput) 'Modelo de ejemplo HUAWEI CX600-X8A-16A - LOP4MU
	
	'MsgBox "empezando a escribir BGP! ESTADO " & ESTADO
	if BGP <> "NO TIENE" then
		fileoutput.Write "BGP	" & "#  BGP a configurar en Cx600" & vbCrLf
		fileoutput.Write "BGP	" & "#" & vbCrLf
		fileoutput.Write "BGP	" &"bgp 7303 " & vbCrLf
		if VRF <> "NO TIENE" then ' si tiene VRF

			fileoutput.Write "BGP	" & "ipv4-family vpn-instance " & VRF &  vbCrLf
			fileoutput.Write "BGP	" & "import-route direct" & vbCrLf
			fileoutput.Write "BGP	" & "import-route static" & vbCrLf
			if BGP_neighbor <> "NO TIENE" then
				fileoutput.Write "BGP	" & "peer " & NEXTHOP & " as-number " & AS_sistemaautonomo & vbCrLf                                                     
				fileoutput.Write "BGP	" & "peer " & NEXTHOP & " description " & Descripcion_BGP & vbCrLf 	
				fileoutput.Write "BGP	" & "peer " & NEXTHOP & " route-limit 100 90 idle-timeout 10" & vbCrLf
				fileoutput.Write "BGP	" & "peer " & NEXTHOP & " advertise-community " & vbCrLf 
				fileoutput.Write "BGP	" & "peer " & NEXTHOP & " advertise-ext-community " & vbCrLf 		
				fileoutput.Write "BGP	" &"#" & vbCrLf
			end if
		else 'tiene BGP pero no tiene VRF
			if BGP_neighbor <> "NO TIENE" then
				fileoutput.Write "BGP	" & "peer " & NEXTHOP & " as-number " & AS_sistemaautonomo & vbCrLf                                                     
				fileoutput.Write "BGP	" & "peer " & NEXTHOP & " description " & Descripcion_BGP & vbCrLf 
				fileoutput.Write "BGP	" & "ipv4-family unicast" & vbCrLf 
				fileoutput.Write "BGP	" & " peer " & NEXTHOP & " enable" & vbCrLf 
				fileoutput.Write "BGP	" & " peer " & NEXTHOP & " route-policy distribute-list_120 export" & vbCrLf
				fileoutput.Write "BGP	" & " peer " & NEXTHOP & " route-limit 100 90 idle-timeout 10" & vbCrLf
				fileoutput.Write "BGP	" & " peer " & NEXTHOP & " default-route-advertise" & vbCrLf
				fileoutput.Write "BGP	" & " peer " & NEXTHOP & " route-policy RED-INTEGRA import" & vbCrLf
				fileoutput.Write "BGP	" & "#" & vbCrLf
			end if
		end if
	else 'si no tiene BGP
		fileoutput.Write "BGP	" &"#  NO TIENE BGP" & vbCrLf
	end if
	

	
	
	
	funcion_configBGP = ESTADO
End Function

Function funcion_configSUBINTERFACE(fileoutput) 'Modelo de ejemplo HUAWEI CX600-X8A-16A - LOP4MU
	
	'MsgBox "empezando a escribir! ESTADO " & ESTADO
	'MsgBox "interface " & Subinterface_completa & "description " & Descripcion 
	
	fileoutput.Write "Subint	" & "#  Configuracion Subinterface FUTURA" & vbCrLf
	fileoutput.Write "Subint	" & "#" & vbCrLf
	fileoutput.Write "Subint	" & "interface " & Subinterface_completa & vbCrLf
	fileoutput.Write "Subint	" & "shutdown " & vbCrLf
	fileoutput.Write "Subint	" & "bandwidth " & Bandwidth & vbCrLf
	fileoutput.Write "Subint	" & "description " & Descripcion & vbCrLf
	if VRF <> "NO TIENE" then fileoutput.Write "Subint	" & "ip binding vpn-instance " & VRF & vbCrLf  
    fileoutput.Write "Subint	" &  "ip address " & IpWAN & " " & Mascara_WAN & vbCrLf
	fileoutput.Write "Subint	" & "statistic enable" & vbCrLf 
	fileoutput.Write "Subint	" & "encapsulation dot1q-termination" & vbCrLf
	fileoutput.Write "Subint	" & "dot1q termination vid " & CVLAN & vbCrLf
	fileoutput.Write "Subint	" & "arp broadcast enable " & vbCrLf
	fileoutput.Write "Subint	" & "trust upstream principal" & vbCrLf
	fileoutput.Write "Subint	" & "trust 8021p " & vbCrLf
	fileoutput.Write "Subint	" & "qos-profile " & Policy_input & " inbound vlan " & CVLAN & " identifier none " & vbCrLf   
    fileoutput.Write "Subint	" & "qos-profile " & Policy_output & " outbound vlan " & CVLAN & " identifier none "  & vbCrLf                                                                                                        
	fileoutput.Write "Subint	" & "#" & vbCrLf
	'fileoutput.Write "Subint	" & "#---------------------------------------" & vbCrLf
	
	
	funcion_configSUBINTERFACE = ESTADO
End Function
                                                                       

Function funcion_estaticas(fileoutput) 
	
	'todavía me falta ver como hacer con las estaticas con salida a internet LAN /29 o /24 esas
	'Estatica(0) = "sin datos"
	indice = 0
	ipEstatica = "sin datos"
	MASCARA_estatica = "sin datos"
	
	if Estatica(indice) <> "sin datos" then
		fileoutput.Write "Estaticas	" & "#  Configuracion estaticas " & vbCrLf
		fileoutput.Write "Estaticas	" & "#" & vbCrLf
		
		Do While indice < Cant_estaticas
			'calcular la ip y su mascara
			lngPos = Instr(Estatica(indice), "/")
			ipEstatica = Mid(Estatica(indice), 1, lngPos-1)
			MASCARA_estatica = Mid(Estatica(indice), lngPos+1, Len(Estatica(indice)))
			
			'Msgbox MASCARA_estatica
			Select Case MASCARA_estatica
				Case "32" 
					MASCARA_estatica = "255.255.255.255"
				Case "30" 
					MASCARA_estatica = "255.255.255.252"
				Case "29" 
					MASCARA_estatica = "255.255.255.248"
				Case "28"
					MASCARA_estatica = "255.255.255.240"
				Case "27"
					MASCARA_estatica = "255.255.255.224"
				Case "26"
					MASCARA_estatica = "255.255.255.192"
				Case "24"
					MASCARA_estatica = "255.255.255.0"
				Case Else
					MASCARA_estatica = "ERROR_MASCARA_" & MASCARA_estatica
			End Select
		
			if VRF <> "NO TIENE" then
				fileoutput.Write "Estaticas	" & "ip route-static vpn-instance " & VRF & " " & ipEstatica & " " & MASCARA_estatica & " " & Subinterface_completa & " " & NEXTHOP & vbCrLf 
				'faltaria el description BGP
				'ip route-static vpn-instance  bna-vpn    181.15.23.230 255.255.255.255 88.32.20.2 description  305801-BNA
			else	
				fileoutput.Write "Estaticas	" & "ip route-static " & ipEstatica & " " & MASCARA_estatica & " " & Subinterface_completa & " " & NEXTHOP & vbCrLf
				'ip route-static   186.153.162.168 255.255.255.248 GigabitEthernet1/0/0.3852 181.15.1.222
			end if
			indice = indice + 1
		loop
		
	else
		fileoutput.Write "Estaticas	" & "#  no tiene estaticas " & vbCrLf
	end if
    'fileoutput.Write "Estaticas	" & "#---------------------------------------" & vbCrLf
	
	'ip route-static 181.96.224.50 255.255.255.255 g2/2/4.2 181.96.17.54
	'tengo que reiniciar las estaticas
	Estatica(indice) = "sin datos"
	Estatica(0) = "sin datos"
	
	funcion_estaticas = ESTADO
End Function

Function funcion_salidaINTENET(fileoutput) 
	indice_salidaINTERNET = 0
	ipEstatica = "sin datos"
	MASCARA_estatica = "sin datos"
	PrimeraVEZ = 0
	
	Do While indice_salidaINTERNET < Cant_IPSalidaInternet
		if SalidaINTERNET(indice_salidaINTERNET) <> "NO TIENE" then
		
			if PrimeraVEZ = 0 then
				fileoutput.Write "Salida internet	" & "#  Configuracion Salida internet " & vbCrLf
				fileoutput.Write "Salida internet	" & "#" & vbCrLf
				fileoutput.Write "Salida internet	" & "bgp 7303" & vbCrLf 
				fileoutput.Write "Salida internet	" & " ipv4-family unicast" & vbCrLf
				fileoutput.Write "Salida internet	" & "  undo synchronization" & vbCrLf
			else
				PrimeraVEZ = 1
			end if
		
			'calcular la ip y su mascara
			lngPos = Instr(SalidaINTERNET(indice_salidaINTERNET), "/")
			if lngPos <> 0 then
				ipEstatica = Mid(SalidaINTERNET(indice_salidaINTERNET), 1, lngPos-1)
				MASCARA_estatica = Mid(SalidaINTERNET(indice_salidaINTERNET), lngPos+1, Len(SalidaINTERNET(indice_salidaINTERNET)))
			else
				ipEstatica = "Error"
				MASCARA_estatica = "Error"
			end if
			
			'Msgbox MASCARA_estatica
			Select Case MASCARA_estatica
				Case "32" 
					MASCARA_estatica = "255.255.255.255"
				Case "29" 
					MASCARA_estatica = "255.255.255.248"
				Case "30" 
					MASCARA_estatica = "255.255.255.252"
				Case "24"
					MASCARA_estatica = "255.255.255.0"
				Case Else
					MASCARA_estatica = "Error_Mascara_no_encontrada" & MASCARA_estatica
			End Select
			
			'network   181.13.43.32 255.255.255.248 route-policy RED-INTEGRA
			'network   ESTATICA MASCARA_estatica route-policy RED-INTEGRA
			fileoutput.Write "Salida internet	" & "network " & ipEstatica & " " & MASCARA_estatica & " route-policy RED-INTEGRA" & vbCrLf 
		end if
		indice_salidaINTERNET = indice_salidaINTERNET + 1
	loop
	
	if PrimeraVEZ = 0 then
		fileoutput.Write "Salida internet	" & "#  no tiene Salida a internet " & vbCrLf
	end if
    fileoutput.Write "Salida internet	" & "#---------------------------------------" & vbCrLf
	
	SalidaINTERNET(indice_salidaINTERNET) = "sin datos"
	SalidaINTERNET(0) = "sin datos"

	funcion_salidaINTENET = ESTADO
End Function

Function funcion_GuardarDatos(TIPOdeDATO, linealeida, fileinput)
	indice = 0
	indice_salidaINTERNET = 0
	
	Do While TIPOdeDATO <> "Resultado ping WAN+1" and ESTADO = 0 
	if fileinput.AtEndOfStream <> True then
		linealeida = fileinput.Readline 'leo la linea
		lngPos = Instr(linealeida, ":")
		if lngPos>0 then 
			TIPOdeDATO = Mid(linealeida, 1, lngPos-1)
		else
			TIPOdeDATO = "---------------------------------------"
		end if
		TIPOdeDATO = trim(TIPOdeDATO)
		DATO = Mid(linealeida, lngPos+1, Len(linealeida))
		DATO = trim(DATO)
		Select Case TIPOdeDATO
		    Case "---------------------------------------" 'inicio no hago nada
			Case "DATOS Configuracion Actual" 'inicio no hago nada
			Case "EQUIPO" 'guardo el equipo     'EQUIPO:	RSC1MB
				EQUIPO = DATO
			Case "PORT" 'PORT:	TenGigE0/1/0/20
				PUERTOANTERIOR = DATO
			Case "Subinterface" 'Subinterface:	13683648  
				Subinterface_anterior = DATO
			Case "Subinterface completa" 'Subinterface completa:	Te0/1/0/20.13683648
				Subinterface_completa_anterior = DATO
			Case "Estado Subinterface" 'Estado Subinterface:	up          up     
				Estado_Subinterface = DATO
			Case "Descripcion" 'Descripcion:	VPN Branch Office - Cliente - Ref:263534 Lin:2721088 Acc:CHL21Dplaca14puerto0  
				Descripcion = DATO
			Case "Bandwidth" 'Bandwidth:	1024     
				Bandwidth = DATO
			Case "Policy input" 'Policy input:	VPN-1M 
				Policy_input = DATO
			Case "Policy output" 'Policy output:	FUERA-DE-PRODUCTO-1M 
				Policy_output = DATO
			Case "VRF" 'VRF:	kellerhoff-vpn 
				VRF = DATO
			Case "Ip de WAN" 'Ip de WAN:	192.168.81.233 
				IpWAN = DATO
			Case "Mascara WAN" 'Mascara WAN:	 255.255.255.252   
				Mascara_WAN = DATO
			Case "WAN+1"
				NEXTHOP = DATO
			Case "CVLAN" 'CVLAN:	
				CVLAN = DATO
				'MsgBox "CVLAN " & CVLAN
			Case "RD" 'RD:	7303:1164 
				RD = DATO
			Case "BGP"  
				BGP = DATO
			Case "BGP neighbor"  
				BGP_neighbor = DATO
			Case "AS" 'AS:	
				AS_sistemaautonomo = DATO
			Case "Descripcion BGP" 'Descripcion BGP:	
				Descripcion_BGP = DATO
			Case "Estatica" 'Estatica: 181.96.224.50/32
				Estatica(indice) = DATO
				indice = indice + 1
			Case "Salida a INTERNET" 
			    'Salida a INTERNET: 181.10.172.8/29
				'Salida a INTERNET: NO TIENE
				SalidaINTERNET(indice_salidaINTERNET) = DATO
				indice_salidaINTERNET = indice_salidaINTERNET + 1
			Case "Resultado ping WAN+1" 'Resultado ping WAN+1:	192.168.81.234 Success rate is 100 percent (5/5)  
				ResultadoNexthop = DATO
				Cant_estaticas = indice
				Cant_IPSalidaInternet = indice_salidaINTERNET
				ESTADO = 1 	'leer mis datos
				TIPOdeDATO = "ninguno"
		End Select
	else ESTADO = ESTADO_FINAL
	end if
		'MsgBox "fin de select case ESTADO:  " & ESTADO & " ESTADO anterior:  " & Estado_Anterior
	loop
		
		funcion_GuardarDatos = ESTADO
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