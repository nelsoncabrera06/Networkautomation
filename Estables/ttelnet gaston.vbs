# $language = "VBScript"
# $interface = "1.0"

Const BUTTON_YESNO = 4		' Yes and No buttons
Const IDYES = 6		' Yes button clicked
Const IDNO = 7		' No button clicked

Sub main
	  
  Dim result, verific, cap, prueba

  
    crt.Screen.Synchronous = True
	crt.Screen.Send Chr(13)
  
   crt.Screen.Send "" & chr(13)
  result = crt.Dialog.Prompt("Ingrese el nombre del equipo")

  verific = crt.Dialog.MessageBox("Desea ingresar a " &result , cap, BUTTON_YESNO)
 


	if verific = IDYES then 

		crt.Screen.Send "ttelnet " & result & chr(13)
		crt.screen.WaitForString "Trying"
		'crt.Screen.Send "" 
		condicional = crt.Screen.WaitForStrings ("Username:","Connection refused","Name or service not known")
		Select Case condicional
			Case 1 
				'crt.Screen.WaitForString "Username:"
				crt.Screen.Send "u564508" & chr(13)
				crt.Screen.WaitForString "Password:"
				crt.Screen.Send "Golf0420" & chr(13)
				Tab = crt.GetTabCount()
				crt.window.caption = result
			Case 2
				crt.Dialog.MessageBox("No se pudo ingresar al equipo reintente")
				exit sub
			Case 3
				crt.Dialog.MessageBox("le mandaste fruta - No se pudo ingresar al equipo reintente")
				exit sub
		End Select
	else 
		crt.Dialog.MessageBox("VUELVA A INTENTAR")
	end if


 

End Sub



