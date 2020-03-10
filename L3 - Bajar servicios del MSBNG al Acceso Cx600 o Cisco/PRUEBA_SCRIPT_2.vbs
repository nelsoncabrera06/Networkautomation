# $language = "VBScript"
# $interface = "1.0"

Sub Main

  ' Send the unix "date" command and wait for the prompt that indicating 
  ' that it completed. In general we want to be in synchronous mode before
  ' doing send/wait operations.
  '
  
  crt.Screen.Synchronous = True
  crt.Screen.Send( "ttelnet CRETIRO" & vbCR )
  crt.Screen.WaitForString( "#" )

  ' When we get here the cursor should be one line below the output
  ' of the 'date' command. Subtract one line and use that value to
  ' read a chunk of text (1 row, 40 characters) from the screen.
  ' 
  fila = crt.screen.CurrentRow ' toma el dato de la fila anterior
  Dim result
  result = crt.Screen.Get(fila, 13, fila, 19 )

  ' Get() reads a fixed size of the screen. So you may need to use
  ' VBScript's regular expression functions or the Split() function to
  ' do some simple parsing if necessary. Just print it out here.
  '
  
  
  'MsgBox result
  
  algodiferente(result)
  
  crt.Screen.Synchronous = False

End Sub

function algodiferente(texto)

	MsgBox "el equipo ingresado es " & texto
	
end function