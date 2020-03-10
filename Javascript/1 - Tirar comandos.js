# $language = "JScript"
# $interface = "1.0"

// Open the file c:\temp\file.txt, read it line by line sending each
// line to the server. Note, to run this script successfully you may need
// to update your script engines to ensure that the filesystemobject runtime
// is available.

var USUARIO, pass, EQUIPO, COMANDO;
USUARIO = "sin datos";
pass = "sin datos";
EQUIPO = "sin datos";
COMANDO = "sin datos";
var ForReading = 1, ForWriting = 2;

function main() 
{
  var fso, f, r;  
  var prueba;
  

  fso = new ActiveXObject("Scripting.FileSystemObject");	
  f = fso.OpenTextFile("inventory.txt", ForReading);


  crt.Screen.Synchronous = true;
	
	 loggearse();
	 
  var str;
  while ( f.AtEndOfStream != true )
  {
		ingresarAlEquipo(f);
		tirarComandos();
		prueba = crt.Screen.get(10,2,10,20);
		crt.Dialog.MessageBox(prueba);
		salirdelequipo();
  }
  
  crt.Dialog.MessageBox("       Fin del script       " + "\015" + "Autor: Nelson Cabrera");
  
};


function salirdelequipo()
{	
	crt.Screen.Send( "exit" + "\015" );
	crt.Screen.WaitForString( "$" );
		
	crt.Screen.Synchronous = true;
	
}



function tirarComandos()
{	
	var fso2, comandos;
	
	fso2 = new ActiveXObject("Scripting.FileSystemObject");	
	comandos = fso2.OpenTextFile("comandos.txt", ForReading);
	
	crt.Screen.Synchronous = true;
	while ( comandos.AtEndOfStream != true )
	{
		COMANDO = comandos.Readline(); //lee una linea
		crt.Screen.Send( COMANDO + "\015" );		// envia el comando
		crt.Screen.WaitForString( "#" );
	}
	crt.Screen.Synchronous = true;
	
	comandos.close();
	
	crt.Dialog.MessageBox("Termine de tirar los comandos");
	
}

function loggearse()
{
	USUARIO = crt.Dialog.Prompt("Ingrese su usuario", "loggearse", "uXXXXXX", 0);
	pass = crt.Dialog.Prompt("Ingrese su password", "Login", "", 1);
	//crt.Dialog.MessageBox("user " + USUARIO);
	//crt.Dialog.MessageBox("pass " + pass);
}

function ingresarAlEquipo(f)
{
	EQUIPO = f.Readline();
	COMANDO = "ttelnet " + EQUIPO + "\015";
    crt.Screen.Send( COMANDO );

    // wait for the prompt before continuing with the next send.
    crt.Screen.WaitForString( "Username:" );
	crt.Screen.Send( USUARIO + "\015" );
	crt.Screen.WaitForString( "Password:" );
	crt.Screen.Send( pass + "\015" );
	crt.Screen.WaitForString( "#" );
	
	//crt.Dialog.MessageBox("hasta aca esta bien ");
}




