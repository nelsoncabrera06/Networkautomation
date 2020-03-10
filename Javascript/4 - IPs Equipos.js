# $language = "JScript"
# $interface = "1.0"

// Open the file c:\temp\file.txt, read it line by line sending each
// line to the server. Note, to run this script successfully you may need
// to update your script engines to ensure that the filesystemobject runtime
// is available.

var USUARIO, pass;
var EQUIPO, MARCA, COMANDO;
USUARIO = "sin datos";
pass = "sin datos";
EQUIPO = "sin datos";
COMANDO = "sin datos";
var ForReading = 1, ForWriting = 2;
var filas = 0, columnas = 0, indice = 0;
var tabla;
var NombreArchivo_Salida = "output.txt"
//creo el archivo de salida
var FileOpener = new ActiveXObject ( "Scripting.FileSystemObject"); 
//var FileOUT = FileOpener.OpenTextFile ("output.csv", 2, true);  //ForWriting 
var FileOUT = FileOpener.OpenTextFile (NombreArchivo_Salida, 2, true);  //ForWriting 

function main() 
{
  var fso, f, r;  
	var IP;
  fso = new ActiveXObject("Scripting.FileSystemObject");	
  f = fso.OpenTextFile("inventory.txt", ForReading);
	


	
	//armarFILA("EQUIPO", "IP");
	//tablaCSV();
	WriteToFile("EQUIPO", "IP");
	
	crt.Screen.Synchronous = true;
	
	loggearse();
	 
	var str;
	while ( f.AtEndOfStream != true )
	{
		ingresarAlEquipo(f);
		//tirarComandos();
		
		
		salirdelequipo();
	}
	
	FileOUT.Close (); 
	
  crt.Dialog.MessageBox("       Fin del script       " + "\015" + "Autor: Nelson Cabrera");
  
	var shell = new ActiveXObject("WScript.Shell");
	//shell.Run("\"C:\\Program Files\\Internet Explorer\\IExplore.exe\" http://www.vandyke.com");
	shell.Run(NombreArchivo_Salida);
  
  
}


function salirdelequipo()
{	
	if ( MARCA == "HUAWEI" )
	{
		crt.Screen.Send( "quit" + "\015" );
	}else crt.Screen.Send( "exit" + "\015" );
	
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
	var fila, readline, esperado;
	
	EQUIPO = f.Readline();
	COMANDO = "ttelnet " + EQUIPO + "\015";
    crt.Screen.Send( COMANDO );

	crt.Screen.WaitForString( "Trying" );
	fila = crt.screen.CurrentRow; //fila actual
	crt.Screen.WaitForString( "..." );
	readline = crt.Screen.Get(fila, 1, fila, 200);
	
	
	var re = /Trying\s(\d+\.\d+\.\d+\.\d+)...\s+/g;
	var IP = readline.replace(re, "$1");
	//crt.Dialog.MessageBox(DATO);
	
	WriteToFile(EQUIPO, IP);
	
	
    // wait for the prompt before continuing with the next send.
    crt.Screen.WaitForString( "Username:" );
	crt.Screen.Send( USUARIO + "\015" );
	crt.Screen.WaitForString( "Password:" );
	crt.Screen.Send( pass + "\015" );
	//crt.Screen.WaitForString( "#" );
	esperado = crt.Screen.WaitForStrings("#", ">");
	if ( esperado == 2 ){
		MARCA = "HUAWEI";
	}else MARCA = "CISCO";
	
	//crt.Dialog.MessageBox("hasta aca esta bien ");
}


function armarFILA(EQUIPO, IP)
{
	tabla[indice] = [EQUIPO, IP];
	indice++;	
}

function tablaCSV()
{
	var A = [['n','sqrt(n)']];  // initialize array of rows with header row as 1st item
		  
	for(var j=1;j<10;++j){ 
		A.push([j, Math.sqrt(j)]) 
	}
	 
	var csvRows = [];
	for(var i=0,l=A.length; i<l; ++i){
			csvRows.push(A[i].join(','));   // unquoted CSV row
	}
	var csvString = csvRows.join('\n');
	 
	var a = document.createElement('a');
	a.href     = 'data:attachment/csv,' + csvString;
	a.target   ='_blank';
	a.download = 'myFile.csv,' + encodeURIComponent(csvString); ;
	a.innerHTML = "Click me to download the file.";
	document.body.appendChild(a);
}

function WriteToFile(texto1, texto2) 
{ 
	FileOUT.WriteLine(texto1 + "," + texto2); 
}

