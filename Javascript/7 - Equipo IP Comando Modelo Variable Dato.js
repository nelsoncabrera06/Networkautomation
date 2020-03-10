# $language = "JScript"
# $interface = "1.0"

// Open the file c:\temp\file.txt, read it line by line sending each
// line to the server. Note, to run this script successfully you may need
// to update your script engines to ensure that the filesystemobject runtime
// is available.

var USUARIO, pass;
var EQUIPO, MARCA, MODELO, IP, COMANDO;
USUARIO = "sin datos";
pass = "sin datos";
EQUIPO = "sin datos";
COMANDO = "sin datos";
MODELO = "sin datos";
var ForReading = 1, ForWriting = 2, ForAppending = 8;
var filas = 0, columnas = 0, indice = 0;
var tabla;
var NombreArchivo_Salida = "output.txt"
//creo el archivo de salida
var FileOpener = new ActiveXObject ( "Scripting.FileSystemObject"); 
//var FileOUT = FileOpener.OpenTextFile ("output.csv", 2, true);  //ForWriting 
var FileOUT = FileOpener.OpenTextFile (NombreArchivo_Salida, ForWriting, true);  //ForWriting 

function main() 
{
	var fso, f, r;  
	var IP;
	fso = new ActiveXObject("Scripting.FileSystemObject");	
	f = fso.OpenTextFile("inventory.txt", ForReading);

	//armarFILA("EQUIPO", "IP");
	//tablaCSV();
	//WriteToFile("EQUIPO", "IP", "MARCA", "MODELO");
	EscribirEnArchivo("EQUIPO", "IP", "COMANDO", "VARIABLE", "DATO");
	
	crt.Screen.Synchronous = true;
	
	loggearse();
	 
	var str;
	while ( f.AtEndOfStream != true )
	{
		ingresarAlEquipo(f);
		tirarComandos();
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
	var fso2, comandos, fila, salir, obtenido, DATO, regular, linealeida, lngPos, variable;
	
	fso2 = new ActiveXObject("Scripting.FileSystemObject");
	fsoRegEx = new ActiveXObject("Scripting.FileSystemObject");
	
	switch(MARCA) {
	  case "HUAWEI":
		comandos = fso2.OpenTextFile("comandos HUAWEI.txt", ForReading);
		regexFile = fsoRegEx.OpenTextFile("regex HUAWEI.txt", ForReading);
		break;
	  case "CISCO":
		comandos = fso2.OpenTextFile("comandos CISCO.txt", ForReading);
		regexFile = fsoRegEx.OpenTextFile("regex CISCO.txt", ForReading);
		break;
	  default:
		// code block
	}
	

	
	crt.Screen.Synchronous = true;
	while ( comandos.AtEndOfStream != true )
	{
		COMANDO = comandos.Readline(); //lee una linea
		crt.Screen.Send( COMANDO + "\015" );		// envia el comando
		//crt.Screen.WaitForString( "#" );
		
		//aca PARSEO
		linealeida = regexFile.Readline();
		lngPos = linealeida.indexOf(",");
		if (lngPos!=-1) { //  encontré una coma ,
		  	variable = linealeida.substring(0, lngPos);
			regular =  linealeida.substring(lngPos+1, linealeida.length-1); //lee la linea de la expresion regular
			//crt.Dialog.MessageBox(regular);
		}else crt.Dialog.MessageBox("ERROR");	
		re = new RegExp(regular);		//crea una regExp a partir de ese string leido
		fila = crt.screen.CurrentRow; //fila actual
		readline = crt.Screen.Get(fila, 1, fila, 200);
		
		switch(MARCA) {
		  case "HUAWEI":
						
			obtenido = readline.match(re);
			if (obtenido != null){	
				crt.Dialog.MessageBox(obtenido);
			}else{
				while ( salir != 2 ){ // si salir = 2 llegue al final del comando
					salir = crt.Screen.WaitForStrings("\015", ">");
					fila = crt.screen.CurrentRow; //fila actual
					readline = crt.Screen.Get(fila, 1, fila, 200);
					obtenido = readline.match(re);
					if (obtenido != null) {
						//crt.Dialog.MessageBox(obtenido[1]); // si coindice con la RegExp muestra la variable
						DATO = obtenido[1]; // DATO es lo que queria encontrar
					}
				}
				salir = 0;
			}
			
			break;
		  case "CISCO":
						
			obtenido = readline.match(re);
			if (obtenido != null){	
				crt.Dialog.MessageBox(obtenido);
			}else{
				while ( salir != 2 ){ // si salir = 2 llegue al final del comando
					salir = crt.Screen.WaitForStrings("\015", "#");
					fila = crt.screen.CurrentRow; //fila actual
					readline = crt.Screen.Get(fila, 1, fila, 200);
					obtenido = readline.match(re);
					if (obtenido != null) {
						//crt.Dialog.MessageBox(obtenido[1]); // si coindice con la RegExp muestra la variable
						DATO = obtenido[1]; // DATO es lo que queria encontrar
					}
				}
				salir = 0;
			}
		  
			/*
			var re = /(ASR\s\d+)\s\d+[\s\S+]+/g
			crt.Screen.WaitForString( "PEM" );
			fila = crt.screen.CurrentRow; //fila actual
			readline = crt.Screen.Get(fila, 1, fila, 200);
			MODELO = readline.replace(re, "$1");*/
			break;
		  default:
			// code block
		}
		
		//crt.Screen.WaitForStrings("#", ">");
		//WriteToFile(EQUIPO, IP, MARCA, DATO);
		EscribirEnArchivo(EQUIPO, IP, COMANDO, variable, DATO);
		
		DATO = "sin datos";
		//obtenido = [];
	}
	crt.Screen.Synchronous = true;
	
	comandos.close();
	
	//crt.Dialog.MessageBox("Termine de tirar los comandos");
	//crt.Dialog.MessageBox(MODELO);
	//WriteToFile(EQUIPO, IP, MARCA, MODELO);
	
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
	
	//aca PARSEO
	var re = /Trying\s(\d+\.\d+\.\d+\.\d+)...\s+/g;
	IP = readline.replace(re, "$1");
	//crt.Dialog.MessageBox(DATO);
	
	
	//[A-Z]\w+\s\d+\s\w+\s\w+\s\w+:\s(\d+)\s\w+/\w+,\s\d+\s\w+/\w+ // trafico
	
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
	
	//WriteToFile(EQUIPO, IP, MARCA, "MODELO");
	
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

function WriteToFile(texto1, texto2, tex3, t4) 
{ 	
	//el orden correcto sería EQUIPO IP Marca Modelo
	FileOUT.WriteLine(texto1 + "," + texto2 +","+ tex3 +","+ t4); 
}

function EscribirEnArchivo(t1, t2, t3, t4, t5) 
{ 	
	//el orden correcto sería EQUIPO IP Marca Modelo
	//FileOUT.WriteLine(t1 +"	"+ t2 +"	"+ t3 +"	"+ t4 +"	"+ t5); 
	FileOUT.WriteLine(t1 +","+ t2 +","+ t3 +","+ t4 +","+ t5); 
}

