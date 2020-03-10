/*
  falta ver bien la parte de las vrf
  que tome las RD import
  y las RD export
*/
var contenido;
var lines;
var linealeida;
var tabla_con_datos;
var cant_filas;
var cant_columnas;
var TRANSPORTE = "NO TIENE";
var SVLAN = "NO TIENE";
var MARCA = "NO TIENE";
var MODELO = "NO TIENE";
var PUERTO = "sin datos";
var TIPOdeDATO = "sin datos";
var ESTADO = 0;
var EquipoFUTURO;
var Subinterface_completa;
var tabladedatos;
var ULTIMAlinealeida = 0;
var Descripcion = "sin datos";
var IpWAN = "sin datos";
var Bandwidth = "sin datos";
var Policy_input = "sin datos";
var Policy_output = "sin datos";
var VRF = "sin datos";
var Mascara_WAN = "sin datos";
var NEXTHOP = "sin datos";
var CVLAN = "sin datos";
var SVLANleida = "sin datos"; //SVLAN
var RD = "sin datos";
var RD_import;
var RD_export;
var BGP = "sin datos";
var BGP_neighbor = "sin datos";
var AS_sistemaautonomo = "sin datos";
var Descripcion_BGP = "sin datos";
var Estatica = new Array(); //creo el array de estaticas
Estatica[0] = "sin datos";
var SalidaINTERNET = new Array();
SalidaINTERNET[0] = "sin datos";
var indice = 0;
var indice_salidaINTERNET = 0;
var Cant_estaticas;
var Cant_IPSalidaInternet = 0;
var PRIMER_SERVICIO = 0;
var CANTIDAD_DE_SERVICIOS = new Array();
var Servicio_index = 0;
var TABLAGLOBAL  = new Array(2);
var fila_tabla = 0;
TABLAGLOBAL[0]  = new Array(); // instancia
TABLAGLOBAL[1]  = new Array(); // Configuración FUTURA


class clase_ESTATICAS {
  constructor ( IP , SalidaINTERNET ){
    this.IP = IP;
    //this.MASCARA = MASCARA;
    this.SalidaINTERNET = SalidaINTERNET;
  }
}

//lo siguiente hace que el gif de cargando no lo muestre al principio
document.getElementById('loading').style.display = 'none';
document.getElementById('mitabla').style.display = 'none';
document.getElementById('file-input').addEventListener('change', leerArchivo, false);
document.getElementById('file-input').addEventListener('change', infoARCHIVO, false);



function subirarchivo(){
  var x = document.getElementById("myFile");
  var txt = "";
  if ('files' in x) {
    if (x.files.length == 0) {
      txt = "Select one or more files.";
    } else {
      for (var i = 0; i < x.files.length; i++) {
        txt += "<br><strong>" + (i+1) + ". file</strong><br>";
        var file = x.files[i];
        if ('name' in file) {
          txt += "name: " + file.name + "<br>";
        }
        if ('size' in file) {
          txt += "size: " + file.size + " bytes <br>";
        }
      }
    }
  }
  document.getElementById ("respuestaMyFile").innerHTML = txt;
}


function ExportFileExcel() {

    $(document).ready(function() {
      $('#mitabla').DataTable( {
            dom: 'Bfrtip',
            buttons: [
                'excelHtml5',
            ]
        } );
      });
}



function descargarTXT(){
  var text = "";

  for(var i=0; i < TABLAGLOBAL[1].length; i++){
    if(TABLAGLOBAL[1][i]) text += TABLAGLOBAL[1][i] + '\n';
  }
  var data = new Blob([text], {type: 'text/plain', endings:'native'});
  var url = window.URL.createObjectURL(data);

  document.getElementById('downloadtxt').href = url;
}


function cargardatos(){
  //alert("hasta aca esta bien");
  var archivo = "Modelos_equipos.txt";
  contenido = loadFile(archivo);

  if(contenido){
    creartabla_bidimensional();
    //alert("tabla cargada");
  }else alert("el archivo NO existe");

  document.getElementById('downloadtxt').style.display = 'none';
  document.getElementById('descargarenexcelcheck').style.display = 'none';
  document.getElementById('descargarenexceltext').style.display = 'none';
}


function lochequearon() {
  // Get the checkbox
  var checkBox = document.getElementById("transportecheck");
  //var cuadrito = document.getElementById("SVLANinput");
  // Get the output text
    // If the checkbox is checked, display the output text
  if (checkBox.checked == true){
    document.getElementById("SVLANinput").style.display = "inline";
    document.getElementById('SVLANtransporte').style.display = "inline";
    document.getElementById('SVLANtransporte').innerHTML = "ingrese la SVLAN de transporte";
    TRANSPORTE = "Si tiene";
    SVLAN = document.getElementById('SVLANinput').value;
  }else {
    document.getElementById("SVLANinput").style.display = "none";
    document.getElementById('SVLANtransporte').style.display = "none";
    TRANSPORTE = "NO TIENE";
    SVLAN = "NO TIENE";
  }
}

function leerArchivo(e) {
  alert("llegue hasta aca");
  var archivo = e.target.files[0];
  if (!archivo) {
		alert("algo salio mal");
    return;
  }
  var lector = new FileReader();
  lector.onload = function(e) {
	//contenido = e.target.result;
  tabladedatos = e.target.result;
    //var contenido = e.target.result;
    //mostrarContenido(contenido);
  };
  lector.readAsText(archivo);

}

function Configfutura(){
  var respuestaDATOS = 0;
  var servicio;
  if(!EquipoFUTURO){
    alert( "todo mal no hay cargado ningun EQUIPO" );
    return;
  }

  PUERTO = document.getElementById('interfaceFUTURA').value;
  //cargar archivo tabla de datos
  var archivo = "TabladeDatos.txt";
  tabladedatos = loadFile(archivo);

  if(tabladedatos){
    //alert( "tabla de datos cargada!");
  }else {
    alert("el archivo NO existe");
    return;
  }

  linealeida = tabladedatos.split('\n'); // aca tengo todo el contenido del archivo cargado en un array de lineas
  cuantosservicios(linealeida);

  for(servicio = 0; servicio < CANTIDAD_DE_SERVICIOS.length; servicio++){

    //alert(CANTIDAD_DE_SERVICIOS[servicio]);
    GuardarDatos();

    if (TRANSPORTE == "NO TIENE"){
      Subinterface_completa = PUERTO + "." + CVLAN;
    } else {
      SVLAN = document.getElementById('SVLANinput').value;
      Subinterface_completa = PUERTO + "." + SVLAN + CVLAN;
    }

    switch (MARCA) {
      case "HUAWEI":
        configFUTURA_HUAWEI(); //esto todavía no esta desarrollado
        break;
      case "CISCO":
        configFUTURA_CISCO();
        break;
      default:
        break;
    }
  }


  alert("fin configuracion futura");
  document.getElementById("downloadtxt").style.display = "inline";
  document.getElementById('descargarenexcelcheck').style.display = 'inline';
  document.getElementById('descargarenexceltext').style.display = 'inline';
  descargarTXT();
  //ExportFileExcel();
}


function cuantosservicios(renglones){
  var servicios = 0;
  var i;
  //alert("servicios " + servicios);

  for (i = 0; i< renglones.length; i++){
    var re = /---------------------------------------/g
    var obtenido = renglones[i].match(re);
    if (obtenido != null) {//encontré un servicio
      CANTIDAD_DE_SERVICIOS[servicios] = i;
      servicios++;
      //alert("servicios " + servicios);

      } //final del if
    } //final del for

  return servicios;
}



function configFUTURA_HUAWEI(){
  HUAWEI_verificar();
  HUAWEI_configVRF();
  HUAWEI_configBGP();
  switch (MODELO) {
    case "CX600-X16A":
      HUAWEI_XA_configSUBINTERFACE();
      break;
    case "CX600-X8A":
      HUAWEI_XA_configSUBINTERFACE();
      break;
    case "CX600-X16":
      HUAWEI_X_configSUBINTERFACE();
      break;
    case "CX600-X8":
      HUAWEI_X_configSUBINTERFACE();
      break;
    default:
      alert("error de modelo HUAWEI ----" + MODELO +"----");
      break;
    } //final del switch
    HUAWEI_estaticas();
    HUAWEI_salidaINTENET();

}


function configFUTURA_CISCO(){
  //alert("funcion CISCO");
  CISCO_verificar();
  CISCO_configVRF();
  CISCO_configBGP();
  CISCO_configSUBINTERFACE();
  CISCO_estaticas();
  CISCO_salidaINTENET();
}


function CISCO_salidaINTENET(){
  indice = 0;
	ipEstatica = "sin datos";
	MASCARA_estatica = "sin datos";
  var PrimeraVEZ = 0;

  /*
  Salida a internet
  router bgp 7303
   address-family ipv4 unicast
    network 181.10.21.16/29 route-policy red-integra

  */

    for (indice = 0; indice < Cant_estaticas; indice++){
      if (Estatica[indice].SalidaINTERNET){
        if (PrimeraVEZ == 0){
          insertarentabla("Salida internet,!  Salida a internet ");
          insertarentabla("Salida internet,router bgp 7303");
          //if(VRF!="NO TIENE")  insertarentabla("Salida internet, vrf "+ VRF); // para el caso de VRF es con neighbor_BGP??
          insertarentabla("Salida internet, address-family ipv4 unicast");
          PrimeraVEZ = 1;
        } //fin del if primera vez

        if (Estatica[indice].SalidaINTERNET != "NO TIENE")
          insertarentabla("Salida internet,   network " + Estatica[indice].SalidaINTERNET + " route-policy red-integra");
      } else{//fin del if (Estatica[indice].SalidaINTERNET)
        //insertarentabla("Salida internet, !   no tiene Salida a internet " + Estatica[indice].IP);
        }
      } //fin del for

      if (PrimeraVEZ == 0) insertarentabla("Salida internet, !   no tiene Salida a internet ");
      insertarentabla("Salida internet,	!---------------------------------------");
}



function CISCO_estaticas() {
  //todavía me falta ver como hacer con las estaticas con salida a internet LAN /29 o /24 esas
	//Estatica(0) = "sin datos"
	indice = 0;
	ipEstatica = "sin datos";
	MASCARA_estatica = "sin datos";
  var PrimeraVEZ = 0;

	//!
	//router static
	// vrf #####VRF#####
	//  address-family ipv4 unicast
	//   #####LOOPBACK#####/32 GigabitEthernet#####INTERFACE L3 X/X/X#####.#####CVLAN##### #####IPWAN CPE##### description REF:#####REFERENCIA#####
	//!

	//!
	//router static
	// address-family ipv4 unicast
	//  #####LOOPBACK#####/32 GigabitEthernet#####INTERFACE L3 X/X/X#####.#####CVLAN##### #####IPWAN CPE##### description REF:#####REFERENCIA#####
	//!

	//if (Estatica[indice].IP != "sin datos" ){


  //alert("hasta aca esta bien HUAWEI SALIDA INTERNET");
  //alert(indice + "    " + Estatica[indice].SalidaINTERNET);


    for (indice = 0; indice < Cant_estaticas; indice++){
      if (Estatica[indice].IP){
        if (PrimeraVEZ == 0){
          insertarentabla("Estaticas,!  Configuracion estaticas ");
          insertarentabla("Estaticas,router static");
          if(VRF!="NO TIENE")  insertarentabla("Estaticas, vrf "+ VRF);
          insertarentabla("Estaticas, address-family ipv4 unicast");
          PrimeraVEZ = 1;
        } //fin del if primera vez

        insertarentabla("Estaticas, " + Estatica[indice].IP + " " + Subinterface_completa + " " + NEXTHOP);


      } else insertarentabla("Estaticas, !  no tiene estaticas ");
  } //fin del for
}

function CISCO_config_Puerto_Fisico(){ //Modelo de ejemplo HUAWEI CX600-X8 - SPE2MU

	insertarentabla( "Interface	,!  Configuracion Puerto Fisico FUTURO" );
	//insertarentabla( "Interface	,!" );
	insertarentabla( "Interface	,interface GigabitEthernet" + PUERTO );
	insertarentabla( "Interface	, description " + Descripcion + " - ACCESO" );
	insertarentabla( "Interface	, load-interval 30" );
	insertarentabla( "Interface	,!" );

  PRIMER_SERVICIO = 1;
}

function CISCO_configSUBINTERFACE(){ //Modelo de ejemplo HUAWEI CX600-X8 - SPE2MU

	//MsgBox "empezando a escribir! ESTADO " & ESTADO
	//MsgBox "interface " & Subinterface_completa & "description " & Descripcion
	if( PRIMER_SERVICIO == 0 ) CISCO_config_Puerto_Fisico();



	insertarentabla( "Subint	,!  Configuracion Subinterface FUTURA" );
	//insertarentabla( "Subint	,!" );
	insertarentabla( "Subint	,interface Gi" + Subinterface_completa ); // ACA SE ARMA Dif(ERENTE SVLAN + CVLAN en caso de Transporte
	insertarentabla( "Subint	, description " + Descripcion );
	insertarentabla( "Subint	, shutdown " );
	if( Bandwidth != "NO TIENE" ) insertarentabla( "Subint	, bandwidth " + Bandwidth );
	if( Policy_input != "NO TIENE" ) insertarentabla( "Subint	, service-policy input " + Policy_input );
	if( Policy_output != "NO TIENE" ) insertarentabla( "Subint	, service-policy output " + Policy_output );

	if( TRANSPORTE == "NO TIENE" ){
		if( VRF != "NO TIENE" ) insertarentabla( "Subint	, vrf " + VRF );
		insertarentabla( "Subint	, ipv4 address " + IpWAN + " " + Mascara_WAN );
		insertarentabla( "Subint	, load-interval 30" );
		insertarentabla( "Subint	, encapsulation dot1q " + CVLAN );
		insertarentabla( "Subint	, !" );
		//insertarentabla( "Subint	,!" );
	}else{
		insertarentabla( "Subint	,ERROR:  No esta contemplado el caso de Conf. FUTURA CISCO por Transporte" );
		alert("ERROR: No esta contemplado el caso de Conf. FUTURA CISCO por Transporte");
		return;
	}

}



function CISCO_configBGP(){

	if ( BGP != "NO TIENE" ){
		insertarentabla( "BGP	,!  BGP a configurar en Cx600" );
		insertarentabla( "BGP	,!" );
		insertarentabla( "BGP	,router bgp 7303 " );

		if ( VRF != "NO TIENE" ) {// si tiene VRF
			insertarentabla( "BGP	, address-family ipv4 unicast " );
			insertarentabla( "BGP	, !" );
			insertarentabla( "BGP	,  vrf " + VRF );
			insertarentabla( "BGP	,   rd " + RD );
			insertarentabla( "BGP	,   address-family ipv4 unicast" );
			insertarentabla( "BGP	,    redistribute connected" );
			insertarentabla( "BGP	,    redistribute static" );
			insertarentabla( "BGP	,   !" );
			if ( BGP_neighbor != "NO TIENE" ){
				insertarentabla( "BGP	,   neighbor " + NEXTHOP );
				insertarentabla( "BGP	,    remote-as " + AS_sistemaautonomo );
				insertarentabla( "BGP	,    use neighbor-group VPN-Prefix-500" );
				insertarentabla( "BGP	,    description " + Descripcion_BGP );
				insertarentabla( "BGP	,    address-family ipv4 unicast" );
				insertarentabla( "BGP	,    maximum-prefix 500 90 restart 10" );
				insertarentabla( "BGP	,    !" );
				insertarentabla( "BGP	,   !" );
				insertarentabla( "BGP	,  !" );
			} else {//tiene BGP pero no tiene VRF
			if ( BGP_neighbor != "NO TIENE" ){
				insertarentabla( "BGP	, !" );
				insertarentabla( "BGP	, neighbor " + NEXTHOP );
				insertarentabla( "BGP	,   remote-as " + AS_sistemaautonomo );
				insertarentabla( "BGP	,  use neighbor-group INTERNET" );
				insertarentabla( "BGP	,   description " + Descripcion_BGP );
				insertarentabla( "BGP	,   address-family ipv4 unicast" );
				insertarentabla( "BGP	,    route-policy red-integra in" );
				insertarentabla( "BGP	,   !" );
				insertarentabla( "BGP	,  !" );
				insertarentabla( "BGP	, !" );
  			} // if ( BGP_neighbor != "NO TIENE" )
  		  } //tiene BGP pero no tiene VRF
      } // si tiene VRF

	 } else insertarentabla( "BGP	,!  NO TIENE BGP" ); //si no tiene BGP

}

function CISCO_configVRF(){

  var RD_imp = new Array();
  var RD_exp = new Array();

  alert ("CISCO_configVRF hasta aca esta bien");

  if (VRF != "NO TIENE") {
    insertarentabla( "VRF,!  VRF a configurar en CISCO ASR9K" );
    insertarentabla( "VRF,!" );
    insertarentabla( "VRF,vrf " + VRF );
    if (Descripcion_BGP != "NO TIENE") insertarentabla( "VRF, description " + Descripcion_BGP );
    insertarentabla( "VRF, address-family ipv4 unicast" );
    if ( RD_import != "sin datos" ){ //si lo hago con RD_export = "sin datos" seria lo mismo
      RD_imp = RD_import.split(' '); // los separo por espacio
      RD_exp = RD_export.split(' '); // los separo por espacio
      insertarentabla( "VRF,  import route-target " );
      for (var i = 0; i < RD_imp.length; i++){
        insertarentabla( "VRF,   "+ RD_imp[i] );
      }
      insertarentabla( "VRF,  !" );
      insertarentabla( "VRF,  export route-policy loopback-vpn" );
      insertarentabla( "VRF,  export route-target" );
      for (var i = 0; i < RD_exp.length; i++){
        insertarentabla( "VRF,   "+ RD_exp[i] );
      }
      insertarentabla( "VRF,  !" );
      insertarentabla( "VRF, !" );
      insertarentabla( "VRF,!" );
      //alert("RD_import " + RD_import + "   RD_export " + RD_export); //esto funciona perfecto
    } else { // aca termina el nuevo modelo de VRF
    insertarentabla( "VRF,   7303:2" );
    insertarentabla( "VRF,   " + RD );
    insertarentabla( "VRF,  !" );
    insertarentabla( "VRF,  export route-policy loopback-vpn" );
    insertarentabla( "VRF,  export route-target" );
    insertarentabla( "VRF,   " + RD );
    insertarentabla( "VRF,  !" );
    insertarentabla( "VRF, !" );
    insertarentabla( "VRF,!" );
    }
  }	else{
    insertarentabla( "VRF,#  NO TIENE VRF" );
  }


}




function CISCO_verificar()
{
  //alert("funcion CISCO verificar");
  insertarentabla("verificar,!  Verificaciones Previas ");
  if (VRF != "NO TIENE"){
    insertarentabla("verificar,show running-config vrf " + VRF);
    insertarentabla("verificar,sh run router bgp 7303 vrf " + VRF + " neighbor " + NEXTHOP);
  }
  insertarentabla("verificar,show running-config interface GigabitEthernet" + PUERTO );
  insertarentabla("verificar,show running-config interface GigabitEthernet" + Subinterface_completa);
  insertarentabla("verificar,!  " );
}

function HUAWEI_salidaINTENET(){
  //indice_salidaINTERNET = 0;
  indice = 0;
	ipEstatica = "sin datos";
	MASCARA_estatica = "sin datos";
	var PrimeraVEZ = 0;

  //alert("hasta aca esta bien HUAWEI SALIDA INTERNET");
  //alert(indice + "    " + Estatica[indice].SalidaINTERNET);
  if (Estatica[indice].SalidaINTERNET){
  for (indice = 0; indice < Cant_estaticas; indice++){

    if (Estatica[indice].SalidaINTERNET != "NO TIENE"){
      if (PrimeraVEZ == 0){
        insertarentabla("Salida internet,#  Configuracion Salida internet ");
        insertarentabla("Salida internet,#");
        insertarentabla("Salida internet,bgp 7303");
        insertarentabla("Salida internet, ipv4-family unicast");
        insertarentabla("Salida internet,  undo synchronization");
        PrimeraVEZ = 1;
      }

			//calcular la ip y su mascara

      var re = /(\d+.\d+.\d+.\d+.)\/(\d+)/g
      var obtenido = Estatica[indice].SalidaINTERNET.match(re);
      if (obtenido != null) {
        ipEstatica = Estatica[indice].SalidaINTERNET.replace(re,"$1");
        MASCARA_estatica = Estatica[indice].SalidaINTERNET.replace(re,"$2").replace(/^\s\s*/, '').replace(/\s\s*$/, ''); // le quito los espacios en blanco;

      }else{
        ipEstatica = "Error";
				MASCARA_estatica = "Error";
      }

      MASCARA_estatica = convertirMASCARA(MASCARA_estatica);

			//network   181.13.43.32 255.255.255.248 route-policy RED-INTEGRA
			//network   ESTATICA MASCARA_estatica route-policy RED-INTEGRA
			insertarentabla("Salida internet,   network " + ipEstatica + " " + MASCARA_estatica + " route-policy RED-INTEGRA");
		} else insertarentabla("Salida internet,#  no tiene Salida a internet " + Estatica[indice].IP);
    // fin del if Estatica[indice].SalidaINTERNET != "NO TIENE"
    } // fin del ciclo for
  } // fin del primer if

	if (PrimeraVEZ == 0) {
		insertarentabla("Salida internet,#  no tiene Salida a internet ");
	}
    insertarentabla("Salida internet,#---------------------------------------");

    //Estatica[indice].SalidaINTERNET = "sin datos";
    //Estatica[0].SalidaINTERNET = "sin datos";

}

function HUAWEI_estaticas(){ //esto esta perfecto!!!
	//todavía me falta ver como hacer con las estaticas con salida a internet LAN /29 o /24 esas
	indice = 0;
  //alert("Cant_estaticas " + Cant_estaticas);
	ipEstatica = "sin datos";
	MASCARA_estatica = "sin datos";

	if (Estatica[indice].IP){
		insertarentabla("Estaticas	,#  Configuracion estaticas " );
		//insertarentabla("Estaticas	,#" );
    //alert("hasta aca esta bien ESTATICAS indice < Cant_estaticas " + indice + " " + Cant_estaticas);
    for (indice = 0; indice < Cant_estaticas; indice++){ //todavía no sabes la cantidad de estaticas
			//calcular la ip y su mascara
      //alert("hasta aca esta bien ESTATICAS indice < Cant_estaticas " + indice + " " + Cant_estaticas);
      //alert(Estatica[indice].IP);
      if(Estatica[indice].IP){
        var re = /(\d+.\d+.\d+.\d+.)\/(\d+)/g
        var obtenido = Estatica[indice].IP.match(re);

        if (obtenido != null) {
          //crt.Dialog.MessageBox(readline.replace(/remote-as (\d+)/,"$1")); //vrf_estado = "import";

          ipEstatica = Estatica[indice].IP.replace(re,"$1");
          MASCARA_estatica = Estatica[indice].IP.replace(re,"$2").replace(/^\s\s*/, '').replace(/\s\s*$/, ''); // le quito los espacios en blanco;;

          MASCARA_estatica = convertirMASCARA(MASCARA_estatica);


  			if (VRF != "NO TIENE"){
          insertarentabla("Estaticas, ip route-static vpn-instance " + VRF + " " + ipEstatica + " " + MASCARA_estatica + " " + Subinterface_completa + " " + NEXTHOP);
        }	else{
          insertarentabla("Estaticas, ip route-static " + ipEstatica + " " + MASCARA_estatica + " " + Subinterface_completa + " " + NEXTHOP);
          }
        } //final del if (obtenido != null)
      }

    } //final del for
  } else insertarentabla("Estaticas,#  no tiene estaticas ");



  	//ip route-static 181.96.224.50 255.255.255.255 g2/2/4.2 181.96.17.54
  	//tengo que reiniciar las estaticas
  	//Estatica[indice].IP = "sin datos";
  	//Estatica[0].IP = "sin datos";


}

function convertirMASCARA(MASCARA){
  switch (MASCARA) {
      case "32":
        MASCARA = "255.255.255.255";
        break;
      case "30":
        MASCARA = "255.255.255.252";
        break;
      case "29":
        MASCARA = "255.255.255.248";
        break;
      case "28":
        MASCARA = "255.255.255.240";
        break;
      case "27":
        MASCARA = "255.255.255.224";
        break;
      case "26":
        MASCARA = "255.255.255.192";
        break;
      case "24":
        MASCARA = "255.255.255.0";
        break;
      default:
        MASCARA = "ERROR_MASCARA_" + MASCARA;
        alert(MASCARA);
        break;
      } //final del switch

      return MASCARA;
}


function HUAWEI_XA_configSUBINTERFACE() { //Modelo de ejemplo HUAWEI CX600-X8A-16A - LOP4MU

  insertarentabla( "Subint,#  Configuracion Subinterface FUTURA" );
  //insertarentabla( "Subint	,#" );
  if( TRANSPORTE == "NO TIENE" ){
    insertarentabla( "Subint,interface " + Subinterface_completa );
    insertarentabla( "Subint, shutdown " );
    insertarentabla( "Subint, bandwidth " + Bandwidth );
    insertarentabla( "Subint, description " + Descripcion );
    if(VRF != "NO TIENE")insertarentabla( "Subint, ip binding vpn-instance " + VRF );
    insertarentabla( "Subint, ip address " + IpWAN + " " + Mascara_WAN );
    insertarentabla( "Subint, statistic enable" );
    insertarentabla( "Subint, encapsulation dot1q-termination" );
    insertarentabla( "Subint, dot1q termination vid " + CVLAN );
    insertarentabla( "Subint, arp broadcast enable " );
    insertarentabla( "Subint, trust upstream principal" );
    insertarentabla( "Subint, trust 8021p " );
    if(Policy_input != "NO TIENE") insertarentabla( "Subint, qos-profile " + Policy_input + " inbound vlan " + CVLAN + " identifier none " );
    if(Policy_output != "NO TIENE") insertarentabla( "Subint, qos-profile " + Policy_output + " outbound vlan " + CVLAN + " identifier none "  );
    insertarentabla( "Subint,#" );
  }else{ //si no tiene TRANSPORTE entonces lo configuro de la siguiente manera
    insertarentabla( "Subint, interface " + Subinterface_completa );
    insertarentabla( "Subint, shutdown " );
    insertarentabla( "Subint, bandwidth " + Bandwidth );
    insertarentabla( "Subint, description " + Descripcion );
    if(VRF != "NO TIENE") insertarentabla( "Subint, ip binding vpn-instance " + VRF );
    insertarentabla( "Subint, ip address " + IpWAN + " " + Mascara_WAN );
    insertarentabla( "Subint, statistic enable" );
    insertarentabla( "Subint, encapsulation qinq-termination" );
    insertarentabla( "Subint, qinq termination pe-vid " + SVLAN + " ce-vid " + CVLAN );
    insertarentabla( "Subint, arp broadcast enable " );
    insertarentabla( "Subint, trust upstream principal" );
    insertarentabla( "Subint, trust 8021p " );
    if(Policy_input != "NO TIENE") insertarentabla( "Subint, qos-profile " + Policy_input + " inbound pe-vid " + SVLAN + " ce-vid " + CVLAN + " identifier none " );
    if(Policy_output != "NO TIENE") insertarentabla( "Subint, qos-profile " + Policy_output + " outbound pe-vid " + SVLAN + " ce-vid " + CVLAN + " identifier none "  );
    insertarentabla( "Subint, #" );
  }


}




function HUAWEI_X_configSUBINTERFACE() { //Modelo de ejemplo HUAWEI CX600-X8 - SPE2MU

	insertarentabla( "Subint	,#  Configuracion Subinterface FUTURA" );
	//insertarentabla( "Subint	,#" );
	if ( TRANSPORTE == "NO TIENE" ){
		insertarentabla( "Subint	,interface " + Subinterface_completa );
    insertarentabla( "Subint	,description " + Descripcion );
    insertarentabla( "Subint	,shutdown " );
    insertarentabla( "Subint	,bandwidth " + Bandwidth );
		insertarentabla( "Subint	,control-vid " + CVLAN + " dot1q-termination" );
		insertarentabla( "Subint	,dot1q termination vid " + CVLAN );
		if ( VRF != "NO TIENE" ) insertarentabla( "Subint	,ip binding vpn-instance " + VRF );
		insertarentabla( "Subint	,ip address " + IpWAN + " " + Mascara_WAN );
		insertarentabla( "Subint	,arp broadcast enable " );
		insertarentabla( "Subint	,trust upstream principal " );
		insertarentabla( "Subint	,trust 8021p " );
		if (Policy_input != "NO TIENE") insertarentabla( "Subint	,qos-profile " + Policy_input + " inbound vlan " + CVLAN + " identifier none " );
		if (Policy_output != "NO TIENE") insertarentabla( "Subint	,qos-profile " + Policy_output + " outbound vlan " + CVLAN + " identifier none "  );
		insertarentabla( "Subint	,statistic enable" );
		insertarentabla( "Subint	,#" );
	}else{
		insertarentabla( "Subint	,interface " + Subinterface_completa ); // ACA SE ARMA DIFERENTE SVLAN + CVLAN
		insertarentabla( "Subint	,description " + Descripcion );
		insertarentabla( "Subint	,shutdown " );
		insertarentabla( "Subint	,bandwidth " + Bandwidth );
		insertarentabla( "Subint	,control-vid " + CVLAN + " qinq-termination" ); //aca la CVLAN en realidad podria ser cualquier id
		insertarentabla( "Subint	,qinq termination pe-vid " + SVLAN + " ce-vid " + CVLAN ); //SVLAN y CVLAN
		if ( VRF != "NO TIENE" ) insertarentabla( "Subint	,ip binding vpn-instance " + VRF );
		insertarentabla( "Subint	,ip address " + IpWAN + " " + Mascara_WAN );
		insertarentabla( "Subint	,arp broadcast enable " );
		insertarentabla( "Subint	,trust upstream principal " );
		insertarentabla( "Subint	,trust 8021p " );
		if ( Policy_input != "NO TIENE" ) insertarentabla( "Subint	,qos-profile " + Policy_input + " inbound pe-vid " + SVLAN +" ce-vid " + CVLAN + " identifier none " );
		if ( Policy_output != "NO TIENE" ) insertarentabla( "Subint	,qos-profile " + Policy_output + " outbound pe-vid " + SVLAN +" ce-vid " + CVLAN + " identifier none "  );
		insertarentabla( "Subint	,statistic enable" );
		insertarentabla( "Subint	,#" );
	}

}


function HUAWEI_configBGP(){

  if (BGP != "NO TIENE"){
    insertarentabla( "BGP	,#  BGP a configurar en Cx600" );
  	insertarentabla( "BGP	,#" );
  	insertarentabla( "BGP	,bgp 7303 " );
  	if (VRF != "NO TIENE"){ //si tiene VRF
      insertarentabla( "BGP	,ipv4-family vpn-instance " + VRF );
      insertarentabla( "BGP	,import-route direct" );
      insertarentabla( "BGP	,import-route static" );
      if (BGP_neighbor != "NO TIENE") {
        insertarentabla( "BGP	,peer " + NEXTHOP + " as-number " + AS_sistemaautonomo );
        insertarentabla( "BGP	,peer " + NEXTHOP + " description " + Descripcion_BGP );
    		insertarentabla( "BGP	,peer " + NEXTHOP + " route-limit 100 90 idle-timeout 10" );
    		insertarentabla( "BGP	,peer " + NEXTHOP + " advertise-community " );
    		insertarentabla( "BGP	,peer " + NEXTHOP + " advertise-ext-community " );
    		insertarentabla( "BGP	,#" );
      }
    }else{ //tiene BGP pero no tiene VRF
      if (BGP_neighbor != "NO TIENE"){
        insertarentabla( "BGP	,peer " + NEXTHOP + " as-number " + AS_sistemaautonomo );
    		insertarentabla( "BGP	,peer " + NEXTHOP + " description " + Descripcion_BGP );
    		insertarentabla( "BGP	,ipv4-family unicast" );
    		insertarentabla( "BGP	, peer " + NEXTHOP + " enable" );
    		insertarentabla( "BGP	, peer " + NEXTHOP + " route-policy distribute-list_120 export" );
    		insertarentabla( "BGP	, peer " + NEXTHOP + " route-limit 100 90 idle-timeout 10" );
    		insertarentabla( "BGP	, peer " + NEXTHOP + " default-route-advertise" );
    		insertarentabla( "BGP	, peer " + NEXTHOP + " route-policy RED-INTEGRA import" );
    		insertarentabla( "BGP	,#" );
      }
    }
  }else { //si no tiene BGP
    insertarentabla( "BGP	,#  NO TIENE BGP" );
  }
}




function HUAWEI_configVRF() {
  /*
  #
  ip vpn-instance wifioffload-vpn
   ipv4-family
    route-distinguisher 7303:1312
    vpn-target 7303:1312 export-extcommunity
    vpn-target 7303:1312 import-extcommunity
  #
  */


  if ( VRF != "NO TIENE" ){
    insertarentabla( "VRF,#  VRF a configurar en Cx600" );
    insertarentabla( "VRF,#" );
    insertarentabla( "VRF,ip vpn-instance " + VRF );
    insertarentabla( "VRF, ipv4-family " );
    if ( RD_import != "sin datos" ){ //si lo hago con RD_export = "sin datos" seria lo mismo
  		insertarentabla( "VRF,  route-distinguisher " + RD_export );
  		insertarentabla( "VRF,  export route-policy loopback-vpn " );
  		insertarentabla( "VRF,  vpn-target " + RD_export + " export-extcommunity" );
  		insertarentabla( "VRF,  vpn-target " + RD_import + " import-extcommunity" );
  		insertarentabla( "VRF,#" );
      //alert("RD_import " + RD_import + "   RD_export " + RD_export); //esto funciona perfecto
    } else { // aca termina el nuevo modelo de VRF
      insertarentabla( "VRF,  route-distinguisher " + RD );
  		insertarentabla( "VRF,  export route-policy loopback-vpn " );
  		insertarentabla( "VRF,  vpn-target " + RD + " export-extcommunity" );
  		insertarentabla( "VRF,  vpn-target " + RD + " 7303:2 import-extcommunity" );
  		insertarentabla( "VRF,#" );
    }

  } else{
    insertarentabla( "VRF,#  NO TIENE VRF" );
  }
}

function HUAWEI_verificar()
{
  //alert("funcion HUAWEI verificar");
  insertarentabla("verificar,#  Verificaciones Previas " );

	if (Policy_input != "NO TIENE") insertarentabla("verificar,disp curr configuration qos-profile " + Policy_input );
	if (Policy_output != "NO TIENE") insertarentabla("verificar	,disp curr configuration qos-profile " + Policy_output );
	if (VRF != "NO TIENE") insertarentabla("verificar	,disp curr conf vpn-instance " + VRF );
	if (BGP != "NO TIENE"){
    if (VRF != "NO TIENE"){
        insertarentabla("verificar	,disp curr conf bgp | begin " + VRF );
    } else{
      insertarentabla("verificar	,disp curr conf bgp | begin " + NEXTHOP );
    }
  }
	insertarentabla("verificar,disp int desc | inc " + PUERTO );
	insertarentabla("verificar,disp int desc | inc " + Subinterface_completa );
	insertarentabla("verificar,disp curr | inc " + NEXTHOP );
	insertarentabla("verificar,#  " );
}







function GuardarDatos(){
  //var linealeida = tabladedatos.split('\n'); // aca tengo todo el contenido del archivo cargado en un array de lineas
  /*delete Estatica;
  Estatica = new Array(); //creo el array de estaticas
  Estatica[0] = "sin datos";*/
  var DATO;
  var i;
  RD_import = "sin datos";
  RD_export = "sin datos";
  TIPOdeDATO = "ninguno";

  //alert("arranco a leer la linea " + (CANTIDAD_DE_SERVICIOS[Servicio_index]+1));
  //alert(linealeida[(CANTIDAD_DE_SERVICIOS[Servicio_index]+1)]);
  for (i = CANTIDAD_DE_SERVICIOS[Servicio_index]+1; TIPOdeDATO != "Resultado ping WAN+1"; i++){
    //var re = /(\w+): ([\S\s]+)/g		//crea una regExp a partir de ese string leido
    var re = /^([\S\s]+): ([\S\s]+)$/g
    var obtenido = linealeida[i].match(re);
    if (obtenido != null) {
      //crt.Dialog.MessageBox(readline.replace(/remote-as (\d+)/,"$1")); //vrf_estado = "import";
      TIPOdeDATO = linealeida[i].replace(re,"$1"); //vrf_estado = "import";
      DATO = linealeida[i].replace(re,"$2"); //vrf_estado = "import";
      DATO = DATO.replace(/^\s\s*/, '').replace(/\s\s*$/, ''); // le quito los espacios en blanco;
      //alert("TIPOdeDATO " + TIPOdeDATO + "   DATO " + DATO);

      switch (TIPOdeDATO) {
        case "Descripcion":  //Descripcion:	VPN Branch Office - Cliente - Ref:263534 Lin:2721088 Acc:CHL21Dplaca14puerto0
  				Descripcion = DATO;
          break;
        case "Ip de WAN": //Ip de WAN:	192.168.81.233
    			IpWAN = DATO;
          break;
        case "Bandwidth": //Bandwidth:	1024
          Bandwidth = DATO;
          break;
        case "Policy input": //Policy input:	VPN-1M
          Policy_input = DATO;
          break;
        case "Policy output": //Policy output:	FUERA-DE-PRODUCTO-1M
          Policy_output = DATO;
          break;
        case "VRF": //VRF:	kellerhoff-vpn
          VRF = DATO;
          break;
        case "Mascara WAN": //Mascara WAN:	 255.255.255.252
          Mascara_WAN = DATO;
          break;
        case "WAN+1":
          NEXTHOP = DATO;
          break;
        case "CVLAN": //CVLAN:
          CVLAN = DATO;
          break;
        case "RD": //RD:	7303:1164
          RD = DATO;
          break;
        case "RD import": //RD import: 7303:2 7303:128
          RD_import = DATO;
          break;
        case "RD export": //RD export: 7303:128
          RD_export = DATO;
          break;
        case "BGP":
          BGP = DATO;
          break;
        case "BGP neighbor":
          BGP_neighbor = DATO;
          break;
        case "AS": //AS:
          AS_sistemaautonomo = DATO;
          break;
        case "Descripcion BGP": //Descripcion BGP:
          Descripcion_BGP = DATO;
          break;
        case "Estatica": //Estatica: 181.96.224.50/32
          Estatica[indice] = new clase_ESTATICAS;
          Estatica[indice].IP = DATO;
  			  //Estatica[indice] = DATO;
          //alert("ESTATICA " + Estatica[indice].IP);
  				//indice++;
          break;
        case "Salida a INTERNET":
         //Salida a INTERNET: 181.10.172.8/29
  			 //Salida a INTERNET: NO TIENE
  				//SalidaINTERNET[indice_salidaINTERNET] = DATO;
  				//indice_salidaINTERNET++;
          Estatica[indice].SalidaINTERNET = DATO;
          //alert("CVLAN "+ CVLAN +" ESTATICA " + Estatica[indice].IP + " Salida INTERNET " + Estatica[indice].SalidaINTERNET);
          indice++;
        case "Resultado ping WAN+1": //Resultado ping WAN+1:	192.168.81.234 Success rate is 100 percent (5/5)
          //TIPOdeDATO = "ninguno";
          Servicio_index++;
          break;
        default:
          break;
        } //final del switch

      } //final del if
    } //final del for

  Cant_estaticas = Estatica.length;
  ULTIMAlinealeida = i;
  //Cant_IPSalidaInternet = SalidaINTERNET.length;
  //alert("termine de leer un servicio ULTIMAlinealeida " + ULTIMAlinealeida);
} // final de la funcion




function buscarmodelo(){
	var columna, fila;
	var Equipo_buscado = document.getElementById('equipoabuscar').value;
	Equipo_buscado = Equipo_buscado.toUpperCase(); //lo paso a mayusculas
	var comparar;
	//alert( Equipo_buscado + " " + tabla_con_datos[0][3]); //aca muestro la columna 0 fila 3

	for(fila = 0; fila < cant_filas; fila++){
		comparar = Equipo_buscado.localeCompare(tabla_con_datos[0][fila]);
		if (comparar == 0){
			//alert("encontre el equipo en la fila " + fila); //esto esta bien
			var mensaje = Equipo_buscado + " es un " + tabla_con_datos[2][fila] + " "  + tabla_con_datos[3][fila];
			//alert(mensaje); //esto esta bien
			document.getElementById('respuesta').innerHTML = mensaje;
      EquipoFUTURO = Equipo_buscado;
      MARCA = tabla_con_datos[2][fila];
      MODELO = tabla_con_datos[3][fila].replace(/^\s\s*/, '').replace(/\s\s*$/, '');
			return;
		}
	}

	if (fila == cant_filas) alert("no se encontró el equipo " + Equipo_buscado); //esto esta bien

	return;
}



function creartabla_bidimensional(){
		lines = contenido.split('\n'); // aca tengo todo el contenido del archivo cargado en un array de lineas

		cant_filas = lines.length;
		var filaleida = lines[0].split(',');
		cant_columnas = filaleida.length;
		var datosenfila;
		var columna, fila;
		//alert("hasta aca esta bien");

		//Declaramos el array bidimensional
		//var tabla_con_datos = new Array(cant_columnas);
		tabla_con_datos = new Array(cant_columnas);
		for(var columna=0; columna<cant_columnas; columna++){
				tabla_con_datos[columna] = new Array(cant_filas);
		}
		//Metemos un dato en cada posición

		for( fila = 0; fila < cant_filas; fila++){
			filaleida = lines[fila].split(',');
			for( columna = 0; columna < cant_columnas; columna++){
				tabla_con_datos[columna][fila] = filaleida[columna];
			}
		} //alert(tabla_con_datos[1][2]); //aca muestro la fila 1 columna 2

}


function insertarentabla(texto){ 	//esto anda bien pero solo agrega un fila en la tabla
	//var readline = filedatos.Readline(); //aca lee la subinterface
	var texto_en_celda = texto.split(','); // aca tengo todo el contenido del archivo cargado en un array de lineas
	//alert(texto_en_celda.length); //esto esta perfecto //dice que son 5
	var tablacuerpo = document.getElementById("tablacuerpo"); // identifico el cuerpo de la tablacuerpo
	var newRow   = tablacuerpo.insertRow(); // inserto una nueva fila

	for (var i = texto_en_celda.length -1; i >= 0; i-- ){
		var newCell  = newRow.insertCell(0);	 //inserto una celda en la fila en el indice 0
		var newText = document.createTextNode(texto_en_celda[i]);
    TABLAGLOBAL[i][fila_tabla] = texto_en_celda[i]; // en este caso entonces i es la columna
		newCell.appendChild(newText);
	}
  fila_tabla++;
}

function loadFile(filePath) {
  var result = null;
  var xmlhttp = new XMLHttpRequest();
  xmlhttp.open("GET", filePath, false);
  xmlhttp.send();
  if (xmlhttp.status==200) {
    result = xmlhttp.responseText;
  }
  return result;
}




function infoARCHIVO(evt) {
    var files = evt.target.files; // FileList object // objeto FileList
    // files is a FileList of File objects. List some properties.
    var output = [];
		//var output;
    for (var i = 0, f; f = files[i]; i++) {
      output.push('<li><strong>', escape(f.name), '</strong> (', f.type || 'n/a', ') - ',
                  f.size, ' bytes, last modified: ',
                  f.lastModifiedDate.toLocaleDateString(), '</li>');
    }
    document.getElementById('list').innerHTML = '<ul>' + output.join('') + '</ul>';
		//document.getElementById('list').innerHTML = "algo cambio";
}
