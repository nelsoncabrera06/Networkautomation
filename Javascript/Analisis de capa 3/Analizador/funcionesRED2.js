var contenido;

var EQUIPO = "sin datos";
var MARCA = "sin datos";
var SUBINTERFACES = "sin datos";
var Subint, Cant_subint;
var lines;
var SUBINTERFACE = "sin datos";
var Procesando = false;
var Texto_en_Archivo = "Analisis Configuracion Actual";
var Salida;
var miString;

//lo siguiente hace que el gif de cargando no lo muestre al principio
document.getElementById('loading').style.display = 'none';

document.getElementById('mitabla').style.display = 'none';


var tableRef = document.getElementById('mitabla').getElementsByTagName('tbody')[0];

document.getElementById('file-input').addEventListener('change', leerArchivo, false);
document.getElementById('file-input').addEventListener('change', infoARCHIVO, false);

function leerArchivo(e) {
  
  var archivo = e.target.files[0];
  if (!archivo) {
    return;
  }
  var lector = new FileReader();
  lector.onload = function(e) {
	contenido = e.target.result;
    //var contenido = e.target.result;
    //mostrarContenido(contenido);
  };
  lector.readAsText(archivo);
  
 
}

function mostrarContenido(contenido) {
  var elemento = document.getElementById('contenido-archivo');
  elemento.innerHTML = contenido;
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




(function crear() {
	var textFile = null,
	makeTextFile = function crear(text) {
    var data = new Blob([text], {type: 'text/plain'});

    //If we are replacing a previously generated file we need to
    //manually revoke the object URL to avoid memory leaks.
    if (textFile !== null) {
      window.URL.revokeObjectURL(textFile);
    }

    textFile = window.URL.createObjectURL(data);

    return textFile;
  };



   /* ORIGINAL
    var create = document.getElementById('create'),
    textbox = document.getElementById('textbox');

  create.addEventListener('click', function () {
    var link = document.getElementById('downloadlink');
    link.href = makeTextFile(textbox.value);
    link.style.display = 'block';*/
	
  var create = document.getElementById('create');
  //var create = document.getElementById('runscript');
  //var textbox = document.getElementById('contenido-archivo');
	
	//alert(document.getElementById("contenido-archivo").innerHTML);
	
	//var mitexto = "preto";
	
  create.addEventListener('click', function crear(){
	  
	//alert(mitexto); //esto esta perfecto  
	  
	var link = document.getElementById('downloadlink');
    link.href = makeTextFile(Texto_en_Archivo);
    link.style.display = 'block';
  }, false);
})()


function Analizar() {
	//leo la primera linea del archivo del EQUIPO
	lines = contenido.split('\n');
	//RP/0/9/CPU0:C2BELGRANO#
	
	
	
	//document.getElementById('loading').style.display = 'block'; //muestra el loading...
	
	MARCA_EQUIPO(lines[0]);
	
	//alert("llegue al loading"); //esto esta perfecto
}

function MARCA_EQUIPO(primeralinea)
{
	var inicio, fin;
	var lngPos = primeralinea.indexOf("#");
	
	Procesando = true;
	document.getElementById('loading').style.display = 'block'; //muestra el loading...
	
	if (lngPos!=-1) {
	  //  estoy en un equipo CISCO
	  MARCA = "CISCO";
	  inicio = primeralinea.indexOf(":")+1;
	  fin = primeralinea.indexOf("#");
	  EQUIPO = primeralinea.substring(inicio,fin); // ok! trae el nombre del EQUIPO para el caso CISCO
	  
	  //esto anda perfecto para una sola subinterface
	  //SUBINTERFACE = document.getElementById("subinterface").value;
	  //pero si fueran muchas?? ahi acabo de crear un textarea en html creo que va mejor con lo que quiero hacer
	  SUBINTERFACES = document.getElementById("subint").value;
	  Subint = SUBINTERFACES.split('\n');    // ok! tengo todas las subinterfaces en un array
	  Cant_subint = Subint.length;
	  
	  //alert("Cant_subint " + Cant_subint);
	  
		for (var i = 0; i < Cant_subint; i++) {
			
			SUBINTERFACE = Subint[i];
			//alert(Subint); //esto esta perfecto
			funcion_CISCO();
		}
	  
	} else {
	  //  sino estoy en HUAWEI - esto habria que chequearlo
	  MARCA = "HUAWEI";
	}
	
	Procesando = true;
	document.getElementById('loading').style.display = 'none'; // no anda muy bien el loading revisar dsp
	//messagetoSend = messagetoSend.replace("\n", "<br />");
	Texto_en_Archivo = document.getElementById("contenido-archivo").innerHTML;
	Texto_en_Archivo = Texto_en_Archivo.replace(/<br>/g, "\n");
	Salida = Texto_en_Archivo.split('\n');
	//Texto_en_Archivo = Texto_en_Archivo.split(',').join('\n');
	
	
	
}

function WriteToFile(){
	
	
	
	var myArr = ["Orange", "Banana", "Mango", "Kiwi" ];
	console.log(myArr);
	 
}


function funcion_CISCO()
{
	/*
	var TABLA(30,10);
	var SVLAN(30);
	var Estatica(50);
	var pingEstatica(50);
	var SalidaINTERNET(50);
	var MASCARA_estatica(50); 
	var ipPING(50);
	var Resultado_pingEstatica(50);
	var Redes_aprendidas(50);
	var linea;
	var bgp_Array;
	var exclamacion, max_exclamacion;
	var exclamacion = 0;
	var max_exclamacion = 0;
	
	var Cant_redes = 0;
	var aux = 0;
	var indice = 0;
	var indice_SVLAN = 0;
	var indice_max = 0;
	var i_redes = 0;
	var Estatica(0) = "NO TIENE";
	var SalidaINTERNET(0) = "NO TIENE"; */
	var _static = "NO TIENE";
	var Cant_estaticas = 0;
	var ultima_linea_del_bloque;
	var PUERTO = "NO TIENE";
	var Salir = 0;
	var linealeida, lngPos, linea_recortada, StringsArray, encontrado;
	var VRF = "NO TIENE";
	var ip_WAN = "NO TIENE";
	var mascara_WAN = "NO TIENE";
	var NEXTHOP = "NO TIENE";
	var BGP = "NO TIENE";
	var neighbor_BGP = "NO TIENE";
	var AS_actual = "NO TIENE";
	var policy_input = "NO TIENE";
	var policy_output = "NO TIENE";
	var BANDWIDTH = "NO TIENE";
	var DESCRIPCION = "NO TIENE";
	var RD = "NO TIENE";
	var PEAKFLOW = "NO TIENE";
	var descripcion_BGP = "NO TIENE";
	/*var Subint = "NO TIENE";
	//SUBINTERFACE = "NO TIENE";
	var Estado_Subint = "NO TIENE";
	var CVLAN = "NO TIENE";
	
	var PalabrasCLAVE = array("arg", "ARG", "GMT");
	var Trafico_input = 0;
	var Trafico_output = 0;
	
	*/
	PUERTO = SUBINTERFACE;
	
	var elemento = document.getElementById('contenido-archivo');
	
	miString = "	Descripcion: sh run int " + SUBINTERFACE ;
	elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + miString + "<br/>";
	agregar_fila(miString);
	
	for (var i = 0; i < SUBINTERFACE.length; i++) {
		var Caracter = SUBINTERFACE.charAt(i);
		
		if (!isNaN(Caracter)) {
			linea_recortada = SUBINTERFACE.substring(i, SUBINTERFACE.length);
			//alert(linea_recortada); //esto esta perfecto
			break; 
		} 
	}
	
	
	
	/*
	interface TenGigE0/5/0/16.18140005
	 description VPN - GPBA - Ref.:712583 - Linea.:2821850 - vrf gpba
	 bandwidth 10000
	 service-policy input VPN-10M
	 service-policy output rt0-mc90-10M
	 vrf vpn-gpba
	 ipv4 address 10.247.81.1 255.255.255.252
	 encapsulation dot1q 1814 second-dot1q 5
	!
	*/
	
	//Aca arranco a ver la configuracion actual de la subinterface
	for ( var i = 0; i < lines.length; i++)
	{
		linealeida = lines[i];
		lngPos = linealeida.indexOf(linea_recortada);
		if (lngPos >= 0) { //encontró la interface
			for ( i = i ; Salir != 1; i++){
				linealeida = lines[i];
				miString = "	" +	linealeida ;
				agregar_fila(miString);
				elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + miString;
				
				StringsArray = ["nelson", "#", "bandwidth ", "policy input ", "policy output ", "vrf ", "ipv4 address ", "dot1q vlan ", "interface ", "description ", "encapsulation dot1q", "flow ipv4 monitor"];
				encontrado = encontrarString(linealeida, StringsArray);
				
				
			//	alert(linealeida + " encontrado " + encontrado);
				
				switch(encontrado) {
				  case 3: // bandwidth
					BANDWIDTH = linealeida.substring(10, linealeida.length);
					BANDWIDTH = BANDWIDTH.trim();
					//alert("BANDWIDTH " + BANDWIDTH);
					break;
				  case 4: // policy input
				    policy_input = linealeida.substring(22, linealeida.length);
					policy_input = policy_input.trim();
					//alert("policy_input___" + policy_input + "______");
					break;
				  case 5: // policy output
					policy_output = linealeida.substring(22, linealeida.length);
					policy_output = policy_output.trim();
					//alert("policy_output " + policy_output + linealeida);
					break;
				  case 6:
					VRF = linealeida.substring(5, linealeida.length);
					VRF = VRF.trim();
					//alert("VRF " + VRF);
					break;
				  case 7: // ipv4 address
					// ipv4 address 190.224.11.153 255.255.255.252
				    var lngPos2 = linealeida.indexOf(" ", 14); //cadena.indexOf(valorBusqueda[, indiceDesde])
					ip_WAN = linealeida.substring(14, lngPos2);
					ip_WAN = ip_WAN.trim();
					mascara_WAN = linealeida.substring(lngPos2, linealeida.length);
					mascara_WAN = mascara_WAN.trim();
					//MsgBox ip_WAN & " mascara: " & mascara_WAN
					//alert("ip_WAN " + ip_WAN + " mascara: " + mascara_WAN);
					break;
				  case 8: // dot1q vlan -> caso C2BELGRANO y esos
					//  dot1q vlan 2009 27
					CVLAN = linealeida.substring(11, linealeida.length);
					CVLAN = CVLAN.trim();
					
					/*			lngPos = Instr(CVLAN, " ")
								if lngPos > 0 then
									CVLAN = Mid(CVLAN, lngPos, Len(CVLAN))
									CVLAN = trim(CVLAN)
								end if*/
					alert("CVLAN " + CVLAN);
					break;
				  case 9: // Interface
					//Subint = linealeida.substring(9, linealeida.length);
					//Subint = Subint.trim();
					break;
				  case 10: // Description
					DESCRIPCION = linealeida.substring(12, linealeida.length);
					DESCRIPCION = DESCRIPCION.trim();
					break;
				  case 11: // encapsulation dot1q -> para caso MSBNG ejemplo RSC1MB
					// encapsulation dot1q 369 second-dot1q 3621
					lngPos2 = linealeida.indexOf("second-dot1q ");
					if (lngPos2 >= 0) { //encontró la VRF
						CVLAN = linealeida.substring(lngPos2 + 13, linealeida.length-1);
						//alert("CVLAN " + CVLAN); // esto esta perfecto
					}
					
					/*lngPos = 0;
					CVLAN = trim(CVLAN)
								lngPos = Instr(CVLAN, "second-dot1q")
								if lngPos > 0 then
									CVLAN = Mid(CVLAN, lngPos+13, Len(CVLAN))
									CVLAN = trim(CVLAN)
								end if
								'MsgBox CVLAN*/
					break;
			      case 12: //flow ipv4 monitor FMMAP sampler FNF_SAMPLER_MAP ingress
					//PEAKFLOW = "Si tiene" ' HABRIA que ver un caso de estos
					break;
				  default:
					// code block
				}
				
				
				lngPos = linealeida.indexOf("!");
				if (lngPos >= 0){ //encontró ! para salir
					Salir = 1;
					
				}
			}
		}
	}
	//termine de ver la configuracion actual de la subinterface
	
	if (ip_WAN != "NO TIENE"){
		NEXTHOP = SiguienteIP(ip_WAN);
		//alert("NEXTHOP " + NEXTHOP); 
	}
	
	//ahora tengo que ver la VRF;
	if (VRF.localeCompare("NO TIENE")!=0){
		miString = "	" ;
		agregar_fila(miString);
		elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + miString + "<br/>";
		
		miString = "	config VRF: sh run vrf " + VRF ;
		agregar_fila(miString);
		elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + miString + "<br/>";
		Salir = 0;
		for ( var i = 0; i < lines.length; i++)
		{
			linealeida = lines[i];
			lngPos = linealeida.indexOf(VRF);
			if (lngPos >= 0) { //encontró la VRF
				for ( i = i ; Salir != 1; i++){
					linealeida = lines[i];
					
					miString = linealeida ;
					agregar_fila(miString);
					
					elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	" +	miString ; //+ "<br/>";
					
					var lngPosRD = linealeida.indexOf("7303:");
					if (lngPosRD >= 0){ //encontró ! para salir
						RD = linealeida.substring(3, linealeida.length);
						RD = RD.trim();
						//alert("RD " + RD); 
					}
					
					lngPos = linealeida.indexOf("!");
					if (lngPos == 0){ //encontró ! para salir
						Salir = 1;
						
					}
				}
			}
		}
	}
	//termine de ver la VRF
	
	//ahora tengo que ver el BGP
	if (ip_WAN.localeCompare("NO TIENE")!=0){
		if (VRF.localeCompare("NO TIENE")!=0){
			
			miString = "	" ;
			agregar_fila(miString);	
			elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "<br/>";
			miString = "	Config Actual BGP: sh run router bgp 7303 vrf " + VRF + " neighbor " + NEXTHOP ;
			agregar_fila(miString);	
			elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + miString + "<br/>";
			
			var Inicio_bgp7303 = inicio_BLOQUE(lines, 0, -1, "router bgp 7303");
			var Fin_bgp7303 = fin_BLOQUE(lines, 0, -1, "router bgp 7303");
			var Inicio_bgp_VRF = inicio_BLOQUE(lines, Inicio_bgp7303, Fin_bgp7303, VRF);
			var Fin_bgp_VRF = fin_BLOQUE(lines, Inicio_bgp_VRF, Fin_bgp7303, VRF);
			var primeraParte_bgp_VRF = inicio_BLOQUE(lines, Inicio_bgp_VRF, Fin_bgp_VRF, "!");
			var Inicio_neighborBGP = inicio_BLOQUE(lines, primeraParte_bgp_VRF, Fin_bgp_VRF, NEXTHOP);
			//if (Inicio_neighborBGP == -1) alert("no encontro Inicio_neighborBGP " + NEXTHOP);
			var fin_neighborBGP = fin_BLOQUE(lines, Inicio_neighborBGP, Fin_bgp_VRF, NEXTHOP);
			//if (fin_neighborBGP == -1) alert("no encontro fin_neighborBGP " + NEXTHOP);
			
			print_BLOQUE(lines, Inicio_bgp7303, Inicio_bgp7303); // imprimo la primera linea del BGP 7303
			if (Inicio_bgp_VRF == -1) {
				miString = "	% No se encontro la VRF en BGP " +  VRF ;
				agregar_fila(miString);
				elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + miString + "<br/>";
			} else{
				print_BLOQUE(lines, Inicio_bgp_VRF, primeraParte_bgp_VRF); // esto esta perfecto!!
				BGP = "Si tiene";
			}
			
			
			if (Inicio_neighborBGP == -1) {
				miString = "	% No se encontro neighbor BGP " +  NEXTHOP ;
				agregar_fila(miString);
				elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + miString + "<br/>";
			} else {
				print_BLOQUE(lines, Inicio_neighborBGP, fin_neighborBGP);  // esto esta perfecto!!
				neighbor_BGP = "Si tiene";
				
				for ( i = Inicio_neighborBGP ; i < fin_neighborBGP; i++){
					linealeida = lines[i];
					
					var lngPosAS = linealeida.indexOf("remote-as");
					if (lngPosAS >= 0) { //encontró remote-as
						AS_actual = linealeida.substring(lngPosAS + 10, linealeida.length);
						AS_actual = AS_actual.trim();
						//alert("AS_actual " + AS_actual); 
					}
					
					var lngPosDesc = linealeida.indexOf("description ");
					if (lngPosDesc >= 0) { //encontró lngPosDesc BGP
						descripcion_BGP = linealeida.substring(lngPosAS + 16, linealeida.length);
						descripcion_BGP = descripcion_BGP.trim();
						//alert("descripcion_BGP " + descripcion_BGP); 
					}
					
					
				}
			
				
				
			}
		} else { // si no tiene VRF
			//sh run router bgp 7303  neighbor 181.15.37.98
			
			miString = "	" ;
			agregar_fila(miString);
			elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + miString + "<br/>";
			
			miString = "	Config Actual BGP: sh run router bgp 7303 neighbor " + NEXTHOP;
			agregar_fila(miString);
			elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + miString + "<br/>";
			
			var Inicio_bgp7303 = inicio_BLOQUE(lines, 0, -1, "router bgp 7303");
			var Fin_bgp7303 = fin_BLOQUE(lines, 0, -1, "router bgp 7303");
			var Inicio_neighborBGP = inicio_BLOQUE(lines, Inicio_bgp7303, Fin_bgp7303, NEXTHOP);
			var fin_neighborBGP = fin_BLOQUE(lines, Inicio_neighborBGP, Fin_bgp7303, NEXTHOP);
			//if (fin_neighborBGP == -1) alert("no encontro fin_neighborBGP " + NEXTHOP);
			
			print_BLOQUE(lines, Inicio_bgp7303, Inicio_bgp7303); // imprimo la primera linea del BGP 7303

			
			if (Inicio_neighborBGP == -1) {
				miString = "	% No se encontro neighbor BGP " +  NEXTHOP ;
				agregar_fila(miString);
				elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + miString + "<br/>";
				//elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	% No se encontro neighbor BGP " +  NEXTHOP + "<br/>";
			} else {
				print_BLOQUE(lines, Inicio_neighborBGP, fin_neighborBGP);  // esto esta perfecto!!
				neighbor_BGP = "Si tiene";
				BGP = "Si tiene";
				
				for ( i = Inicio_neighborBGP ; i < fin_neighborBGP; i++){
					linealeida = lines[i];
					
					var lngPosAS = linealeida.indexOf("remote-as");
					if (lngPosAS >= 0) { //encontró remote-as
						AS_actual = linealeida.substring(lngPosAS + 10, linealeida.length);
						AS_actual = AS_actual.trim();
						//alert("AS_actual " + AS_actual); 
					}
					
					var lngPosDesc = linealeida.indexOf("description ");
					if (lngPosDesc >= 0) { //encontró lngPosDesc BGP
						descripcion_BGP = linealeida.substring(lngPosAS + 16, linealeida.length);
						descripcion_BGP = descripcion_BGP.trim();
						//alert("descripcion_BGP " + descripcion_BGP); 
					}
					
				}	
			}
			
		}
	}
	//termine de ver BGP para el caso con VRF, habria que ver en las integras
	
	//AHORA TENGO QUE VER LAS ESTATICAS
	/*
	RP/0/RSP0/CPU0:BEL1MB#sh run router static vrf iafas-vpn | inc 192.168.89.30
	Mon Dec  9 15:47:57.440 ARG
	   190.227.201.108/32 TenGigE0/5/0/16.3513618 192.168.89.30
	   190.227.201.109/32 TenGigE0/5/0/16.3513618 192.168.89.30 */
	if (ip_WAN.localeCompare("NO TIENE")!=0){
		if (VRF.localeCompare("NO TIENE")!=0){
			elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "<br/>";
			
			elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	Config Actual Estaticas: sh run router static vrf " + VRF + " | inc " + linea_recortada + "<br/>";
	
			var Inicio_routerstatic = inicio_BLOQUE(lines, 0, -1, "router static");
			var Fin_routerstatic = fin_BLOQUE(lines, 0, -1, "router static");
			var Inicio_estaticasVRF = inicio_BLOQUE(lines, Inicio_routerstatic, Fin_routerstatic, VRF);
			var Fin_estaticasVRF = fin_BLOQUE(lines, Inicio_routerstatic, Fin_routerstatic, VRF);
			
			if (Inicio_estaticasVRF != -1 && Fin_estaticasVRF != -1) {
				print_BLOQUE_inc(lines, Inicio_estaticasVRF, Fin_estaticasVRF, linea_recortada);
				_static = "Si tiene";
			}
		} else{
			elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "<br/>";
			
			elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	Config Actual Estaticas: sh run router static | inc " + linea_recortada + "<br/>";
	
			var Inicio_routerstatic = inicio_BLOQUE(lines, 0, -1, "router static");
			var Fin_routerstatic = fin_BLOQUE(lines, 0, -1, "router static");
			
			if (Inicio_routerstatic != -1 && Fin_routerstatic != -1) {
				print_BLOQUE_inc(lines, Inicio_routerstatic, Fin_routerstatic, linea_recortada);
				_static = "Si tiene";
			}
		}
	}
	elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	Estaticas " + _static + "<br/>";
	//termine de VER LAS ESTATICAS
	
	//ahora tengo que ver los policy
	/*
	RP/0/RSP0/CPU0:BEL1MB#sh run policy-map VPN-650M
	Mon Dec  9 16:13:35.574 ARG
	policy-map VPN-650M
	 class class-default
	  police rate 650000000 bps burst 112500000 bytes peak-burst 225000000 bytes 
	   conform-action transmit
	   exceed-action drop
	  ! 
	 ! 
	 end-policy-map
	!
	*/
	
	if (policy_input.localeCompare("NO TIENE")!=0){
		elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "<br/>";
		elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	Config Actual Policy input: sh run policy-map " + policy_input + "<br/>";
		
		var Inicio_policy_input = inicio_BLOQUE(lines, 0, -1, "policy-map " + policy_input);
		var Fin_policy_input = fin_BLOQUE(lines, Inicio_policy_input-1, Inicio_policy_input+20, "policy-map " + policy_input);
		
		if (Inicio_policy_input != -1 && Fin_policy_input != -1) {
			print_BLOQUE(lines, Inicio_policy_input, Fin_policy_input); 
			//alert("Inicio_policy_input " + Inicio_policy_input + " Fin_policy_input " + Fin_policy_input); 
		}
	} //else alert("el servicio no tiene policy_input"); 

	
	if (policy_output.localeCompare("NO TIENE")!=0){
		elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "<br/>";
		elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	Config Actual Policy output: sh run policy-map " + policy_output + "<br/>";
		
		var Inicio_policy_output = inicio_BLOQUE(lines, 0, -1, "policy-map " + policy_output);
		var Fin_policy_output = fin_BLOQUE(lines, 0, -1, "policy-map " + policy_output);
		
		if (Inicio_policy_output != -1 && Fin_policy_output != -1) {
			print_BLOQUE(lines, Inicio_policy_output, Fin_policy_output);  
		}
	} //else alert("el servicio no tiene policy_output"); 
	
	
	elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	--------------------------<br/>";
	elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	" + "Lista de comandos completa: " + "<br/>";
	elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	" + "sh int " + SUBINTERFACE + "<br/>"; 
	elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	" + "sh run int " + SUBINTERFACE + "<br/>";
	elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	" + "sh run vrf " + VRF + "<br/>"; 
	elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	" + "sh run router bgp 7303 vrf " + VRF + " neighbor " + NEXTHOP + "<br/>"; 
	elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	" + "sh bgp vrf " + VRF + " neighbors " + NEXTHOP + " routes " + "<br/>"; 
	elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	" + "sh arp vrf " + VRF + " " + SUBINTERFACE + "<br/>"; 
	elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	" + "sh bgp vrf " + VRF + " summary | inc " + NEXTHOP + "<br/>"; 
	elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	" + "sh run router static vrf " + VRF + " | inc " + linea_recortada + "<br/>"; // SUBINTERFACE
	elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	" + "sh run policy-map " + policy_input + "<br/>"; 
	elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	" + "sh run policy-map " + policy_output + "<br/>"; 
	elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	" + "ping vrf " + VRF  + " " + NEXTHOP + "<br/>";
	
	for ( i = 0; i < Cant_estaticas; i++)
	{
		elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	" + "ping vrf " + VRF  + " " + Estatica(i) + "<br/>"; // esta mal porque muestra la "/"
	}
	
	
	elemento.innerHTML +=    "---------------------------------------" + "<br/>";
	elemento.innerHTML +=    "DATOS Configuracion Actual: " + "<br/>";
	elemento.innerHTML +=    "EQUIPO:	" + EQUIPO + "<br/>";
	elemento.innerHTML +=    "Subinterface:	" + SUBINTERFACE + "<br/>";
	elemento.innerHTML +=    "Descripcion:	" + DESCRIPCION + "<br/>";
	elemento.innerHTML +=    "Bandwidth:	" + BANDWIDTH + "<br/>";
	elemento.innerHTML +=    "Policy input:	" + policy_input + "<br/>";
	elemento.innerHTML +=    "Policy output:	" + policy_output + "<br/>";
	elemento.innerHTML +=    "VRF:	" + VRF + "<br/>";
	elemento.innerHTML +=    "Ip de WAN:	" + ip_WAN + "<br/>";
	elemento.innerHTML +=    "Mascara WAN:	" + mascara_WAN + "<br/>";
	elemento.innerHTML +=    "WAN+1:	" + NEXTHOP + "<br/>";
	elemento.innerHTML +=    "CVLAN:	" + CVLAN + "<br/>";
	elemento.innerHTML +=    "RD:	" + RD + "<br/>";
	elemento.innerHTML +=    "BGP:	" + BGP + "<br/>";
	elemento.innerHTML +=    "BGP neighbor:	" + neighbor_BGP + "<br/>";
	elemento.innerHTML +=    "AS:	" + AS_actual + "<br/>";
	elemento.innerHTML +=    "Peakflow:	" + PEAKFLOW + " ESTO TODAVIA NO ESTA PROBADO!" + "<br/>"; //peakflow habria que buscar un caso para armarlo TODAVIA NO ANDA
	elemento.innerHTML +=    "Descripcion BGP:	" + descripcion_BGP + "<br/>";
	
	/*
	
	
	do while indice < Cant_estaticas
				file.Write   "Estatica " & indice & " :	" & Estatica(indice) & vbCrLf
				if SalidaINTERNET(indice) <> "NO TIENE" then
					file.Write   "Salida a INTERNET:	" & SalidaINTERNET(indice) & vbCrLf
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
	*/
	
	
	elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	--------------------------<br/>";
	elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	--------------------------<br/>";
	
}



function print_BLOQUE_inc(linea, inicio, fin, textoBuscado){
	var i, fin_fin, lngPos, coincidencia;
	var elemento = document.getElementById('contenido-archivo');
	//elemento.innerHTML = EQUIPO + " " + SUBINTERFACE + "	" +	linealeida + "<br/>";
	
	lngPos = -1;
	coincidencia = 0;
	
	if (inicio == -1){
		elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	% No such configuration item(s)"; // + "<br/>";
		return;
	}
	
	if (fin == -1){
		fin_fin = linea.length;
	}else{
		fin_fin = fin;
	}
	
	if (inicio > fin){
		//alert("print_BLOQUE: todo mal cuando quiero imprimir el inicio es mas grande que el fin");
		return -1;
	}
	
	
	for (i = inicio; i <= fin_fin; i++){
		linealeida = linea[i];
		lngPos = linealeida.indexOf(textoBuscado);
		if (lngPos > 0) {
			elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	" +	linealeida; // + "<br/>";
			coincidencia = 1;
		}
	}
	
	if (coincidencia == 0) elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	% No se encontro " + textoBuscado + "<br/>";
	
}

function fin_BLOQUE(linea, inicio, fin, textoBuscado){
	
	var re =  /(\s+)/;  //regular expression para los espacios al inicio
	
	var StringsArray, encontrado;
	var i, fin_fin, sin_espacios, cant_espacios, lngPos, lngPosfin, largo, inicio_Bloque;
	//var ultima_linea_del_bloque = imprimir_BLOQUE(linea, 0, -1, "router bgp 7303")
	
	var flag = false;
	var exclamacion = 0;
	
	if (inicio == -1){
		return -1; // si devuelvo -1 quiere decir que no encontré nada
	}
	
	
	if (fin == -1){
		fin_fin = linea.length;
	}else{
		fin_fin = fin;
	}
	
	
	// tengo que buscar "!"
	
	for (i = inicio; i < fin_fin; i++){
		linealeida = linea[i];
		lngPos = linealeida.indexOf(textoBuscado);
		
		if (lngPos >= 0){
			inicio_Bloque = i;	// aca encontré el inicio del bloque
			//debo contar la cantidad de espacios antes
			sin_espacios = linealeida.trim();
			cant_espacios = linealeida.length - sin_espacios.length;
			//alert(i + " linealeida " + linealeida);
		}
	}
	
	for (i = inicio_Bloque; i < fin_fin; i++){
		linealeida = linea[i];
		lngPosfin = linealeida.indexOf("!");
		
		//if (lngPos <= cant_espacios){
		if ( lngPosfin >= 0){
			if ( lngPosfin < cant_espacios){
				fin_bloque = i;	
				//alert("cant_espacios " + cant_espacios + " fin_BLOQUE " + i + " " + linealeida + " " + lngPosfin);
				return fin_bloque;
			}
		}
	}
	
	//alert(textoBuscado + " fin_BLOQUE: error -1");
	return -1; // si devuelvo -1 quiere decir que no encontré nada
}

function inicio_BLOQUE(linea, inicio, fin, textoBuscado){
	
	var StringsArray, encontrado;
	var i, fin_fin, sin_espacios, cant_espacios, lngPos, largo;
	//var ultima_linea_del_bloque = imprimir_BLOQUE(linea, 0, -1, "router bgp 7303")
	
	var flag = false;
	var exclamacion = 0;
	
	if (fin == -1){
		fin_fin = linea.length;
	}else{
		fin_fin = fin;
		if (inicio > fin){
		alert("inicio_BLOQUE: todo mal cuando quiero imprimir el inicio es mas grande que el fin");
		return -1;
		}
	}
	
	
	
	StringsArray = [textoBuscado, "!"];
	//encontrado = encontrarString(linealeida, StringsArray);
	
	//alert("inicio " + inicio + " fin " + fin_fin);
	for (i = inicio; i < fin_fin; i++){
		linealeida = linea[i];
		lngPos = linealeida.indexOf(textoBuscado);
		
		if (lngPos >= 0){
			return i;
		}

	}
	
	
	//alert(textoBuscado + " inicio_BLOQUE: error -1");
	return -1; // si devuelvo -1 quiere decir que no encontré nada
}
	

function encontrarString(linealeida, StringsArray){
	var i, largo;
	//esta funcion tiene problemas con el primero de la lista
	
	largo = StringsArray.length;
	
	//alert("estoy adentro de encontrarString " + linealeida + " " + StringsArray[0]); 
	
	for (i = 0; i < largo; i++)
	{
		lngPos = linealeida.indexOf(StringsArray[i]);
		//alert(linealeida + " buscado " + StringsArray[i]);
		if (lngPos >= 0){ //encontró la palabra buscada
			return i + 1;
		} 
	}
	
	if (i = largo){ //esto quiere decir que llegue al final del array y no encontré lo que queria
		return -1;
	}
}

function SiguienteIP(IP){
	var ultocteto;
	
	var re =  /(\d+.\d+.\d+.)(\d+)/;
	var primeros;
	var largo;
	var siguiente;
	
	
	ultocteto = IP.replace(re, "$2");
	siguiente = IP.replace(re, "$1") + (parseInt(ultocteto) + 1);
	// alert(siguiente); 
	

	return siguiente;
}


function print_BLOQUE(linea, inicio, fin){
	var i, fin_fin;
	var elemento = document.getElementById('contenido-archivo');
	//elemento.innerHTML = EQUIPO + " " + SUBINTERFACE + "	" +	linealeida + "<br/>";
	
	if (inicio == -1){
		elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	% No such configuration item(s)"; // + "<br/>";
		return;
	}
	
	if (fin == -1){
		fin_fin = linea.length;
	}else{
		fin_fin = fin;
	}
	
	if (inicio > fin){
		//alert("print_BLOQUE: todo mal cuando quiero imprimir el inicio es mas grande que el fin");
		return -1;
	}
	
	
	for (i = inicio; i <= fin_fin; i++){
		linealeida = linea[i];
		
		elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	" +	linealeida; // + "<br/>";
		
	}
}

function imprimir_BLOQUE(linea, inicio, fin, textoBuscado){
	
	var StringsArray, encontrado;
	var i, fin_fin, sin_espacios, cant_espacios, lngPos, largo;
	//var ultima_linea_del_bloque = imprimir_BLOQUE(linea, 0, -1, "router bgp 7303")
	
	var flag = false;
	var exclamacion = 0;
	
	if (fin == -1){
		fin_fin = linea.length;
	}else{
		fin_fin = fin;
	}
	
	if (inicio > fin){
		alert("imprimir_BLOQUE: todo mal cuando quiero imprimir el inicio es mas grande que el fin");
		return -1;
	}
	
	elemento.innerHTML += EQUIPO + " " + SUBINTERFACE + "	" + linealeida + "<br/>";
	
	
	StringsArray = [textoBuscado, "!"];
	//encontrado = encontrarString(linealeida, StringsArray);
	
	for (var i = inicio; i < fin_fin; i++){
		linealeida = linea[i];
		lngPos = linealeida.indexOf(textoBuscado);
		
		if (lngPos >= 0){
			flag = true
			//alert("estoy adentro de imprimir_BLOQUE " + linealeida); 
			//MsgBox linealeida
			//dato importante en este momento en la variable i se encuentra el num de linea donde comienza en bloque
			//aca deberia contar la cantidad de espacios en blanco antes que arranque el bloque
			largo = linealeida.length;
			sin_espacios = linealeida.trim().length;
			cant_espacios = largo - sin_espacios;
			//alert(cant_espacios);
		}
		
		if (flag == true){
			//alert(linealeida);
		}
		
		exclamacion = linealeida.indexOf("!");
		if (exclamacion == (cant_espacios + 1) && flag == true){
			//alert(linealeida);
			flag = false; //dato importante en este momento en la variable i se encuentra el num de linea donde termina en bloque
			return i;
		}
	}
	
	
	//alert(linealeida);
	return 0;
}


//------------------------------- exportar a excel

function fnExcelReport()
{
    var tab_text="<table border='2px'><tr bgcolor='#87AFC6'>";
    var textRange; var j=0;
    //tab = document.getElementById('headerTable'); // id of table
	tab = document.getElementById('testTable'); // id of table
	

    for(j = 0 ; j < tab.rows.length ; j++) 
    {     
        tab_text=tab_text+tab.rows[j].innerHTML+"</tr>";
        //tab_text=tab_text+"</tr>";
    }

    tab_text=tab_text+"</table>";
    tab_text= tab_text.replace(/<A[^>]*>|<\/A>/g, "");//remove if u want links in your table
    tab_text= tab_text.replace(/<img[^>]*>/gi,""); // remove if u want images in your table
    tab_text= tab_text.replace(/<input[^>]*>|<\/input>/gi, ""); // reomves input params

    var ua = window.navigator.userAgent;
    var msie = ua.indexOf("MSIE "); 

    if (msie > 0 || !!navigator.userAgent.match(/Trident.*rv\:11\./))      // If Internet Explorer
    {
        txtArea1.document.open("txt/html","replace");
        txtArea1.document.write(tab_text);
        txtArea1.document.close();
        txtArea1.focus(); 
        sa=txtArea1.document.execCommand("SaveAs",true,"Say Thanks to Sumit.xls");
    }  
    else                 //other browser not tested on IE 11
        sa = window.open('data:application/vnd.ms-excel,' + encodeURIComponent(tab_text));  

    return (sa);
}

//--------------------------------------arranco con la parte de tablas



function agregar_fila(texto_en_celda) {
	document.getElementById('mitabla').style.display = 'block';
	SUBINTERFACES = document.getElementById("subint").value;
	Subint = SUBINTERFACES.split('\n');    // ok! tengo todas las subinterfaces en un array
	Cant_subint = Subint.length;
	var columnas = 2;
	var filas;
	
	//var tableRef = document.getElementById('mitabla').getElementsByTagName('tbody')[0];
	
	//for(var j = 0; j < Cant_subint; j++) { //esto es para las filas
		//inserto una fila en la tabla en la ultima fila
		var newRow   = tableRef.insertRow();

		for(var i = 0; i < columnas ; i++) { //esto es para las columnas
			//inserto una celda en la fila en el indice 0
			var newCell  = newRow.insertCell(i);
			var newText;
			
			if (i == 0){
				newText = document.createTextNode(EQUIPO + " " + SUBINTERFACE + "	");
			}else{
				//newText = document.createTextNode("aca va la bajada del gestor");
				newText = document.createTextNode(texto_en_celda);
			}
			newCell.appendChild(newText);
		}
	//}
	
	//alert("hasta aca esta bien " + Cant_subint);
	
	return;
}

function start() {
	var columnas;
	var filas;
    // get the reference for the body //toma referencia del body
    var mybody = document.getElementsByTagName("body")[0];

    // creates <table> and <tbody> elements
    mytable     = document.createElement("table");
    mytablebody = document.createElement("tbody");

    // creating all cells
    for(var j = 0; j < 2; j++) {
    // creates a <tr> element
    mycurrent_row = document.createElement("tr");

    for(var i = 0; i < 2; i++) {
        // creates a <td> element
        mycurrent_cell = document.createElement("td");
        // creates a Text Node
		currenttext = document.createTextNode("fila " + j + ", columna " + i);
        // appends the Text Node we created into the cell <td>
                mycurrent_cell.appendChild(currenttext);
                // appends the cell <td> into the row <tr>
                mycurrent_row.appendChild(mycurrent_cell);
            }
            // appends the row <tr> into <tbody>
            mytablebody.appendChild(mycurrent_row);
        }

        // appends <tbody> into <table>
        mytable.appendChild(mytablebody);
        // appends <table> into <body>
        mybody.appendChild(mytable);
        // sets the border attribute of mytable to 2;
        //mytable.setAttribute("border","1");
}

function comprobar(){
	// row es la fila
	// cells es la columna
	var mostrar = document.getElementById('mitabla').rows[3].cells[1].innerText ;
	//var mostrar = "algo para mostrar";
	
	
	alert("hasta aca esta bien " + mostrar);
}