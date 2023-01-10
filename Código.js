const archivoOrdenes = SpreadsheetApp.getActiveSpreadsheet();
const hojaPrincipal = archivoOrdenes.getSheetByName("Base Principal");
const datosPrincipal = hojaPrincipal.getDataRange().getDisplayValues();
const hojaTerminados = archivoOrdenes.getSheetByName("Log de Terminados");
const hojaPruebas = archivoOrdenes.getSheetByName("Pruebas");

function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu("Gestionar órdenes").addItem("Abrir módulo","cargarSideBar").addToUi()

};

// Application settings
const CSV_HEADER_EXIST = true;  // Set to true if CSV files have a header row, false if not.

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
};

function cargarSideBar(){
  const html = HtmlService.createTemplateFromFile("UI").evaluate();
  html.setTitle("Payments Monitor Bot")
  SpreadsheetApp.getUi().showSidebar(html)

};

// INICIO CSV órdenes pendientes nuevas

function cargarCSVNuevos(obj){
  const blob = Utilities.newBlob(Utilities.base64Decode(obj.data),obj.mimeType,obj.fileName);
  const id = '13fy-iESZl51U5XHNYvcgg3QWWpXxqQGi';
  const folder = DriveApp.getFolderById(id);
  const file = folder.createFile(blob);
  const fileURL = file.getUrl();
  const response = {
    'fileName' : obj.fileName,
    'url' : fileURL,
    'status' :true,
    'data' : JSON.stringify(obj)
  };

  // let archivoPrincipal = SpreadsheetApp.getActiveSpreadsheet();
  // let hojaPrueba = archivoPrincipal.getSheetByName("Pruebas");
   procesarCSVNuevos(file);


  // Para borrar el archivo una vez tomada la información
  // file.setTrashed(true);

 
  return response;
};


/**
 * Parses CSV data into an array and appends it after the last row in the destination spreadsheet.
 * 
 * @return {boolean} true if the update is successful, false if unexpected errors occur.
*/
function procesarCSVNuevos(csvFile) {

  try {
    // Gets the sheet of the destination spreadsheet.
    // let sheet = hojaDestino;

    // Parses CSV file into data array.
    let data = Utilities.parseCsv(csvFile.getBlob().getDataAsString());

    // Omits header row if application variable CSV_HEADER_EXIST is set to 'true'.
    // if (CSV_HEADER_EXIST) {
      // data.splice(0, 1);
    //   // data.shift();
    // }

    //Se deben quitar dos filas iniciales en base historica
    data.splice(0, 2);

    // Logger.log(data)
    // // Gets the row and column coordinates for next available range in the spreadsheet. 
    // let startRow = sheet.getLastRow() + 1;
    // let startCol = 1;
    // // Determines the incoming data size.
    // let numRows = data.length;
    // let numColumns = data[0].length;
    

    let registrosActuales = datosPrincipal.map(dato=>{let llave = String(dato[0])+String(dato[3]);return llave});
    // Logger.log(registrosActuales);
    

    data.map((dato)=>{
      let llave = String(dato[0])+String(dato[2]);
      let existeRegistro = registrosActuales.includes(llave);
      Logger.log(llave);
      Logger.log(existeRegistro);
      let numCompania=dato[0];
      let fechaCreacionOrden=dato[1];
      let numOrdenPago=dato[2];
      let tipoOrdenPago=dato[3];

      let valorOrden=dato[5].replace(".",",");
      let estadoOrden=dato[6];
      let estadoPago=dato[8];
      let causal=dato[9];
      let formaPago=dato[10];
      let tipoIdentificacionTercero=dato[11];
      let numIdentificacionTercero=dato[12];
      let codigoBancoDestino=dato[13];
      let estadoCuentaTercero=dato[15];
      let fechaPagoProgramada=dato[16];
      if(!existeRegistro){
        // Logger.log(existeRegistro);
        hojaPrincipal.appendRow([numCompania,"",fechaCreacionOrden,numOrdenPago,tipoOrdenPago,valorOrden,estadoOrden,estadoPago,causal,formaPago,tipoIdentificacionTercero,numIdentificacionTercero,codigoBancoDestino,estadoCuentaTercero,fechaPagoProgramada]);
      };
    });

    // Appends data into the sheet.
    // sheet.getRange(startRow, startCol, numRows, numColumns).setValues(data);
    Browser.msgBox("Se cargó el archivo");
    return true; // Success.

  } catch (e) {
    Browser.msgBox("No se pudo cargar el archivo   /" + e);
    return false; // Failure. Checks for CSV data file error.
  };
};

// FIN CSV órdenes pendientes nuevas




// INICIO CSV estado pagos

// function cargarCSVActualizaEstado(obj){
//   const blob = Utilities.newBlob(Utilities.base64Decode(obj.data),obj.mimeType,obj.fileName);
//   const id = '13fy-iESZl51U5XHNYvcgg3QWWpXxqQGi';
//   const folder = DriveApp.getFolderById(id);
//   const file = folder.createFile(blob);
//   const fileURL = file.getUrl();
//   const response = {
//     'fileName' : obj.fileName,
//     'url' : fileURL,
//     'status' :true,
//     'data' : JSON.stringify(obj)
//   };

//   // let archivoPrincipal = SpreadsheetApp.getActiveSpreadsheet();
//   // let hojaPrueba = archivoPrincipal.getSheetByName("Pruebas");
//    procesarCSVActualizaEstado(file);


//   // Para borrar el archivo una vez tomada la información
//   // file.setTrashed(true);

 
//   return response;
// };

// Versión inicial con estructura separada por numero de columnas en archivo
/**
* Parses CSV data into an array and appends it after the last row in the destination spreadsheet.
* 
* @return {boolean} true if the update is successful, false if unexpected errors occur.
*/


// function procesarCSVActualizaEstado(csvFile) {

//   // Parses CSV file into data array.
//   let data = Utilities.parseCsv(csvFile.getBlob().getDataAsString());

  
//   // Si el archivo trae 16 columnas---------------------
  
//   if(data[1].length === 16){  
//     try {
//       // Gets the sheet of the destination spreadsheet.
//       // let sheet = hojaDestino;

//       let datosPrincipal = hojaPrincipal.getDataRange().getValues(); 

//       data.map((dato)=>{
      
//         let banco=dato[0];
//         let numIdentificacionTercero=dato[1];
//         let numOrdenPago= String(dato[2]);
//         // let numAutorizacion=dato[3];
//         let valorOrden=dato[4].replace(".",",");
//         // let numCuentaDestino=dato[5];
//         // let codBancoDestino=dato[6];
//         let codCompania=dato[7];
//         let fechaPago=dato[8];
//         // let fechaCreacion=dato[9];
//         let fechaEnvioCentralizador=dato[10];
//         // let horaAplicacion=dato[11];
//         let tipoProceso=dato[12];
//         let estado=dato[13];
//         let codMotivoRechazo=dato[14];
//         // let gestionador=dato[15];

//         // let arrayDatosCruce = Array.from([banco,fechaEnvioCentralizador,numIdentificacionTercero,fechaPago,tipoProceso]);

//         let posicionArray = datosPrincipal.findIndex(dato=>dato[3]==numOrdenPago);

//         let existeEnPrincipal = posicionArray!= -1 ? true : false ;

//         if(existeEnPrincipal){

//           let numFilaCambio = posicionArray+1
//           hojaPrincipal.getRange(numFilaCambio,16).setValue(banco);
//           hojaPrincipal.getRange(numFilaCambio,17).setValue(fechaEnvioCentralizador);
//           hojaPrincipal.getRange(numFilaCambio,18).setValue(numIdentificacionTercero);
//           hojaPrincipal.getRange(numFilaCambio,19).setValue(fechaPago);
//           hojaPrincipal.getRange(numFilaCambio,20).setValue(tipoProceso);

          
//           switch (estado) {
//             case 'RECHAZADO':
//               hojaPrincipal.getRange(numFilaCambio,8).setValue("Rechazado").setBackground("red");
//               hojaPrincipal.getRange(numFilaCambio,9).setValue(codMotivoRechazo).setBackground("orange");;


              
//               break;
//             case 'TERMINADO':

//               // Crear log corto de Terminado y borrar la fila cuando terminado
//               let logTerminado= [fechaPago,codCompania,numOrdenPago,numIdentificacionTercero,valorOrden];
//               hojaTerminados.appendRow(logTerminado)
//               hojaPrincipal.deleteRow(numFilaCambio); 

              

//               // Cambiar el estado de la orden de pago
//               // hojaPrincipal.getRange(numFilaCambio,8).setValue("Terminado").setBackground("green");
              


//               break;
//             // case 'RECHAZADO':
//             //   console.log('Oranges are $0.59 a pound.');
//             //   break;
//             default:
//               console.log(`${estado} no es un estado válido.`);
//           }
//         }   
//       });

//       Browser.msgBox("Información actualizada");
//       return true; // Success.

//     } catch(e) {
//       Browser.msgBox("No se pudo cargar el archivo  /"+ e);
//       return false; // Failure. Checks for CSV data file error.
//     };
//   }else{


//       // Si el archivo trae 17 columnas---------------------

//       try {
//       // Gets the sheet of the destination spreadsheet.
//       // let sheet = hojaDestino;
//       // let reenvios = [];
//       let datosPrincipal = hojaPrincipal.getDataRange().getValues();

      
//       data.map((dato)=>{
      
//         let banco=dato[0];
//         let numIdentificacionTercero=dato[1];
//         let numOrdenPago= String(dato[2]);
//         // let numAutorizacion=dato[3];
//         let valorOrden=dato[4].replace(".",",");
//         // let numCuentaDestino=dato[5];
//         // let codBancoDestino=dato[6];
//         let codCompania=dato[7];
//         let fechaPago=dato[8];
//         // let fechaCreacion=dato[9];
//         let fechaEnvioCentralizador=dato[10];
//         // let horaAplicacion=dato[11];
//         let tipoProceso=dato[12];
//         let estado=dato[13];
//         let codMotivoRechazo=dato[14];
//         // let gestionador=dato[15];
//         let reenvio=dato[15];  
//          if(reenvio == "SI") {

//            let arrayLogReenvio = [fechaPago,numOrdenPago]
          
//            hojaReenvios.appendRow(arrayLogReenvio)
//           //  reenvios.push(numOrdenPago);
//          }  
        
//         let posicionArray = datosPrincipal.findIndex(dato=>dato[3]==numOrdenPago);

//         let existeEnPrincipal = posicionArray!= -1 ? true : false ;

//         if(existeEnPrincipal){

//           let numFilaCambio = posicionArray+1
//           hojaPrincipal.getRange(numFilaCambio,16).setValue(banco);
//           hojaPrincipal.getRange(numFilaCambio,17).setValue(fechaEnvioCentralizador);
//           hojaPrincipal.getRange(numFilaCambio,18).setValue(numIdentificacionTercero);
//           hojaPrincipal.getRange(numFilaCambio,19).setValue(fechaPago);
//           hojaPrincipal.getRange(numFilaCambio,20).setValue(tipoProceso);
          
//           switch (estado) {
//             case 'RECHAZADO':
//               hojaPrincipal.getRange(numFilaCambio,8).setValue("Rechazado").setBackground("red");
//               hojaPrincipal.getRange(numFilaCambio,9).setValue(codMotivoRechazo).setBackground("orange");
              
//               break;
//             case 'TERMINADO':

//               // Crear log corto de Terminado y borrar la fila cuando terminado
//               let logTerminado= [fechaPago,codCompania,numOrdenPago,numIdentificacionTercero,valorOrden];
//               hojaTerminados.appendRow(logTerminado)
//               hojaPrincipal.deleteRow(numFilaCambio);               

//               // Cambiar el estado de la orden de pago
//               // hojaPrincipal.getRange(numFilaCambio,8).setValue("Terminado").setBackground("green");         

//               break;
//             // case 'RECHAZADO':
//             //   console.log('Oranges are $0.59 a pound.');
//             //   break;
//             default:
//               console.log(`${estado} no es un estado válido.`);



             
//           }
//         }   
//       });

      
//       // Logger.log("llego")
//       // const cantidadReenvios = reenvios.reduce((contadorReenvio,reenvio)=>{
//       //   contadorReenvio[reenvio] = (contadorReenvio[reenvio]||0)+1;
//       //   return cantidadReenvios
//       // },{});

//       // const cantidadReenvios = reenvios.reduce((contadorReenvio,reenvio)=>{(
//       //   contadorReenvio[reenvio] ? contadorReenvio[reenvio] += 1: contadorReenvio[reenvio]=1,contadorReenvio)
//       //   // return cantidadReenvios
//       // },{});

//       // Logger.log(cantidadReenvios);

      
//       // for (let property in cantidadReenvios) {
//       //   console.log(`${property}: ${cantidadReenvios[property]}`);

//       //   let hoy = Utilities.formatDate(new Date(), "Bogota/America", "dd/MM/yyyy");

//       //   let arrayLogRechazada = [hoy,property,cantidadReenvios[property]];
//       //   hojaReenvios.appendRow(arrayLogRechazada)        
//       // }

//       Browser.msgBox("Información actualizada");
//       return true; // Success.

//     } catch(e) {
//       Browser.msgBox("No se pudo cargar el archivo  /"+ e);
//       return false; // Failure. Checks for CSV data file error.
//     };
//   }
// };

// FIN CSV estado pagos


// Versión Refactor 1 ---------------------------------

// function procesarCSVActualizaEstado(csvFile) {

//   // Parses CSV file into data array.
//   let data = Utilities.parseCsv(csvFile.getBlob().getDataAsString());

  
//   // Si el archivo trae 16 columnas---------------------
  
//   if(data[1].length === 16){  
//     try {

//       let datosPrincipal = hojaPrincipal.getDataRange().getValues(); 

//       data.map((dato)=>{
      
//         let banco=dato[0];
//         let numIdentificacionTercero=dato[1];
//         let numOrdenPago= String(dato[2]);
//         // let numAutorizacion=dato[3];
//         let valorOrden=dato[4].replace(".",",");
//         // let numCuentaDestino=dato[5];
//         // let codBancoDestino=dato[6];
//         let codCompania=dato[7];
//         let fechaPago=dato[8];
//         // let fechaCreacion=dato[9];
//         let fechaEnvioCentralizador=dato[10];
//         // let horaAplicacion=dato[11];
//         let tipoProceso=dato[12];
//         let estado=dato[13];
//         let codMotivoRechazo=dato[14];
//         // let gestionador=dato[15];

//         let posicionArray = datosPrincipal.findIndex(dato=>dato[3]==numOrdenPago);

//         let existeEnPrincipal = posicionArray!= -1 ? true : false ;

//         if(existeEnPrincipal){

//           let numFilaCambio = posicionArray+1
//           hojaPrincipal.getRange(numFilaCambio,16).setValue(banco);
//           hojaPrincipal.getRange(numFilaCambio,17).setValue(fechaEnvioCentralizador);
//           hojaPrincipal.getRange(numFilaCambio,18).setValue(numIdentificacionTercero);
//           hojaPrincipal.getRange(numFilaCambio,19).setValue(fechaPago);
//           hojaPrincipal.getRange(numFilaCambio,20).setValue(tipoProceso);

//           switch (estado) {
//             case 'RECHAZADO':
//               hojaPrincipal.getRange(numFilaCambio,8).setValue("Rechazado").setBackground("red");
//               hojaPrincipal.getRange(numFilaCambio,9).setValue(codMotivoRechazo).setBackground("orange");;

//               break;
//             case 'TERMINADO':

//               // Crear log corto de Terminado y borrar la fila cuando terminado
//               let logTerminado= [fechaPago,codCompania,numOrdenPago,numIdentificacionTercero,valorOrden];
//               hojaTerminados.appendRow(logTerminado)
//               hojaPrincipal.deleteRow(numFilaCambio); 

//               break;

//             default:
//             break
//           }
//         }   
//       });
//       Browser.msgBox("Información actualizada");
//       return true; // Success.
//     } catch(e) {
//       Browser.msgBox("No se pudo cargar el archivo  /"+ e);
//       return false; // Failure. Checks for CSV data file error.
//     };
//   }else{

//       // Si el archivo trae 17 columnas---------------------

//       try {

//       let datosPrincipal = hojaPrincipal.getDataRange().getValues();      
//       data.map((dato)=>{
      
//         let banco=dato[0];
//         let numIdentificacionTercero=dato[1];
//         let numOrdenPago= String(dato[2]);
//         // let numAutorizacion=dato[3];
//         let valorOrden=dato[4].replace(".",",");
//         // let numCuentaDestino=dato[5];
//         // let codBancoDestino=dato[6];
//         let codCompania=dato[7];
//         let fechaPago=dato[8];
//         // let fechaCreacion=dato[9];
//         let fechaEnvioCentralizador=dato[10];
//         // let horaAplicacion=dato[11];
//         let tipoProceso=dato[12];
//         let estado=dato[13];
//         let codMotivoRechazo=dato[14];
//         // let gestionador=dato[15];
//         let reenvio=dato[15];  
//          if(reenvio == "SI") {

//            let arrayLogReenvio = [fechaPago,numOrdenPago]
          
//            hojaReenvios.appendRow(arrayLogReenvio);          
//          }  
        
//         let posicionArray = datosPrincipal.findIndex(dato=>dato[3]==numOrdenPago);
//         let existeEnPrincipal = posicionArray!= -1 ? true : false ;

//         if(existeEnPrincipal){

//           let numFilaCambio = posicionArray+1
//           hojaPrincipal.getRange(numFilaCambio,16).setValue(banco);
//           hojaPrincipal.getRange(numFilaCambio,17).setValue(fechaEnvioCentralizador);
//           hojaPrincipal.getRange(numFilaCambio,18).setValue(numIdentificacionTercero);
//           hojaPrincipal.getRange(numFilaCambio,19).setValue(fechaPago);
//           hojaPrincipal.getRange(numFilaCambio,20).setValue(tipoProceso);
          
//           switch (estado) {
//             case 'RECHAZADO':
//               hojaPrincipal.getRange(numFilaCambio,8).setValue("Rechazado").setBackground("red");
//               hojaPrincipal.getRange(numFilaCambio,9).setValue(codMotivoRechazo).setBackground("orange");
              
//               break;
//             case 'TERMINADO':

//               // Crear log corto de Terminado y borrar la fila cuando terminado
//               let logTerminado= [fechaPago,codCompania,numOrdenPago,numIdentificacionTercero,valorOrden];
//               hojaTerminados.appendRow(logTerminado)
//               hojaPrincipal.deleteRow(numFilaCambio);       
 
//               break;
//             default:
//             break            
//           }
//         }   
//       });
//       Browser.msgBox("Información actualizada");
//       return true; // Success.
//     } catch(e) {
//       Browser.msgBox("No se pudo cargar el archivo  /"+ e);
//       return false; // Failure. Checks for CSV data file error.
//     };
//   }
// };

// Versión Refactor 2 ---------------------------------


// var objetoBaseEstadoRechazado = [];
// var objetoBaseEstadoTerminado = [];


// function procesarCSVActualizaEstado(csvFile) {

//   try {
//     // Parses CSV file into data array.
//     objetoBaseEstadoRechazado = [];
//     objetoBaseEstadoTerminado = [];

//     // let csvFile = DriveApp.getFileById("1e1tMZNgshANqHjBec25RtLD6zyQISReY")
//     let data = Utilities.parseCsv(csvFile.getBlob().getDataAsString());

//     data.map(dato => {      
//         switch (dato [13]) {
//           case 'RECHAZADO':
//             // Logger.log(dato[0]);
//             objetoBaseEstadoRechazado.push({        
//               banco: dato[0],
//               numIdentificacionTercero: String(dato[1]),
//               numOrdenPago: String(dato[2]),
//               numAutorizacion: dato[3],
//               valorOrden: dato[4],
//               numCuentaDestino: dato[5],
//               codBancoDestino: String(dato[6]),
//               codCompania: String(dato[7]),
//               fechaPago: dato[8],
//               fechaCreacion: dato[9],
//               fechaEnvioCentralizador: dato[10],
//               horaAplicacion: dato[11],
//               tipoProceso: dato[12],
//               estado: dato[13],
//               codMotivoRechazo: dato[14],
//               reenvio: dato[15],
//             });
//             break
//           case 'TERMINADO':
//             objetoBaseEstadoTerminado.push({
//               banco: dato[0],
//               numIdentificacionTercero: String(dato[1]),
//               numOrdenPago: String(dato[2]),
//               numAutorizacion: dato[3],
//               valorOrden: dato[4],
//               numCuentaDestino: dato[5],
//               codBancoDestino: String(dato[6]),
//               codCompania: String(dato[7]),
//               fechaPago: dato[8],
//               fechaCreacion: dato[9],
//               fechaEnvioCentralizador: dato[10],
//               horaAplicacion: dato[11],
//               tipoProceso: dato[12],
//               estado: dato[13],
//               codMotivoRechazo: dato[14],
//               reenvio: dato[15],
//             });
//             break
//           default:
//             break
//         };  
//     });

//     Logger.log(objetoBaseEstadoRechazado);
//     Logger.log(objetoBaseEstadoTerminado);

//     // Actualización de RECHAZADO

//     datosPrincipal.forEach((dato,index)=>{
//       let llave = String(dato[0])+String(dato[3]);
//       let objetoCoincidencia = objetoBaseEstadoRechazado.find((prop)=>{
//         let llaveRechazados = prop.codCompania+prop.numOrdenPago;
//         Logger.log(llaveRechazados)
//         return llaveRechazados == llave
//       });

      

//       if(objetoCoincidencia!= null){
//         Logger.log(index)
//         let numFilaCambio = index+1
//         hojaPrincipal.getRange(numFilaCambio,16).setValue(objetoCoincidencia.banco);
//         hojaPrincipal.getRange(numFilaCambio,17).setValue(objetoCoincidencia.fechaEnvioCentralizador);
//         hojaPrincipal.getRange(numFilaCambio,18).setValue(objetoCoincidencia.numIdentificacionTercero);
//         hojaPrincipal.getRange(numFilaCambio,19).setValue(objetoCoincidencia.fechaPago);
//         hojaPrincipal.getRange(numFilaCambio,20).setValue(objetoCoincidencia.tipoProceso);
//         hojaPrincipal.getRange(numFilaCambio,8).setValue("Rechazado").setBackground("purple");
//         hojaPrincipal.getRange(numFilaCambio,9).setValue(objetoCoincidencia.codMotivoRechazo).setBackground("gray");    
//       };
//     });

//     // Actualización de TERMINADOS

//     let arrayTerminar= [];

//     datosPrincipal.forEach((dato,index)=>{
//       let objetoCoincidencia = objetoBaseEstadoTerminado.find((prop)=>{
//         let llave = String(dato[0])+String(dato[3]);
//         let llaveTerminados = String(prop.codCompania)+String(prop.numOrdenPago);
//         Logger.log(llaveTerminados)
//         return llaveTerminados == llave    
//       }); 

      

//       if(objetoCoincidencia!= null){
//         // Logger.log(index);
//         let numFilaCambio = index+1;
//         Logger.log(numFilaCambio);
//         let logTerminado= [objetoCoincidencia.fechaPago,objetoCoincidencia.codCompania,objetoCoincidencia.numOrdenPago,objetoCoincidencia.numIdentificacionTercero,objetoCoincidencia.valorOrden.replace(".",",")];
//         hojaTerminados.appendRow(logTerminado)
//         arrayTerminar.push(numFilaCambio);
  
//       };
//     });
//     arrayTerminar.reverse().forEach(dato=>{     

//       hojaPrincipal.deleteRow(dato)
//     })  


//     // Logger.log(objetoBaseEstadoRechazado);
//     // Logger.log(objetoBaseEstadoRechazado.sort((a, b) =>{ Date(a.fechaPago) - Date(b.fechaPago)}));

//     // Logger.log("Tamaño de Terminado  /"+objetoBaseEstadoTerminado.length);
//     // Logger.log("Tamaño de Otros  /"+objetoBaseEstadoRechazado.length);

//   } catch (e) {
//     Browser.msgBox("No se pudo cargar el archivo  /" + e);
//     return false; // Failure. Checks for CSV data file error.
//   };

// };


// Versión Refactor 3 ---------------------------------





var objetoBaseEstadoRechazado = [];



function cargarCSVActualizaEstadoRechazado(obj){
  const blob = Utilities.newBlob(Utilities.base64Decode(obj.data),obj.mimeType,obj.fileName);
  const id = '13fy-iESZl51U5XHNYvcgg3QWWpXxqQGi';
  const folder = DriveApp.getFolderById(id);
  const file = folder.createFile(blob);
  const fileURL = file.getUrl();
  const response = {
    'fileName' : obj.fileName,
    'url' : fileURL,
    'status' :true,
    'data' : JSON.stringify(obj)
  };

  // let archivoPrincipal = SpreadsheetApp.getActiveSpreadsheet();
  // let hojaPrueba = archivoPrincipal.getSheetByName("Pruebas");
   procesarCSVActualizaEstadoRechazado(file);


  // Para borrar el archivo una vez tomada la información
  // file.setTrashed(true);

 
  return response;
};


function procesarCSVActualizaEstadoRechazado(csvFile) {

  try {
    // Parses CSV file into data array.
    objetoBaseEstadoRechazado = [];

    // let csvFile = DriveApp.getFileById("1e1tMZNgshANqHjBec25RtLD6zyQISReY")
    let data = Utilities.parseCsv(csvFile.getBlob().getDataAsString());

    data.map(dato => {      
        switch (dato [13]) {
          case 'RECHAZADO':
            // Logger.log(dato[0]);
            objetoBaseEstadoRechazado.push({        
              banco: dato[0],
              numIdentificacionTercero: String(dato[1]),
              numOrdenPago: String(dato[2]),
              numAutorizacion: dato[3],
              valorOrden: dato[4],
              numCuentaDestino: dato[5],
              codBancoDestino: String(dato[6]),
              codCompania: String(dato[7]),
              fechaPago: dato[8],
              fechaCreacion: dato[9],
              fechaEnvioCentralizador: dato[10],
              horaAplicacion: dato[11],
              tipoProceso: dato[12],
              estado: dato[13],
              codMotivoRechazo: dato[14],
              reenvio: dato[15],
            });
            break

          default:
            break
        };  
    });

    // Logger.log(objetoBaseEstadoRechazado);

    // Actualización de RECHAZADO

    datosPrincipal.forEach((dato,index)=>{
      let llave = String(dato[0])+String(dato[3]);
      let objetoCoincidencia = objetoBaseEstadoRechazado.find((prop)=>{
        let llaveRechazados = prop.codCompania+prop.numOrdenPago;
        Logger.log(llaveRechazados)
        return llaveRechazados == llave
      });      

      if(objetoCoincidencia!= null){
        Logger.log(index)
        let numFilaCambio = index+1
        hojaPrincipal.getRange(numFilaCambio,16).setValue(objetoCoincidencia.banco);
        hojaPrincipal.getRange(numFilaCambio,17).setValue(objetoCoincidencia.fechaEnvioCentralizador);
        hojaPrincipal.getRange(numFilaCambio,18).setValue(objetoCoincidencia.numIdentificacionTercero);
        hojaPrincipal.getRange(numFilaCambio,19).setValue(objetoCoincidencia.fechaPago);
        hojaPrincipal.getRange(numFilaCambio,20).setValue(objetoCoincidencia.tipoProceso);
        hojaPrincipal.getRange(numFilaCambio,8).setValue("Rechazado").setBackground("purple");
        hojaPrincipal.getRange(numFilaCambio,9).setValue(objetoCoincidencia.codMotivoRechazo).setBackground("gray");    
      };
    });

  } catch (e) {
    Browser.msgBox("No se pudo cargar el archivo  /" + e);
    return false; // Failure. Checks for CSV data file error.
  };

}; 

var objetoBaseEstadoTerminado = [];


function cargarCSVActualizaEstadoTerminado(obj){
  const blob = Utilities.newBlob(Utilities.base64Decode(obj.data),obj.mimeType,obj.fileName);
  const id = '13fy-iESZl51U5XHNYvcgg3QWWpXxqQGi';
  const folder = DriveApp.getFolderById(id);
  const file = folder.createFile(blob);
  const fileURL = file.getUrl();
  const response = {
    'fileName' : obj.fileName,
    'url' : fileURL,
    'status' :true,
    'data' : JSON.stringify(obj)
  };

  // let archivoPrincipal = SpreadsheetApp.getActiveSpreadsheet();
  // let hojaPrueba = archivoPrincipal.getSheetByName("Pruebas");
   procesarCSVActualizaEstadoTerminado(file);


  // Para borrar el archivo una vez tomada la información
  // file.setTrashed(true);

 
  return response;
};

function procesarCSVActualizaEstadoTerminado(csvFile) {

  try {

    objetoBaseEstadoTerminado = [];

    // let csvFile = DriveApp.getFileById("1e1tMZNgshANqHjBec25RtLD6zyQISReY")
    let data = Utilities.parseCsv(csvFile.getBlob().getDataAsString());

    data.map(dato => {      
        switch (dato [13]) {
          case 'TERMINADO':
            objetoBaseEstadoTerminado.push({
              banco: dato[0],
              numIdentificacionTercero: String(dato[1]),
              numOrdenPago: String(dato[2]),
              numAutorizacion: dato[3],
              valorOrden: dato[4],
              numCuentaDestino: dato[5],
              codBancoDestino: String(dato[6]),
              codCompania: String(dato[7]),
              fechaPago: dato[8],
              fechaCreacion: dato[9],
              fechaEnvioCentralizador: dato[10],
              horaAplicacion: dato[11],
              tipoProceso: dato[12],
              estado: dato[13],
              codMotivoRechazo: dato[14],
              reenvio: dato[15],
            });
            break
          default:
            break
        };  
    });


    // Actualización de TERMINADOS

    let arrayTerminar= [];

    datosPrincipal.forEach((dato,index)=>{
      let objetoCoincidencia = objetoBaseEstadoTerminado.find((prop)=>{
        let llave = String(dato[0])+String(dato[3]);
        let llaveTerminados = String(prop.codCompania)+String(prop.numOrdenPago);
        // Logger.log(llaveTerminados)
        return llaveTerminados == llave    
      });

      if(objetoCoincidencia!= null){
        // Logger.log(index);
        let numFilaCambio = index+1;
        Logger.log(numFilaCambio);
        let logTerminado= [objetoCoincidencia.fechaPago,objetoCoincidencia.codCompania,objetoCoincidencia.numOrdenPago,objetoCoincidencia.numIdentificacionTercero,objetoCoincidencia.valorOrden.replace(".",",")];
        hojaTerminados.appendRow(logTerminado)
        arrayTerminar.push(numFilaCambio);  
      };
    });
    arrayTerminar.reverse().forEach(dato=>{ 
      hojaPrincipal.deleteRow(dato)
    })  

    // Logger.log("Tamaño de Terminado  /"+objetoBaseEstadoTerminado.length);


  } catch (e) {
    Browser.msgBox("No se pudo cargar el archivo  /" + e);
    return false; // Failure. Checks for CSV data file error.
  };

};












// Pruebas ---------------------------------

function pruebasGenerales(){

  Logger.log(datosPrincipal)




  // let registrosActuales = datosPrincipal.map(dato=>{
  //   let llave = String(dato[0])+String(dato[3]);

  //   return llave
  //   });

  // 


  // let csvFile = DriveApp.getFileById("1e1tMZNgshANqHjBec25RtLD6zyQISReY");
  //   let data = Utilities.parseCsv(csvFile.getBlob().getDataAsString());
    
  //   Logger.log(data);

  //   data.map(dato => {      
  //       switch (dato [13]) {
  //         case 'RECHAZADO':
  //           // Logger.log(dato[0]);
  //           objetoBaseEstadoRechazado.push({        
  //             banco: dato[0],
  //             numIdentificacionTercero: String(dato[1]),
  //             numOrdenPago: String(dato[2]),
  //             numAutorizacion: dato[3],
  //             valorOrden: dato[4],
  //             numCuentaDestino: dato[5],
  //             codBancoDestino: String(dato[6]),
  //             codCompania: String(dato[7]),
  //             fechaPago: dato[8],
  //             fechaCreacion: dato[9],
  //             fechaEnvioCentralizador: dato[10],
  //             horaAplicacion: dato[11],
  //             tipoProceso: dato[12],
  //             estado: dato[13],
  //             codMotivoRechazo: dato[14],
  //             reenvio: dato[15],
  //           });
  //           break
  //         case 'TERMINADO':
  //           objetoBaseEstadoTerminado.push({
  //             banco: dato[0],
  //             numIdentificacionTercero: String(dato[1]),
  //             numOrdenPago: String(dato[2]),
  //             numAutorizacion: dato[3],
  //             valorOrden: dato[4],
  //             numCuentaDestino: dato[5],
  //             codBancoDestino: String(dato[6]),
  //             codCompania: String(dato[7]),
  //             fechaPago: dato[8],
  //             fechaCreacion: dato[9],
  //             fechaEnvioCentralizador: dato[10],
  //             horaAplicacion: dato[11],
  //             tipoProceso: dato[12],
  //             estado: dato[13],
  //             codMotivoRechazo: dato[14],
  //             reenvio: dato[15],
  //           });
  //           break
  //         default:
  //           break
  //       };  
  //   });

  //   Logger.log(objetoBaseEstadoRechazado);
  //   Logger.log(objetoBaseEstadoTerminado);
  

  // let reenvios = datosPrincipal.map(dato=>String(dato[3]))

  // Logger.log(reenvios);
  // let numOrdenPago="92102021001800";

  // let filaActualizar = datosPrincipal.findIndex(dato=>dato[3]==numOrdenPago); 
  // Logger.log(filaActualizar);

  // let existeEnPrincipal = filaActualizar!= -1 ? true : false ;
  // Logger.log(existeEnPrincipal);
  // const encabezadosBasePrincipal = datosPrincipal[0];

  // const datosPrueba = hojaPruebas.getDataRange().getValues();
  // const encabezadosBasePruebas = datosPrueba[0];
  // var datoraw = datosPrincipal[36][5];
  // var datotiporaw = typeof datosPrincipal[36][5];
  // var dato = Number (datosPrincipal[36][5]);
  // var datotipo = typeof dato;

  // Logger.log(datoraw);
  // Logger.log(datotiporaw);
  // Logger.log(dato);
  // Logger.log(datotipo);



  // let registrosActuales = datosPrincipal.map(dato=>String(dato[3]));
  // Logger.log(registrosActuales);

  // let incluirRegistro = registrosActuales.includes("92102021001369");

  // Logger.log(incluirRegistro);
  // Logger.log(`Encabezados Pruebas: ${encabezadosBasePruebas}`);

};