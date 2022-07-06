var id_hojaConsolidado = Cambiolib.llamarHoja("Hoja_Consolidado");
var hoja_consolidado = SpreadsheetApp.openById(id_hojaConsolidado);
var libro_consolidado = hoja_consolidado.getSheetByName("Consolidado");

var id_hojaInstanciadora = Cambiolib.llamarHoja("Hoja_instanciadora");
var hoja_instanciadora = SpreadsheetApp.openById(id_hojaInstanciadora);
var libro_agentes = hoja_instanciadora.getSheetByName("Agentes");
var libro_correos = hoja_instanciadora.getSheetByName("Cuerpos_Correos");
var libro_Responsables = hoja_instanciadora.getSheetByName("Responsables");
var libro_procedimientos = hoja_instanciadora.getSheetByName("Procedimientos");

var id_folderEliminada = Cambiolib.llamarCarpeta("Eliminadas");


function doPost(request) {
  
  var params = JSON.stringify(request);
  var codigo = request.parameter.prodId;
  codigo = codigo.trim();
  var finalizar_creacion = request.parameter.finalizar_creacion;
  var finalizar_eliminacion = request.parameter.Finalizar_eliminacion;
  var finalizar_actualizacion = request.parameter.Finalizar_actualizacion;
  
  
  if (finalizar_creacion) {
    fin_creacion (codigo);
    return HtmlService.createTemplateFromFile('Finalizacion').evaluate();
  }
  
  if (finalizar_eliminacion) {
    fin_eli (codigo);
    return HtmlService.createTemplateFromFile('finalizacion_eli').evaluate();
  }
  
  if (finalizar_actualizacion) {
    fin_actu(codigo);
    return HtmlService.createTemplateFromFile('finalizacion_actua').evaluate();
  }
}










//    MailApp.sendEmail("camilo.rojas@davincitech.co","prueba",request.parameter.tipo)
//  return HtmlService.createHtmlOutput(params);









function fin_creacion (codigo){
  var fila_consol =  buscarFila (libro_consolidado,libro_consolidado.getLastRow(),1,codigo);
  if (fila_consol > 0 ){
    var agente = libro_consolidado.getRange(fila_consol, 11).getValue();
    var fila_agente = buscarFila (libro_agentes,libro_agentes.getLastRow(),1,agente);
    
    if (fila_agente > 0){
      var correo_agente = libro_agentes.getRange(fila_agente, 2).getValue();
      var proceso = libro_consolidado.getRange(fila_consol, 7).getValue();
      var tipo_solicitud = libro_consolidado.getRange(fila_consol, 9).getValue();
      var motivo_cambio = libro_consolidado.getRange(fila_consol, 10).getValue();
      var link_proceso = libro_consolidado.getRange(fila_consol, 17).getValue();
      var fila_encargados = buscarFila (libro_Responsables,libro_Responsables.getLastRow(), 1,proceso);
      
      if(fila_encargados > 0 ){
        
        
        var datos_responsables = libro_Responsables.getRange(fila_encargados, 1, 1, libro_Responsables.getLastColumn()).getValues();
        
        var col = 2;
        var buscar = false;
        while(col <= libro_Responsables.getLastColumn()-1){
          
          if(datos_responsables[0][col] == correo_agente){
            buscar = true;
            break;
          }
          col = col+2;
        }
        
        
        if (buscar){
          var asunto = libro_correos.getRange("B7").getValue();
          asunto = asunto.replace("NUMERO_1",codigo);
          var body = libro_correos.getRange("C7").getValue();
          body= body.replace("NUMERO_1 ",codigo);
          body= body.replace("NOMBRE_ASESOR",agente);
          body= body.replace("NOMBRE_PROCESO",proceso);
          
          body= body.replace("LINK_PROCESO",link_proceso);
          body= body.replace("CODIGO",codigo);
          body= body.replace("TIPO_SOLICITUD", tipo_solicitud);
          body= body.replace("MOTIVO_SOLICITUD", motivo_cambio);
          
          MailApp.sendEmail({
            to: correo_agente,
            subject: asunto,
            htmlBody: body,
            noReply: true
            
          });
        }else {
          var asunto = libro_correos.getRange("B6").getValue();
          asunto = asunto.replace ("NUMERO_1", codigo);
          var body = libro_correos.getRange("C6").getValue();
          body = body.replace ("NUMERO_1", codigo);
          body = body.replace ("NOMBRE_PROCESO", proceso);
          body = body.replace ("LINK_PROCESO", link_proceso);
          body = body.replace ("CODIGO", codigo);
          
          MailApp.sendEmail({
            to: correo_agente,
            subject: asunto,
            htmlBody: body,
            noReply: true
            
          });
        }
        
        
        var now = new Date();
        libro_consolidado.getRange(fila_consol, 15).setValue("Finalizado");
        libro_consolidado.getRange(fila_consol, 16).setValue(Utilities.formatDate(now, "GMT-5", "dd-MM-yyyy"));
      }
      
      
      
    }
    
    
    
    
  }
}

function fin_eli (codigo){
  var fila_consol =  buscarFila (libro_consolidado,libro_consolidado.getLastRow(),1,codigo);
  if (fila_consol > 0 ){
    var agente = libro_consolidado.getRange(fila_consol, 11).getValue();
    var correo_solicitante = libro_consolidado.getRange(fila_consol, 3).getValue();
    var fila_agente = buscarFila (libro_agentes,libro_agentes.getLastRow(),1,agente);
    
    if (fila_agente > 0){
      var correo_agente = libro_agentes.getRange(fila_agente, 2).getValue();
      var correo_solicitante = libro_agentes.getRange(fila_agente, 3).getValue();
      var proceso = libro_consolidado.getRange(fila_consol, 7).getValue();
      var link_proceso = libro_consolidado.getRange(fila_consol, 17).getValue();
      var tipo_solicitud = libro_consolidado.getRange(fila_consol, 9).getValue();
      var motivo_cambio = libro_consolidado.getRange(fila_consol, 10).getValue();
      var url2 = link_proceso.toString();
      var url3 = url2.split("/");
      var id_doc = url3[5];
      
      var fila_encargados = buscarFila (libro_Responsables,libro_Responsables.getLastRow(), 1,proceso);
      
      if(fila_encargados > 0 ){
        
        
        var datos_responsables = libro_Responsables.getRange(fila_encargados, 1, 1, libro_Responsables.getLastColumn()).getValues();
        
        var col = 2;
        var buscar = false;
        while(col <= libro_Responsables.getLastColumn()-1){
          
          if(datos_responsables[0][col] == correo_solicitante){
            buscar = true;
            break;
          }
          col = col+2;
        }
        
        var folder_eliminadas = DriveApp.getFolderById(id_folderEliminada); 
        var id_doc = DocumentApp.openByUrl(link_proceso).getId();
        
        var folder = DriveApp.getFolderById(id_folderEliminada).addFile(DriveApp.getFileById(id_doc));
        
        var fila_carpeta = buscarFila (libro_procedimientos,libro_procedimientos.getLastRow(),1,proceso);
        
        if (fila_carpeta > 0){
          var url_busqueda = libro_procedimientos.getRange(fila_carpeta, 4).getValue();
          var contador = 0;
          var array_url_busqueda = url_busqueda.split("/");
          var id_busqueda = array_url_busqueda[array_url_busqueda.length - 1];
          id_busqueda = id_busqueda.split("?");
          
          var id_carpeta = id_busqueda[0];
          var documento = DriveApp.getFileById(id_doc);
          DriveApp.getFolderById(id_carpeta).removeFile(DriveApp.getFileById(documento.getId()))
          var editores = documento.getEditors();
        Logger.log(editores);
        
        for (var i in editores){
          documento.removeEditor(editores[i])
        }
        }
        
        
        
        var fila_encargados = buscarFila (libro_procedimientos,libro_procedimientos.getLastRow(), 1,proceso);
        
        if (fila_encargados>1){
          var datos_responsables = libro_Responsables.getRange(fila_encargados, 1, 1, libro_Responsables.getLastColumn()).getValues();
          
          var col = 2;
          while(col <= libro_Responsables.getLastColumn()-1){
            if(datos_responsables[0][col] != ""){
              documento.addEditor(datos_responsables[0][col]);
            }
            col = col+2;
          }
        }
        
        
        
        
        
        
        if (buscar){
          
          
          
          var asunto = libro_correos.getRange("B14").getValue();
          asunto = asunto.replace("NUMERO_1",codigo);
          var body = libro_correos.getRange("C14").getValue();
          body= body.replace("NUMERO_1 ",codigo);
          body= body.replace("NOMBRE_ASESOR",agente);
          body= body.replace("NOMBRE_PROCESO",proceso);
          
          body= body.replace("LINK_PROCESO",link_proceso);
          body= body.replace("CODIGO",codigo);
          body= body.replace("TIPO_SOLICITUD", tipo_solicitud);
          body= body.replace("MOTIVO_SOLICITUD", motivo_cambio);
          MailApp.sendEmail({
            to: correo_agente,
            subject: asunto,
            htmlBody: body,
            noReply: true
            
          });
        }else {
          var asunto = libro_correos.getRange("B15").getValue();
          asunto = asunto.replace ("NUMERO_1", codigo);
          var body = libro_correos.getRange("C15").getValue();
          body = body.replace ("NUMERO_1", codigo);
          body = body.replace ("NOMBRE_PROCESO", proceso);
          body = body.replace ("LINK_PROCESO", link_proceso);
          body = body.replace ("CODIGO", codigo);
          
          MailApp.sendEmail({
            to: correo_agente,
            subject: asunto,
            htmlBody: body,
            noReply: true
            
          });
          
          var asunto = libro_correos.getRange("B16").getValue();
          asunto = asunto.replace ("NUMERO_1", codigo);
          var body = libro_correos.getRange("C16").getValue();
          body = body.replace ("NUMERO_1", codigo);
          body = body.replace ("NOMBRE_PROCESO", proceso);
          body = body.replace ("LINK_PROCESO", link_proceso);
          
          MailApp.sendEmail({
            to: correo_solicitante,
            subject: asunto,
            htmlBody: body,
            noReply: true
            
          });
          
          
          
          
        }
        
        
        var now = new Date();
        libro_consolidado.getRange(fila_consol, 15).setValue("Finalizado");
        libro_consolidado.getRange(fila_consol, 16).setValue(Utilities.formatDate(now, "GMT-5", "dd-MM-yyyy"));
      }
      
      
      
    }
    
    
    
    
  }
}

function fin_actu(codigo) {
  var fila_consol =  buscarFila (libro_consolidado,libro_consolidado.getLastRow(),1,codigo);
  if (fila_consol > 0 ){
    var agente = libro_consolidado.getRange(fila_consol, 11).getValue();
    var fila_agente = buscarFila (libro_agentes,libro_agentes.getLastRow(),1,agente);
    
    if (fila_agente > 0){
      var correo_agente = libro_agentes.getRange(fila_agente, 2).getValue();
      var proceso = libro_consolidado.getRange(fila_consol, 7).getValue();
      var link_proceso = libro_consolidado.getRange(fila_consol, 17).getValue();
      var url2 = link_proceso.toString();
      var url3 = url2.split("/");
      var id_doc = url3[5];
      
      var fila_encargados = buscarFila (libro_Responsables,libro_Responsables.getLastRow(), 1,proceso);
      
      if(fila_encargados > 0 ){
        
        
        var datos_responsables = libro_Responsables.getRange(fila_encargados, 1, 1, libro_Responsables.getLastColumn()).getValues();
        
        var col = 2;
        var buscar = false;
        while(col <= libro_Responsables.getLastColumn()-1){
          
          if(datos_responsables[0][col] == correo_agente){
            buscar = true;
            break;
          }
          col = col+2;
        }
        
        
        if (buscar){
          var asunto = libro_correos.getRange("B7").getValue();
          asunto = asunto.replace("NUMERO_1",codigo);
          var body = libro_correos.getRange("C7").getValue();
          body= body.replace("NUMERO_1 ",codigo);
          body= body.replace("NOMBRE_ASESOR",agente);
          body= body.replace("NOMBRE_PROCESO",proceso);
          
          body= body.replace("LINK_PROCESO",link_proceso);
          body= body.replace("CODIGO",codigo);
          
          MailApp.sendEmail({
            to: correo_agente,
            subject: asunto,
            htmlBody: body,
            noReply: true
            
          });
        }else {
          var asunto = libro_correos.getRange("B6").getValue();
          asunto = asunto.replace ("NUMERO_1", codigo);
          var body = libro_correos.getRange("C6").getValue();
          body = body.replace ("NUMERO_1", codigo);
          body = body.replace ("NOMBRE_PROCESO", proceso);
          body = body.replace ("LINK_PROCESO", link_proceso);
          body = body.replace ("CODIGO", codigo);
          
          MailApp.sendEmail({
            to: correo_agente,
            subject: asunto,
            htmlBody: body,
            noReply: true
            
          });
        }
        
        
        var now = new Date();
        libro_consolidado.getRange(fila_consol, 15).setValue("Finalizado");
        libro_consolidado.getRange(fila_consol, 16).setValue(Utilities.formatDate(now, "GMT-5", "dd-MM-yyyy"));
      }
      
      
      
    }
    
    
    
    
  }
}




function buscarFila (hojaBusqueda,ultimaFila,columnaBusqueda,dato) {
  var busqueda = false;
  var datosBusqueda = hojaBusqueda.getRange(1,columnaBusqueda, ultimaFila, 1).getValues();
  for (var filaBusqueda in datosBusqueda){
    var dato2 = datosBusqueda[filaBusqueda];
    if (dato == dato2){
      var filaDato = filaBusqueda;
      filaDato++;
      busqueda = true;
      break;
    }
    if (busqueda == false ) {
      var filaDato = -1; 
    }
    
  }
  return filaDato;
}