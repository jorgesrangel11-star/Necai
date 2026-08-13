const ID_CARPETAS_RAIZ = "1Yil306--S02_ZMarTso42AvEKaWWkyxa";

function conectarCarpetasAlExcel() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getSheets()[0];
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  const carpetaRaiz = DriveApp.getFolderById(ID_CARPETAS_RAIZ);
  const subcarpetas = carpetaRaiz.getFolders();
  
  let mapaCarpetas = {};
  while (subcarpetas.hasNext()) {
    let sub = subcarpetas.next();
    let nombre = sub.getName();
    let numeroStr = nombre.split("-")[0].trim();
    let numero = parseInt(numeroStr, 10);
    
    if (!isNaN(numero)) {
      mapaCarpetas[numero] = sub.getUrl();
    }
  }

  const inventarios = hoja.getRange(2, 2, ultimaFila - 1, 1).getValues();
  let linksParaPegar = [];

  for (let i = 0; i < inventarios.length; i++) {
    let inv = parseInt(inventarios[i][0], 10);
    if (mapaCarpetas[inv]) {
      linksParaPegar.push([mapaCarpetas[inv]]);
    } else {
      linksParaPegar.push([""]); 
    }
  }

  hoja.getRange(2, 14, linksParaPegar.length, 1).setValues(linksParaPegar);
}

function obtenerListaParques() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  
  var datos = hoja.getRange(2, 2, ultimaFila - 1, 9).getValues(); 
  var listaParques = [];
  
  for (var i = 0; i < datos.length; i++) {
    if (datos[i][0] != "") { 
      listaParques.push({ 
        id: datos[i][0].toString().trim(),         
        nombre: datos[i][1].toString().trim(),     
        m2: datos[i][8] !== "" ? datos[i][8] : "No registrado" 
      });
    }
  }
  return listaParques; 
}

function formatearTextoFecha(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return valor ? valor.toString() : "";
}

function doPost(e) {
  try {
    var datos = JSON.parse(e.postData.contents);
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = libro.getSheets()[0]; 
    var ultimaFila = hoja.getLastRow();
    
    var filaParque = -1;
    var inventarios = hoja.getRange(2, 2, ultimaFila - 1, 1).getValues(); 
    var nombreParqueEncontrado = "";
    
    for (var i = 0; i < inventarios.length; i++) {
      if (inventarios[i][0].toString().trim() === datos.inventario.toString().trim()) {
        filaParque = i + 2; 
        nombreParqueEncontrado = hoja.getRange(filaParque, 3).getValue();
        break;
      }
    }
    
    if (filaParque === -1) {
      return ContentService.createTextOutput(JSON.stringify({"result": "error", "error": "Inventario no encontrado en el Excel"})).setMimeType(ContentService.MimeType.JSON);
    }
    
    // CORRECCIÓN: En lugar de acumular texto y hacer la celda gigante, SOBREESCRIBE con el mantenimiento más reciente
    hoja.getRange(filaParque, 1).setValue(datos.fecha);
    hoja.getRange(filaParque, 7).setValue(datos.gps_coordenadas || "Sin GPS");
    hoja.getRange(filaParque, 8).setValue(datos.supervisor);
    
    var textoDesc = "--- Último Reporte (" + datos.fecha + ") ---\n📐 Superficie: " + datos.m2_parque + " m²\n" + datos.descripcion;
    hoja.getRange(filaParque, 9).setValue(textoDesc);
    
    hoja.getRange(filaParque, 1, 1, 14).setWrap(true).setVerticalAlignment("top");
    
    var urlCarpeta = hoja.getRange(filaParque, 14).getValue(); 
    if (!urlCarpeta) {
      return ContentService.createTextOutput(JSON.stringify({"result": "error", "error": "El parque no tiene carpeta asignada"})).setMimeType(ContentService.MimeType.JSON);
    }
    
    var idCarpeta = urlCarpeta.split("folders/")[1].split("?")[0].trim();
    var carpetaParque = DriveApp.getFolderById(idCarpeta);
    var tiempoUnico = new Date().getTime(); 
    
    var linksAntes = [];
    var linksDurante = [];
    var linksDespues = [];

    // PROCESAMIENTO DE FOTOS EN LA CARPETA DE DRIVE CORRESPONDIENTE
    function procesarFotos(arrayFotos, prefijo, columna, arregloLinks) {
      if (arrayFotos && arrayFotos.length > 0) {
        for (var f = 0; f < arrayFotos.length; f++) {
          var nombreArchivo = prefijo + "_" + datos.inventario + "_" + tiempoUnico + "_" + f + ".jpg";
          var blob = Utilities.newBlob(Utilities.base64Decode(arrayFotos[f]), "image/jpeg", nombreArchivo);
          var archivo = carpetaParque.createFile(blob);
          archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          arregloLinks.push(archivo.getUrl());
        }
        
        // En la pestaña principal actualiza con los links de la última visita solamente
        var celda = hoja.getRange(filaParque, columna);
        var encabezado = "-- Visita del " + datos.fecha + " --\n";
        var urlsTexto = arregloLinks.join("\n");
        var textoFinal = encabezado + urlsTexto;
        
        var constructorRichText = SpreadsheetApp.newRichTextValue().setText(textoFinal);
        var regex = /https?:\/\/[^\s]+/g;
        var match;
        while ((match = regex.exec(textoFinal)) !== null) {
          constructorRichText.setLinkUrl(match.index, match.index + match[0].length, match[0]);
        }
        celda.setRichTextValue(constructorRichText.build());
      }
    }

    procesarFotos(datos.fotoAntes, "Antes", 11, linksAntes);     
    procesarFotos(datos.fotoDurante, "Durante", 12, linksDurante); 
    procesarFotos(datos.fotoDespues, "Despues", 13, linksDespues); 
    
    // ==========================================
    // SECCIÓN DE HISTORIAL EN BITÁCORA (Fila nueva automática siempre)
    // ==========================================
    var hojaBitacora = libro.getSheetByName("Historial_Bitácora");
    if (!hojaBitacora) {
      hojaBitacora = libro.insertSheet("Historial_Bitácora");
      var encabezados = ["Fecha y Hora", "No. Inventario", "Nombre del Parque", "m²", "Supervisor", "Coordenadas GPS", "Descripción del Trabajo", "Fotos Antes", "Fotos Durante", "Fotos Después", "Carpeta del Parque"];
      hojaBitacora.appendRow(encabezados);
      hojaBitacora.getRange("A1:K1").setFontWeight("bold").setBackground("#1e3a8a").setFontColor("#ffffff");
      hojaBitacora.setFrozenRows(1);
    }
    
    var marcaTiempo = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    
    var nuevaFilaBitacora = [
      marcaTiempo,
      datos.inventario,
      nombreParqueEncontrado,
      datos.m2_parque,
      datos.supervisor,
      datos.gps_coordenadas || "Sin GPS",
      datos.descripcion,
      linksAntes.join("\n"),
      linksDurante.join("\n"),
      linksDespues.join("\n"),
      urlCarpeta
    ];
    
    // INSERTAR FILA NUEVA Y AUTOMÁTICA EN LA PESTAÑA HISTORIAL_BITÁCORA
    hojaBitacora.appendRow(nuevaFilaBitacora);
    var filaInsertada = hojaBitacora.getLastRow();
    hojaBitacora.getRange(filaInsertada, 1, 1, 11).setWrap(true).setVerticalAlignment("top");

    return ContentService.createTextOutput(JSON.stringify({"result": "success"})).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({"result": "error", "error": error.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function procesarDesdeFront(datosString) {
  var e = { postData: { contents: datosString } };
  return JSON.parse(doPost(e).getContent());
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Sistema de Registro - Necai')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no') 
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}