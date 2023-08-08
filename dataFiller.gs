function infoBoat(boat_name, s_botes, s_invoice){
  //con el nombre del bote, busco en la hoja de botes por la info correspondiente a dicho barco,
  //y la completo en la hoja del invoice
  var rng_boats = s_botes.getDataRange();
  var cells_boats = rng_boats.getValues();
  //antes de llenar la hoja de invoice con info, la vacio de info previa
  s_invoice.getRange("B10").clearContent();
  s_invoice.getRange("B11").clearContent();
  s_invoice.getRange("B12").clearContent();
  s_invoice.getRange("B13").clearContent();
  s_invoice.getRange("B14").clearContent();
  //procedo a buscar por la hoja de barcos por el nombre del barco elegido por el usuario
  for (var i = 0; i < rng_boats.getLastRow(); i++) {
    if (cells_boats[i][0] == boat_name) {
      Logger.log("Barco encontrado en la fila " + i);
      //guardo la info del barco en la hoja del invoice
      var dir_fiscal_barco = cells_boats[i][1];
      s_invoice.getRange("B10").setValue(dir_fiscal_barco);
      var dir_fisica_barco = cells_boats[i][2];
      s_invoice.getRange("B11").setValue(dir_fisica_barco);
      var dir_fisica2_barco = cells_boats[i][3];
      s_invoice.getRange("B12").setValue(dir_fisica2_barco);
      var dir_fisica3_barco = cells_boats[i][4];
      s_invoice.getRange("B13").setValue(dir_fisica3_barco);
      var dir_area_barco = cells_boats[i][5];
      s_invoice.getRange("B14").setValue(dir_area_barco);
    }
  }
}

function completarInvoice(rng_res, i, s_botes, s_invoice, nwss, tipoPago){
  //creo las variables para guardar la info importante que voy
  //a querer pegar en la hoja de invoice
  //creo la variable de fecha, de la hoja resumen
  var dateVal = rng_res.getCell(i,1).getValue();
  var date = Utilities.formatDate(dateVal,"GMT", "ddMMYYYY");
  //la paso a la hoja del invoice
  s_invoice.getRange("F7").clearContent();
  s_invoice.getRange("F7").setValue(date);
  //creo la variable del numero de invoice, de la hoja resumen
  var inv_num = rng_res.getCell(i, 2).getValue();
  //se la paso a la hoja del invoice
  s_invoice.getRange("B7").clearContent();
  s_invoice.getRange("B7").setValue("INVOICE Nº " + inv_num);
  //creo la variable del nombre del bote, de la hoja resumen
  var boat_name = rng_res.getCell(i, 3).getValue();
  //primero, utilizo el nombre del barco para modificar el invoice
  s_invoice.getRange("B9").clearContent();
  s_invoice.getRange("B9").setValue("INVOICE TO M/Y " + boat_name);
  //luego con el nombre del barco voy a la hoja de botes y obtengo otra info para modificar el invoice
  infoBoat(boat_name, s_botes, s_invoice)
  //creo la variable de los pacientes, de la hoja resumen
  var patients = rng_res.getCell(i, 4).getValue();
  //le paso esta info a la hoja de invoice
  s_invoice.getRange("B18").clearContent();
  s_invoice.getRange("B18").setValue(patients);
  //creo la variable para el tipo de visita
  var type_visit = rng_res.getCell(i, 5).getValue();
  //guardo dicha info en la hoja de invoice
  s_invoice.getRange("B27").clearContent();
  s_invoice.getRange("B27").setValue(type_visit);
  //creo la variable notes de la hoja resumen
  var notes = rng_res.getCell(i, 13).getValue();
  //guardo la info en invoice?                 ????????
  
  //creo la variable de diagnostico de la hoja resumen
  var diagnoses = rng_res.getCell(i, 14).getValue();
  //guardo la info de los diagnosticos en el invoice
  s_invoice.getRange("E18").clearContent();
  s_invoice.getRange("E18").setValue(diagnoses);
  //creo una variable para ver si hay que hacer fup
  var fup = rng_res.getCell(i, 15).getValue();
  //guardo la info en invoice
  s_invoice.getRange("C24").clearContent();
  s_invoice.getRange("C24").setValue(fup);

  //creo las variables para los precios de los pacientes
  var p1 = rng_res.getCell(i, 17).getValue();
  var p2 = rng_res.getCell(i, 18).getValue();
  var p3 = rng_res.getCell(i, 19).getValue();
  var p4 = rng_res.getCell(i, 20).getValue();
  var p5 = rng_res.getCell(i, 21).getValue();
  var p6 = rng_res.getCell(i, 22).getValue();
  //limpio las celdas donde voy a guardar estos valores en el invoice
  s_invoice.getRange("C27").clearContent();
  s_invoice.getRange("C28").clearContent();
  s_invoice.getRange("C29").clearContent();
  s_invoice.getRange("C30").clearContent();
  s_invoice.getRange("C31").clearContent();
  s_invoice.getRange("C32").clearContent();
  //guardo las variables de precios en el invoice
  s_invoice.getRange("C27").setValue(p1);
  s_invoice.getRange("C28").setValue(p2);
  s_invoice.getRange("C29").setValue(p3);
  s_invoice.getRange("C30").setValue(p4);
  s_invoice.getRange("C31").setValue(p5);
  s_invoice.getRange("C32").setValue(p6);

  //con la info recolectada, renombro a la spreadsheet que estamos creando
  nwss.setName("SANTE_MEDICAL-" + inv_num + "-" + boat_name + "-" + date + "-" + tipoPago);
}


function main_crearInvoices() {
  //defino la variable de la spreadsheet modelo
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //creo las sheets que voy a utilizar
  var s_resumen = ss.getSheetByName("Resumen");
  var s_cash = ss.getSheetByName("CASH");
  var s_evo = ss.getSheetByName("PAYABLE BY EVOLUTION");
  var s_botes = ss.getSheetByName("Botes");
  //itero por las filas que van a crear los invoices, con el rango seleccionado por el usuario. D
  //ependiendo de si se pago en cash o por evolution, uso una sheet u otra
  //defino el rango activo (seleccionado por el usuario)
  var rng_res = s_resumen.getActiveRange();
  //itero por todas las
  for (var i = 1; i <= rng_res.getNumRows(); i++){
    //creo la nueva worksheet que tendra al invoice
    var nwss = SpreadsheetApp.create("invoice_" + i)
    //muevo la spreadsheet creada a la carpeta deseada
    var folderId = "1Qr5jd-L21MHOhQQE8vUxZiGUV3lUJ18L";
    var folder = DriveApp.getFolderById(folderId);
    DriveApp.getFileById(nwss.getId()).moveTo(folder);
    //la columna con el valor de payable by evolution es la 9, empezando desde 1
    var evo = rng_res.getCell(i, 9).getValue();
    //Logger.log(evo);
    if (evo != 0){
      //si se pago con evolution, entonces trabajamos con la hoja de payable by evolution
      Logger.log("Pago por evo");
      //copio la hoja de payable by evolution a la nueva spreadsheet
      s_evo.copyTo(nwss);
      //borro la primera hoja que aparece por default
      nwss.deleteSheet(nwss.getSheets()[0]);
      //defino la nueva hoja evo perteneciente a nwss
      var ns_evo = nwss.getSheets()[0];
      //le cambio el nombre que viene por deafult
      var tipoPago = "PAYABLE_BY_EVOLUTION";
      ns_evo.setName(tipoPago);
      //se la paso a la funcion para que la complete
      completarInvoice(rng_res, i, s_botes, ns_evo, nwss, tipoPago);
       //una vez creado el invoice, cambiamos la fila de pending a done en la hoja resumen
      var cellPending = rng_res.getCell(i, 7);
      cellPending.setValue("DONE");
  } else {
      //si no se pago con evolution, se pago con cash. Trabajamos con dicha hoja
      Logger.log("Pago con Cash");
      //copio la hoja de payable by evolution a la nueva spreadsheet
      s_cash.copyTo(nwss);
      //borro la primera hoja que aparece por default
      nwss.deleteSheet(nwss.getSheets()[0]);
      //defino la nueva hoja cash perteneciente a nwss
      var ns_cash = nwss.getSheets()[0];
      //le cambio el nombre que viene por deafult
      var tipoPago = "CASH";
      ns_cash.setName(tipoPago);
      //se la paso a la funcion para que la complete
      completarInvoice(rng_res, i, s_botes, ns_cash, nwss,tipoPago);
      //una vez creado el invoice, cambiamos la fila de pending a done en la hoja resumen
      var cellPending = rng_res.getCell(i, 7);
      cellPending.setValue("DONE");
  }
  }
}


