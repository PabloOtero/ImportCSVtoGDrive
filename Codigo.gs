function onOpen() {
   
  var ss = SpreadsheetApp.getActiveSpreadsheet();   
  var csvMenuEntries = [{name: "Cargar archivo CSV", functionName: "importFromCSV"}];
  ss.addMenu("PROTECMA-Importar", csvMenuEntries);
}

function importFromCSV() {
 var fileName = Browser.inputBox("Introduce el nombre del archivo que has previamente guardado en tu lista de documentos (ej. user_list.csv):");
 
 var files = DocsList.getFiles();
 var csvFile = "";

  for (var i = 0; i < files.length; i++) {
    if (files[i].getName() == fileName) {
      csvFile = files[i].getContentAsString();
      break;
    }
  }
  var csvData = CSVToArray(csvFile, ",");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  sheet.clear();
  
  for (var i = 0; i < csvData.length; i++) {
    sheet.getRange(i+1, 1, 1, csvData[i].length).setValues(new Array(csvData[i]));
  }
  
  
  formattable();
  
}

// http://www.bennadel.com/blog/1504-Ask-Ben-Parsing-CSV-Strings-With-Javascript-Exec-Regular-Expression-Command.htm
// This will parse a delimited string into an array of
// arrays. The default delimiter is the comma, but this
// can be overriden in the second argument.

function CSVToArray( strData, strDelimiter ){
  // Check to see if the delimiter is defined. If not,
  // then default to comma.
  strDelimiter = (strDelimiter || ",");

  // Create a regular expression to parse the CSV values.
  var objPattern = new RegExp(
    (
      // Delimiters.
      "(\\" + strDelimiter + "|\\r?\\n|\\r|^)" +

      // Quoted fields.
      "(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +

      // Standard fields.
      "([^\"\\" + strDelimiter + "\\r\\n]*))"
    ),
    "gi"
  );


  // Create an array to hold our data. Give the array
  // a default empty first row.
  var arrData = [[]];

  // Create an array to hold our individual pattern
  // matching groups.
  var arrMatches = null;


  // Keep looping over the regular expression matches
  // until we can no longer find a match.
  while (arrMatches = objPattern.exec( strData )){

    // Get the delimiter that was found.
    var strMatchedDelimiter = arrMatches[ 1 ];

    // Check to see if the given delimiter has a length
    // (is not the start of string) and if it matches
    // field delimiter. If id does not, then we know
    // that this delimiter is a row delimiter.
    if (
      strMatchedDelimiter.length &&
      (strMatchedDelimiter != strDelimiter)
    ){

      // Since we have reached a new row of data,
      // add an empty row to our data array.
      arrData.push( [] );

    }


    // Now that we have our delimiter out of the way,
    // let's check to see which kind of value we
    // captured (quoted or unquoted).
    if (arrMatches[ 2 ]){

      // We found a quoted value. When we capture
      // this value, unescape any double quotes.
      var strMatchedValue = arrMatches[ 2 ].replace(
        new RegExp( "\"\"", "g" ),
        "\""
      );

    } else {

      // We found a non-quoted value.
      var strMatchedValue = arrMatches[ 3 ];

    }


    // Now that we have our value string, let's add
    // it to the data array.
    arrData[ arrData.length - 1 ].push( strMatchedValue );
  }

  // Return the parsed data.
  return( arrData );
   
}

function formattable() {
  
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheets()[0];
     
 // Columns start at "1" - this will delete one column starting at 5
 sheet.deleteColumns(5, 1);
  
 //Delete last column 
 var lastColumn = sheet.getLastColumn();
 sheet.deleteColumns(lastColumn, 1); 
 
  // Shifts all columns by one
 sheet.insertColumns(1); 
  
 // Freezes the first row
 sheet.setFrozenRows(1); 
    
 // Compute again size of the table
 var lastRow = sheet.getLastRow();
 var lastColumn = sheet.getLastColumn();
 
 //PROTECMA OR PREVECMA? 
 //If date is previous to 1/5/2010 is PREVECMA, otherwise is PROTECMA 
 var date02 = new Date(2010, 5, 1); 
 //Numbering of rows and columsn starts at 1...index of stored values at 0 
 var values = sheet.getRange(1,6,lastRow).getValues();
 for (var i = 0; i < values.length; i++) {
     //sheet.getRange(i+1,1).setValue(values[i][0]);
   if(values[i][0].valueOf() < date02.valueOf()) {
     sheet.getRange(i+1,1).setValue("PREVECMA");
   } else {
     sheet.getRange(i+1,1).setValue("PROTECMA");
   }  
 }
     
 //CHANGE "undefined" by "blank" cell
 //Numbering of rows and columsn starts at 1...index of stored values at 0 
 var values = sheet.getRange(1,1,lastRow,lastColumn).getValues();
 for (var i = 1; i < lastRow; i++) {
   for (var j = 1; j < lastColumn; j++) {    
      if( values[i][j] === "undefined" ) {
          sheet.getRange(i+1,j+1).setValue("");
      }
   }
 }


 var lastRow = sheet.getLastRow();
 var lastColumn = sheet.getLastColumn(); 
  
 //Insert columns for the 8 Working Groups
 sheet.insertColumns(lastColumn+1,8);

                                              
  //Search key word in the text                                                   
        var lastRow = sheet.getLastRow();
        var lastColumn = sheet.getLastColumn(); 
                                                                   
        var values = sheet.getRange(1,7,lastRow,1).getValues();

        item = 'contingencia';                     
        for(var i = 0; i < values.length; i++) {
            if(values[i].toString().match(item)==item){
               sheet.getRange(i+1,14).setValue("X");           
             }
        }                           
        item = 'operacional';                     
        for(var i = 0; i < values.length; i++) {
            if(values[i].toString().match(item)==item){
               sheet.getRange(i+1,15).setValue("X");           
             }
        }                         
        item = 'residuos';                     
        for(var i = 0; i < values.length; i++) {
            if(values[i].toString().match(item)==item){
               sheet.getRange(i+1,16).setValue("X");           
             }
        }                                                       
        item = 'Directivas';                     
        for(var i = 0; i < values.length; i++) {
            if(values[i].toString().match(item)==item){
               sheet.getRange(i+1,17).setValue("X");           
             }
        }                                       
        item = 'SNPP';                     
        for(var i = 0; i < values.length; i++) {
            if(values[i].toString().match(item)==item){
               sheet.getRange(i+1,18).setValue("X");           
             }
        }                    
        item = 'ambiental';                     
        for(var i = 0; i < values.length; i++) {
            if(values[i].toString().match(item)==item){
               sheet.getRange(i+1,19).setValue("X");           
             }
        }                                                        
        item = 'Dispersantes';                     
        for(var i = 0; i < values.length; i++) {
            if(values[i].toString().match(item)==item){
               sheet.getRange(i+1,20).setValue("X");           
             }
        }                                                     
        item = 'portuaria';                     
        for(var i = 0; i < values.length; i++) {
            if(values[i].toString().match(item)==item){
               sheet.getRange(i+1,21).setValue("X");           
             }
        }
                                  
        //Delete original colum with information about Groups            
        sheet.deleteColumns(7, 1);
        
        //Set new names to colums            
        sheet.getRange(1,1).setValue("PROYECTO");       
        sheet.getRange(1,2).setValue("ID");   
        sheet.getRange(1,3).setValue("Nombre");
        sheet.getRange(1,4).setValue("nick");
        sheet.getRange(1,5).setValue("e-mail");
        sheet.getRange(1,6).setValue("Fecha");
        sheet.getRange(1,7).setValue("Entidad");
        sheet.getRange(1,8).setValue("Cargo");
        sheet.getRange(1,9).setValue("Tlf");
        sheet.getRange(1,10).setValue("Web");
        sheet.getRange(1,11).setValue("Twitter");      
        sheet.getRange(1,12).setValue("Linkedin"); 
  
        sheet.getRange(1,13).setValue("Planes contingencia");
        sheet.getRange(1,14).setValue("Oceanografía Operacional");
        sheet.getRange(1,15).setValue("Gestión residuos"); 
        sheet.getRange(1,16).setValue("Directivas Marco");
        sheet.getRange(1,17).setValue("SNPP");
        sheet.getRange(1,18).setValue("Resturación ambiental");   
        sheet.getRange(1,19).setValue("Dispersantes");                    
        sheet.getRange(1,20).setValue("Gestión acuática");   
 
  
        var lastRow = sheet.getLastRow();
        var lastColumn = sheet.getLastColumn(); 
  
        var cell = sheet.getRange(1,1,1,lastColumn);
        cell.setFontWeight("bold"); 
  
        var cell = sheet.getRange(1,1,1,12);
        cell.setBackgroundRGB(64, 187, 249);
          
        var cell = sheet.getRange(1,13,1,8);
        cell.setBackgroundRGB(140, 219, 79);
                    
        var cell = sheet.getRange(2,13,lastRow-1,8);
        cell.setHorizontalAlignment("center");
        cell.setFontWeight("bold");
        cell.setBorder(true, true, true, true, true, true);
        cell.setBackgroundRGB(227, 253, 207);
                  
               
 }


