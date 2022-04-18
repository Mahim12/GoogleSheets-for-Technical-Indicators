function onEdit(e)
{
  var row = e.range.getRow();
  var sheet = e.range.getSheet();
  var col = e.range.getColumn();
  
  
  var headerRows = 0;  // # header rows to ignore
  
  // Skip header rows
  if (row <= headerRows) return;
  
  // We're only interested in column F (aka 6)
  if (e.range.getColumn() != 6 ) return;
  {
    var sma1 = [];
    var sma2 = [];
    var std = [];
    var upperbb = [];
    var lowerbb = [];
    var smaformacd1 = [];
    var smaformacd2 = [];
    
    
    var cellref = sheet.getRange("F1").getValue();
    
    var cellref2 = sheet.getRange("F2").getValue();
    
    var diffdate = sheet.getRange("C2").getValue();  
    
    
    for (var i=0; i < diffdate; i++)
    {
      
      sma1[i] = ['=SUM(OFFSET(E' + (i+4).toString() + ',0,0,$F$1,1))/$F$1' ];
      sma2[i] = ['=SUM(OFFSET(E' + (i+4).toString() + ',0,0,$F$2,1))/$F$2' ];
      std[i] = ['=STDEV.P(OFFSET(E' + (i+4).toString() + ',0,0,$F$1,1))' ];
      upperbb[i]= ['=I' + (i+cellref+3).toString() + '+(J'+ (i+cellref+3).toString() + '*$I$1)' ];
      lowerbb[i]= ['=I' + (i+cellref+3).toString() + '-(J'+ (i+cellref+3).toString() + '*$I$1)' ];
      
      
    }
    
    
    var sma1range = sheet.getRange("I4:I");
    sma1range.clearContent();
    var stdrange = sheet.getRange("J4:J");
    stdrange.clearContent();
    var upperbbrange = sheet.getRange("K4:K");
    upperbbrange.clearContent();
    var lowerbbrange = sheet.getRange("l4:l");
    lowerbbrange.clearContent();
    var sma2range = sheet.getRange("m4:m");
    sma2range.clearContent();
    
    
    
    sheet.getRange(cellref+3,9,diffdate,1).setFormulas(sma1);  
    sheet.getRange(cellref+3,10,diffdate,1).setFormulas(std); 
    sheet.getRange(cellref+3,11,diffdate,1).setFormulas(upperbb);
    sheet.getRange(cellref+3,12,diffdate,1).setFormulas(lowerbb);
    sheet.getRange(cellref2+3,13,diffdate,1).setFormulas(sma2);  
    
    
    
    sheet.getRange("I3").setFormula(['IF(ISBLANK(I3),"",CONCAT(F1," SMA"))' ]);
    sheet.getRange("J3").setFormula(['IF(ISBLANK(J3),"","STD DEV")' ]);
    sheet.getRange("K3").setFormula(['IF(ISBLANK(K3),"","UPPER BB")' ]);
    sheet.getRange("L3").setFormula(['IF(ISBLANK(L3),"","LOWER BB")' ]);
    sheet.getRange("M3").setFormula(['IF(ISBLANK(M3),"",CONCAT(F2," SMA"))' ]);  
    
    
    var ema1 = [];
    var ema2 = [];
    var smaofmacdline = [];
    
    var avgupmove = [];
    var avgdownmove = [];
    
    var macdref = sheet.getRange("Q1").getValue();
    
    var macdref2 = sheet.getRange("S1").getValue();
    
    var smaofmacdline2 = sheet.getRange("U1").getValue();
    
    var rsiduration = sheet.getRange("Y1").getValue();
    
    
    
    for (var i=0; i < diffdate; i++)
    {
      
      ema1[i] = ['=SUM(OFFSET(E' + (i+4).toString() + ',0,0,$Q$1,1))/$Q$1' ];
      ema2[i] = ['=SUM(OFFSET(E' + (i+4).toString() + ',0,0,$S$1,1))/$S$1' ];
      smaofmacdline[i] = ['=SUM(OFFSET(T' + (i+macdref2+3).toString() + ',0,0,$U$1,1))/$U$1' ];
      avgupmove[i] = ['=SUM(OFFSET(X' + (i+rsiduration+3).toString() + ',0,0,$Y$1,1))/$Y$1' ];
      avgdownmove[i] = ['=SUM(OFFSET(Y' + (i+rsiduration+3).toString() + ',0,0,$Y$1,1))/$Y$1' ];
      
      
    }
    
    
    var ema1range = sheet.getRange("P4:P");
    ema1range.clearContent();
    
    var ema2range = sheet.getRange("Q4:Q");
    ema2range.clearContent();
    
    var smamacdlinerange = sheet.getRange("U4:U");
    smamacdlinerange.clearContent();
    
    var rsiavgup = sheet.getRange("Z4:Z");
    rsiavgup.clearContent();
    
    var rsiavgdown = sheet.getRange("AA4:AA");
    rsiavgdown.clearContent();
    
    
    
    sheet.getRange(macdref+3,16,diffdate,1).setFormulas(ema1);  
    
    sheet.getRange(macdref2+3,17,diffdate,1).setFormulas(ema2);  
    
    
    sheet.getRange(macdref2+smaofmacdline2+2,21,diffdate,1).setFormulas(smaofmacdline); 
    
    sheet.getRange(rsiduration+4,26,diffdate,1).setFormulas(avgupmove);
    sheet.getRange(rsiduration+4,27,diffdate,1).setFormulas(avgdownmove);
    
    
    
    sheet.getRange("P3").setFormula(['IF(ISBLANK(P3),"",CONCAT(Q1," SMA-MACD"))' ]);
    
    sheet.getRange("Q3").setFormula(['IF(ISBLANK(Q3),"",CONCAT(S1," SMA-MACD"))' ]);  
    
    sheet.getRange("U3").setFormula(['IF(ISBLANK(U3),"",CONCAT(U1," SMA-MACD-LINE"))' ]);  
    
    
    
    
    this.readRows();
    this.onOpen();
    
  }
  
  
  
}



/*** Deletes rows in the active spreadsheet that contain 0 or
* a blank value in column "C". 
* For more information on using the Spreadsheet API, see
* https://developers.google.com/apps-script/service_spreadsheet
*/

function readRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  
  var rowsDeleted = 0;
  for (var i = 0; i <= numRows - 1; i++) {
    var row = values[i];
    if (row[2] == 0 || row[2] == '') {
      sheet.deleteRow((parseInt(i)+1) - rowsDeleted);
      rowsDeleted++;
    }
  }
};

/**
* Adds a custom menu to the active spreadsheet, containing a single menu item
* for invoking the readRows() function specified above.
* The onOpen() function, when defined, is automatically invoked whenever the
* spreadsheet is opened.
* For more information on using the Spreadsheet API, see
* https://developers.google.com/apps-script/service_spreadsheet
*/
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Remove rows where column C and D is 0 or blank",
    functionName : "readRows"
  }];
  sheet.addMenu("Script Center Menu", entries);
};
