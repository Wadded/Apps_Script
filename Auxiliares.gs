function onOpen(){
  Criar_Menu()
}

function copiarFormula() {
  for(x=13;x<=24;x++){
    var end = columnToLetter(23+(21*(x-1)))
    var d1 = columnToLetter(19+(21*(x-1)))
    var d2 = columnToLetter(20+(21*(x-1)))

    Diario.getRange("T3:T93").copyTo(Diario.getRange(3,20+(21*(x-1))), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    Diario.getRange("W3:W93").copyTo(Diario.getRange(3,23+(21*(x-1))), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    Diario.getRange(3,20+(21*(x-1)),aMES).setValues(Diario.getRange(3,20+(21*(x-1)),aMES).getValues())
    Diario.getRange(3,19+(21*(x-1)),aMES).setValues(Diario.getRange(3,23+(21*(x-1)),aMES).getValues())
    Diario.getRange(3,19+(21*(x-1))).setValue('F. Emp');
    Diario.getRange(end+':'+end).clearFormat().clear({contentsOnly: true});
    Diario.setColumnWidths(19+(21*(x-1)),2,50);
    Diario.getRange(d1+':'+d2).setNumberFormat('0')
  };
};

function Formatacao_Padrao_Diario(){
  for(x=0;x<=30;x++){
    var end1 = columnToLetter(3+(21*x))
    var end2 = columnToLetter(22+(21*x))
    var end3 = columnToLetter(14+(21*x))
    //var faixa = Diario.getRange(end3+"4:"+end3).getValues().filter(function(r){ return r[0] != ""}).length+4

    if(Diario.getRange(4,3+(21*x)).getBandings().length >0){
      Logger.log(Diario.getRange(4,3+(21*x)).getBandings().length)
      var bd = Diario.getRange(end1+'4:'+end2+'93').getBandings()[0]
      bd.remove()
    }
    Ref.getRange("B2:U"+(aMES+5)).copyTo(Diario.getRange(2,3+(21*x)), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false)
    //Diario.getRange(end1+"4:"+end2 + faixa).sort({column: 3+(21*x), ascending: true});
  };
};

function Marcar_Todos(){
  var fx = Divulga.getRange("B6:B").getValues().filter(function(r){return r[0] != ""}).length+5
  if(Divulga.getRange("C6").getValue() == "NAO"){
    Divulga.getRange("C6:C"+fx).setValue("SIM")
  }else{
    Divulga.getRange("C6:C"+fx).setValue("NAO")
  };
};

function onEdit(e){
  if(e.range.getColumn() == 3 && app.getSheetName() == "Lotes"){
    if(e.value == "ENCERRADO" || e.value == "CANCELADO"){
      var lin = e.range.getRow();
      var l_hist = Hist.getRange("B1000").getNextDataCell(SpreadsheetApp.Direction.UP).getRow()+1;

      Hist.getRange(l_hist,2,1,cLOTES).setValues(Lotes.getRange(lin,2,1,cLOTES).getValues());
      Lotes.getRange(lin,2,1,cLOTES).clearContent();
      Lotes.getRange("L150").autoFill(Lotes.getRange("L6:L2150"), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
      Lotes.getRange("N150:O150").autoFill(Lotes.getRange("N6:O150"), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
      Lotes.getRange("R150").autoFill(Lotes.getRange("R6:R2150"), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    };
    Lotes.getRange("B6:"+clLOTES).sort([{column: 3, ascending: true},{column: 17, ascending: true}]);
  };
};

function tirar_Cor_Alternada(){
  for(x=0;x<=1;x++){
    var faixa = Diario.getRange(end3+"4:"+end3).getValues().filter(function(r){ return r[0] != ""}).length+4
    Logger.log(Diario.getRange(4,3+(21*x)).getBandings())
  }
};

function tirar_FormatCond(){
  var format = Diario.getConditionalFormatRules();
  x=Diario.getConditionalFormatRules().length
  format.splice(1, x);
  Diario.setConditionalFormatRules(format);
};

function tira_DataValid(){
  for(x=1;x<=9;x++){
    Diario.getRange(4,5+(21*(x-1)),90,2).clearDataValidations();
  };
};

function columnToLetter(column){
  var temp, letter = '';
  while (column > 0){
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

