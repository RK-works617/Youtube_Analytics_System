function highlightTop10(){
  let sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート２");
  let lastc = sheet2.getLastColumn();
  let lastr = sheet2.getLastRow();
  let numbers = sheet2.getRange(2,lastc,lastr-1,1).getValues().sort((a,b)=>b-a);
  let numbers1d = numbers.flat();
  for (let i=2; i<= lastr; i++){
    let cell = sheet2.getRange(i,lastc).getValue();
    if (cell == numbers1d[0]){
      sheet2.getRange(i,lastc).setBackground("#FFD700");
    }
    else if (cell == numbers1d[1]){
      sheet2.getRange(i,lastc).setBackground("#C0C0C0");
    }
    else if (cell == numbers1d[2]){
      sheet2.getRange(i,lastc).setBackground("#CD7F32");
    }
    else if (numbers1d.slice(3,10).includes(cell)){
      sheet2.getRange(i,lastc).setBackground("#D9EAF7");
    }  
  }
}
