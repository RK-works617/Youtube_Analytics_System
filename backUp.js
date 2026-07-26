function backUp() {
  let date = new Date();
  let date2 = Utilities.formatDate(date,Session.getScriptTimeZone(),"yyyyMMdd");
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  let backUp1 = sheet.insertSheet().setName(date2+"シート1");//シート1のバックアップ用シートを作成
  let backUp2 = sheet.insertSheet().setName(date2+"シート２");//シート2のバックアップ用シートを作成
  let sheet2 = SpreadsheetApp.openById("1RRvnbNPhzrSDOtBIA233D7rtGmGPkeuyJat_1Ov2ohE");//元データのシートを呼び出す
  let original1 = sheet2.getSheetByName("シート1");
  let original2 = sheet2.getSheetByName("シート２");
  let lastr1 = original1.getLastRow();
  let lastc1 = original1.getLastColumn();
  let lastr2 = original2.getLastRow();
  let lastc2 = original2.getLastColumn();
  let values1 = original1.getRange(1,1,lastr1,lastc1).getValues();//バックアップ対象のデータを取得
  let values2 = original2.getRange(1,1,lastr2,1).getValues();
  let values3 = original2.getRange(1,lastc2-6,lastr2,7).getValues();
  backUp1.getRange(1,1,lastr1,lastc1).setValues(values1);//バックアップシートに書き込み
  backUp2.getRange(1,1,lastr2,1).setValues(values2);
  backUp2.getRange(1,2,lastr2,7).setValues(values3);
  backUp2.getRange(1,9).setValue("週間再生数");//バックアップシート最終行に゙週間再生数を表示
  for (let i=2; i<=lastr2; i++){
      let weekly = 0;
      let days = backUp2.getRange(i,2,1,7).getValues();//楽曲ごとの日次再生数を配列で取得
    for (let k=0; k<7; k++){
      weekly += days[0][k];// 週間再生数として合計
      backUp2.getRange(i,9).setValue(weekly);//最終列に゙週間再生数として転記
    }
  }
}
