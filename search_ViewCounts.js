function searchViewCounts(){
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート1");
  let sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート２");
  let apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_DATA_API_KEY");
  let values = sheet.getRange(2,5,sheet.getLastRow()-1,1).getValues();
  sheet.getRange(2,6,values.length,1).setValues(values);//前回再生数を隣接行へ転記
  for (let i=2; i<=sheet.getLastRow(); i++){
    let videoId = sheet.getRange(i,4).getValue();
    let url = `https://www.googleapis.com/youtube/v3/videos?part=snippet,statistics&id=${videoId}&key=${apiKey}`;
    let response = UrlFetchApp.fetch(url);
    let data = JSON.parse(response.getContentText());
    if (data.items.length === 0) {//動画データ取得失敗時に処理継続
      Logger.log("取得失敗：" + videoId);
      continue;
    }
    let viewCount = Number(data.items[0].statistics.viewCount);
    sheet.getRange(i,5).setValue(viewCount);
    let weekly = viewCount-Number(sheet.getRange(i,6).getValue());//前回再生数との差分を計算
    sheet.getRange(i,7).setValue(weekly);//差分をスプレッドシートへ転記
  }
  let inc = sheet.getRange(2,7,sheet.getLastRow()-1,1).getValues();//保存用シートに転記する差分データを取得
  let lastc = sheet2.getLastColumn()+1;
  sheet2.getRange(2,lastc,inc.length,1).setValues(inc);//空白列に最新の差分データを転記
  let date = new Date();
  let date2 = Utilities.formatDate(date,Session.getScriptTimeZone(),"yyyy/MM/dd")
  sheet2.getRange(1,lastc).setValue(date2);//先頭行にデータ取得日を転記
}
