function weeklyNewVideosSearch() {
  let date = new Date();
  date = Utilities.formatDate(date,Session.getScriptTimeZone(),"yyyyMMdd");
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  let ss = sheet.insertSheet(1).setName(date+"週間新着動画");//週間データ保存用リストを作成
  let apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_DATA_API_KEY");//APIキーを取得
  ss.getRange(1,1).setValue("動画タイトル");
  ss.getRange(1,2).setValue("公開日");
  ss.getRange(1,3).setValue("動画ID");
  ss.getRange(1,4).setValue("再生数");
  ss.getRange(1,5).setValue("1日あたり再生数");//先頭行の項目名を記入
  let query = 'ft重音テト';//検索キーワードを指定
  let oneWeekAgo = new Date();
  oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);
  let publishedAfter = oneWeekAgo.toISOString();//日付データをISO8601形式へ変換。
  let url = `https://www.googleapis.com/youtube/v3/search?part=snippet&q=${encodeURIComponent(query)}&type=video&maxResults=50&publishedAfter=${publishedAfter}&order=viewCount&key=${apiKey}`;
  let response = UrlFetchApp.fetch(url);
  let data = JSON.parse(response.getContentText());
  let last = sheet.getLastRow()+2;
  for (let i=0 ; i<data.items.length ; i++){
    let title = data.items[i].snippet.title;
    let release = data.items[i].snippet.publishedAt;
    let videoId = data.items[i].id.videoId;
    ss.getRange(last+i,1).setValue(title);
    ss.getRange(last+i,2).setValue(release);
    ss.getRange(last+i,3).setValue(videoId);
  }
}

function searchViewCounts(){
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  let ss = sheet.getSheets()[1];
  let apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_DATA_API_KEY");
  for (let i=2; i<=ss.getLastRow(); i++){
    let videoId = ss.getRange(i,3).getValue();
    let url = `https://www.googleapis.com/youtube/v3/videos?part=snippet,statistics&id=${videoId}&key=${apiKey}`;
    let response = UrlFetchApp.fetch(url);
    let data = JSON.parse(response.getContentText());
    if (data.items.length === 0) {
      Logger.log("取得失敗：" + videoId);
      continue;
    }
    let viewCount = Number(data.items[0].statistics.viewCount);
    ss.getRange(i,4).setValue(viewCount); 
    date = new Date();
    let serialValue1 = Math.floor((date/86400000)+25569);//情報取得日をシリアル値に変換
    let releaseDate = ss.getRange(i,2).getValue();
    let releaseDate2 = new Date(releaseDate);//公開日情報を文字列からDATEオブジェクトへ変換
    let localTime = releaseDate2.getTime() - (releaseDate2.getTimezoneOffset() * 60000);//タイムゾーンをローカルに変換
    let serialValue2 = Math.floor((localTime/86400000)+25569);//公開日をシリアル値に変換
    let days = serialValue1 - serialValue2;//情報取得日から公開日までの日数を計算
    if(days == 0){
      ss.getRange(i,5).setValue(viewCount);
    }else{
      ss.getRange(i,5).setValue(Math.floor(viewCount/days));//1日あたりの再生数を計算して転記
    }
  }
}

function weeklyUpdate(){
  weeklyNewVideosSearch();
  searchViewCounts();
}
