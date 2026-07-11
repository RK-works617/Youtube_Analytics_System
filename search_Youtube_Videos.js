function fetchAndWriteYoutubeVideos() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_DATA_API_KEY");//APIキーを取得
  let videoId = '4aFC2oC-wHA'; // 取得したい動画のIDをここに入力
  let url = `https://www.googleapis.com/youtube/v3/videos?part=snippet,statistics&id=${videoId}&key=${apiKey}`;//YouTube Data APIへリクエスト
  let response = UrlFetchApp.fetch(url);
  let data = JSON.parse(response.getContentText());//JSONレスポンスをJavascriptオブジェクトへ変換
  let title = data.items[0].snippet.title;//必要な情報を抽出
  let release = data.items[0].snippet.publishedAt;
  let channel = data.items[0].snippet.channelTitle;
  let viewcount = data.items[0].statistics.viewCount;
  // Logger.log(data);//JSON構造確認用
  sheet.getRange(2,1).setValue(title);//抽出した情報をスプレッドシートへ転記
  sheet.getRange(2,2).setValue(channel);
  sheet.getRange(2,3).setValue(release);
  sheet.getRange(2,4).setValue(viewcount);
  sheet.getRange(2,5).setValue(videoId);
  
}

function searchYoutubeVideos(){
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_DATA_API_KEY");//APIキーを取得
  let query = 'ft重音テト';//検索キーワードを指定
  let url = `https://www.googleapis.com/youtube/v3/search?part=snippet&q=${encodeURIComponent(query)}&type=video&maxResults=50&key=${apiKey}`;//５０件の検索結果を取得
  let response = UrlFetchApp.fetch(url);
  let data = JSON.parse(response.getContentText());
  let last = sheet.getLastRow()+1;
  //５０件の動画データを取得して、スプレッドシートに転記
  for (i=0 ; i<data.items.length ; i++){
    let title = data.items[i].snippet.title;
    let channel = data.items[i].snippet.channelTitle;
    let release = data.items[i].snippet.publishedAt;
    let videoId = data.items[i].id.videoId;
    sheet.getRange(last+i,1).setValue(title);
    sheet.getRange(last+i,2).setValue(channel);
    sheet.getRange(last+i,3).setValue(release);
    sheet.getRange(last+i,4).setValue(videoId);
  }
  
}

function searchCountviews(){
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  for (let i=2; i<=sheet.getLastRow(); i++){
    let videoId = sheet.getRange(i,4).getValue();//取得した動画の動画IDを順次取得
    let apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_DATA_API_KEY")
    let url = `https://www.googleapis.com/youtube/v3/videos?part=snippet,statistics&id=${videoId}&key=${apiKey}`;
    let response = UrlFetchApp.fetch(url);
    let data = JSON.parse(response.getContentText());
    let viewCount = data.items[0].statistics.viewCount//再生回数を取得
    sheet.getRange(i,5).setValue(viewCount);//取得した再生回数を転記
  }
  
}
