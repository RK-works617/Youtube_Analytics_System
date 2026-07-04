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
  Logger.log(data);//JSON構造確認用
  sheet.getRange(2,1).setValue(title);//抽出した情報をスプレッドシートへ転記
  sheet.getRange(2,2).setValue(channel);
  sheet.getRange(2,3).setValue(release);
  sheet.getRange(2,4).setValue(viewcount);
  sheet.getRange(2,5).setValue(videoId);
  
}
