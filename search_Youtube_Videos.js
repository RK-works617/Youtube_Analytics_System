function searchYoutubeVideos(){
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート1");
  let apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_DATA_API_KEY");//APIキーを取得
  let query = 'ft重音テト';//検索キーワードを指定
  let url = `https://www.googleapis.com/youtube/v3/search?part=snippet&q=${encodeURIComponent(query)}&type=video&maxResults=50&key=${apiKey}`;
  let response = UrlFetchApp.fetch(url);
  let data = JSON.parse(response.getContentText());
  let last = sheet.getLastRow()+1;
  for (let i=0 ; i<data.items.length ; i++){
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
