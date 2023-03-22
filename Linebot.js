// 設定 LINE Bot 資訊和 Google Sheets 資訊
var LINE_CHANNEL_ACCESS_TOKEN = "your line channel access token";
var LINE_CHANNEL_SECRET = "your line user ID";
var SPREADSHEET_ID = "your google sheet id";
var SHEET_NAME = "your google sheet table name";

function doPost(e) {
  var json = JSON.parse(e.postData.contents);
  var replyToken = json.events[0].replyToken;
  var message = json.events[0].message.text;
  var userId = json.events[0].source.userId;
  var timestamp = json.events[0].timestamp;

  // 判斷使用者發送的訊息是否為 "上班" 或 "下班"
  if (message === "上班" || message === "下班") {
    var date = new Date(timestamp);
    // var currentDate = formatDate(date);
    var currentTime = formatTime(date);

    // 判斷使用者打卡時間是否在上班時間內
      var username = getUsername(userId);
      var data = {
        replyToken: replyToken,
        username: username,
        userId: userId,
        // date: currentDate,
        timestamp: currentTime,
        situation: message
      };
      if(message === "上班"){
        writeDataToSheet(data);
      }else if(message === "下班"){
        writeDataToSheet1(data);
      }
      
  }else if(message === "test"){
    pushLocationRequest();
  }
  
  return ContentService.createTextOutput(JSON.stringify({"status": 200}));
}

// 根據 userId 取得使用者名稱
function getUsername(userId) {
  var url = "https://api.line.me/v2/bot/profile/" + userId;
  var headers = {
    "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
  };
  var options = {
    "method": "GET",
    "headers": headers
  };
  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());
  return json.displayName;
}

// 將打卡資料寫入 Google Sheets 中(上班)
function writeDataToSheet(data) {
  var replyToken = data.replyToken;
  var SpreadSheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  var Sheet = SpreadSheet.getSheetByName(SHEET_NAME);
  var LastRow = Sheet.getLastRow();
  var dataRange = Sheet.getDataRange();
  var values = dataRange.getValues();
  var numRows = values.length;
  var idColumn = 0;
  var situationColumn = 4;

  var id = data.userId;
  var name = data.username;
  var time = data.timestamp;
  var situation = data.situation;

  var count = 0;
  var count1 = 0;
  for (var i = 0; i < numRows; i++) {
    var row = values[i];
    if (row[idColumn] == id) {
      count += 1;
    }
    if (row[idColumn] == id && row[situationColumn] == "下班") {
      count1 += 1;
    }
  }

  if(count == count1){
    Sheet.getRange(LastRow+1,1).setValue(id);
    Sheet.getRange(LastRow+1,2).setValue(name);
    Sheet.getRange(LastRow+1,3).setValue(time);
    Sheet.getRange(LastRow+1,5).setValue(situation);
    replyMessage(replyToken, "打卡成功！" + name + time);
  }else{
    replyMessage(replyToken, "請先打卡下班！");
  }
}

// 將打卡資料寫入 Google Sheets 中(下班)
function writeDataToSheet1(data) {
  var replyToken = data.replyToken;
  var SpreadSheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  var Sheet = SpreadSheet.getSheetByName(SHEET_NAME);
  var dataRange = Sheet.getDataRange();
  var values = dataRange.getValues();
  var numRows = values.length;
  var idColumn = 0;
  var situationColumn = 4;

  var id = data.userId;
  var time = data.timestamp;
  var situation = data.situation;

  var count = 0;
  var count1 = 0;
  for (var i = 0; i < numRows; i++) {
    var row = values[i];
    if (row[idColumn] == id) {
      count += 1;
    }
    if (row[idColumn] == id && row[situationColumn] == "下班") {
      count1 += 1;
    }
  }

  if(count == count1){
    replyMessage(replyToken, "請先打卡上班！");
  }else{
    for (var i = 0; i < numRows; i++) {
      var row = values[i];
      if (row[idColumn] == id && row[situationColumn] == "上班") {
        Sheet.getRange(i+1, 5).setValue(situation);
        Sheet.getRange(i+1, 4).setValue(time);
      }
    }
    replyMessage(replyToken, "打卡成功！" + data.username + time);
  }
}

// 回覆訊息給使用者
function replyMessage(replyToken, message) {
  var url = "https://api.line.me/v2/bot/message/reply";
  var headers = {
    "Content-Type": "application/json",
    "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
  };
  var data = {
    "replyToken": replyToken,
    "messages": [{
      "type": "text",
      "text": message
    }]
    };

  var options = {
    "method": "POST",
    "headers": headers,
    "payload": JSON.stringify(data)
    };
    var response = UrlFetchApp.fetch(url, options);
    Logger.log(response);
  }

  // 格式化時間為 "YYYY/MM/DD HH:mm" 的字串
  function formatTime(date) {
    var year = date.getFullYear();
    var month = addLeadingZero(date.getMonth() + 1);
    var day = addLeadingZero(date.getDate());
    var hours = addLeadingZero(date.getHours());
    var minutes = addLeadingZero(date.getMinutes());
    return year + "/" + month + "/" + day + " " + hours + ":" + minutes;
  }

  // 在小於 10 的數字前面加上 "0"
  function addLeadingZero(number) {
  return number < 10 ? "0" + number : number;
  }
  
