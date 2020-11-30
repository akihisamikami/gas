const SPREAD_SHEET_ID          = 'ここにGoogle SpreadsheetのIDを入力';
const SHEET_NAME_ACCUMULATION  = '集積'; // 集積に使うシート名
const SHEET_NAME_TAG           = 'タグマスタ'; // タグを定義するシート名

const FIRST_ROW_NUMBER         = 4; // 入力開始行番号
const COLUMN_NO                = 1; // 入力列番号．1=A列
const COLUMN_DATE              = 2;
const COLUMN_TAG               = 3;
const COLUMN_USER_NAME         = 4;
const COLUMN_BODY              = 5;
const COLUMN_MESSAGE_LINK      = 6;
const COLUMN_TIMESTAMP         = 7;

function getTags() {
  var spreadsheet = SpreadsheetApp.openById(SPREAD_SHEET_ID);
  var sheet = spreadsheet.getSheetByName(SHEET_NAME_TAG);
  
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues();
  return data;
}

function updateHistoryInSpreadsheet(messages) {
  var spreadsheet = SpreadsheetApp.openById(SPREAD_SHEET_ID);
  var sheet = spreadsheet.getSheetByName(SHEET_NAME_ACCUMULATION);
  
  // Logger.log(messages);

  // 最新行数
  var latest_row_number = sheet.getLastRow();
  var last_timestamp = getLatestTimestamp(sheet);

  // 書き出し始めの行にするため1追加
  latest_row_number = latest_row_number + 1;

  // 古いメッセージから出力
  for (var i = messages.length - 1; i >= 0; i--) {
    var message = messages[i];
    var this_ts = parseFloat(message.timestamp);
    var last_ts = parseFloat(last_timestamp);
    var no = latest_row_number - FIRST_ROW_NUMBER + 1;
    
    if (this_ts > last_ts && latest_row_number >= FIRST_ROW_NUMBER) {

      setValue(sheet, latest_row_number, COLUMN_NO, no);
      setValue(sheet, latest_row_number, COLUMN_DATE, message.date);
      setValue(sheet, latest_row_number, COLUMN_TAG, message.tag);
      setValue(sheet, latest_row_number, COLUMN_USER_NAME, message.user_name);
      setValue(sheet, latest_row_number, COLUMN_BODY, message.body);
      setValue(sheet, latest_row_number, COLUMN_MESSAGE_LINK, message.message_link);
      setValue(sheet, latest_row_number, COLUMN_TIMESTAMP, message.timestamp);

      latest_row_number++;
    }
  }
}

function getLatestTimestamp (sheet) {
  // タイムスタンプ列から取得
  var vals = sheet.getRange(FIRST_ROW_NUMBER, COLUMN_TIMESTAMP, sheet.getLastRow() - 1).getValues().filter(String);
  var last_timestamp = vals.pop();
  
  if(!Number(last_timestamp)) return 0;
  
  return last_timestamp;
}

function setValue(sheet, row_number, column_number, value) {
  range = sheet.getRange(row_number, column_number);
  return range.setValue(value);
}