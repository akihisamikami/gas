function onOpen() {
  var myMenu=[
    {name:"集積開始",functionName:"main"}
  ];
  
  SpreadsheetApp.getActiveSpreadsheet().addMenu("集積",myMenu);
}

function main() {
  var slack_members = getSlackMembers();
  var logs = getSlackHistory();
  var tags = getTags();  
  var messages = [];

  for (var i = 0; i < logs.length; i++) {
    var data = logs[i];
    var body = data.text;
    var ts = data.ts;
    var thread_ts = data.thread_ts;
    var date = getFormattedDate(ts)
    var user_id = data.user;
    var user_name = getSlackMemberName(slack_members, user_id);
    var message_link = getSlackMessageLink(ts, thread_ts);
    var tag = getTag(body, tags);
    
    // タグをbodyから除外
    body = body.replace(tag, "");
    
    if (tag) {
      var json = {
        "date": date,
        "tag": tag,
        "user_name": user_name,
        "body" : body,
        "message_link": message_link,
        "timestamp": ts
      };
      
      messages.push(json);
    }
  }
  
  updateHistoryInSpreadsheet(messages);
}

function getTag(message, tags){
  for (var i = 0; i < tags.length; i++)
  {
    var tag = tags[i];
    var regexp = new RegExp('^' + tag + '(.*)');
    
    if (message.match(regexp)) {
      return tag;
    }
  }
  
  return null;
}

function getFormattedDate(timestamp) {
  var d = new Date( timestamp * 1000 );
  var year  = d.getFullYear();
  var month = d.getMonth() + 1;
  var day = d.getDate();
  var dayOfWeek = d.getDay();
  var dayOfWeekStr = [ "日", "月", "火", "水", "木", "金", "土" ][dayOfWeek];
  return year + '/' + month + '/' + day + ' ' + dayOfWeekStr;
}