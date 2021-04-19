const SLACK_API_TOKEN         = "ここにSlackAPIトークンを入力";
const SLACK_CHANNEL_ID        = "ここにSlackチャンネルIDを入力"; // パブリックチャンネルである必要がある
const SLACK_WORKSPACE_NAME    = "ここにSlackワークスペース名を入力";
const SLACK_HISTORY_GET_COUNT = 100; // 1度に取得するメッセージ数

function getSlackHistory(){
  var url = "https://slack.com/api/conversations.history";
  var payload = {
    "token": SLACK_API_TOKEN,
    "channel": SLACK_CHANNEL_ID,
    "count" : SLACK_HISTORY_GET_COUNT
  };
  var options = {
    "headers": {"Accept": "application/json"},
    "payload" : payload
  };
  
  var strResponse = UrlFetchApp.fetch(url, options);
  json = strResponse.getContentText();
  if (json == "") {
    return;
  }
  // Logger.log(json);
  
  var data = JSON.parse(json).messages;
  return data;
}

function getSlackMembers(){
  var url = "https://slack.com/api/users.list";
  var payload = {
    "token": SLACK_API_TOKEN,
  };
  var options = {
    "headers": {"Accept": "application/json"},
    "payload" : payload
  };
  
  var strResponse = UrlFetchApp.fetch(url, options);
  json = strResponse.getContentText();
  if (json == "") {
    return;
  }
  // Logger.log(json);
  
  var data = JSON.parse(json).members;
  return data;
}

function getSlackMemberName(slack_members, user_id){
  for (var i = 0; i < slack_members.length; i++)
  {
    var member = slack_members[i];
    var id = member.id;
    
    if (id == user_id) return member.profile.display_name;
  }
}

function getSlackMessageLink(timestamp, thread_timestamp) {
  var base_url = "https://" + SLACK_WORKSPACE_NAME + ".slack.com/archives/";
  var p_url = "p" + timestamp.replace('.', '');
  var url = base_url + SLACK_CHANNEL_ID + "/" + p_url;
  
  if (thread_timestamp != null)
  {
    url = url + "?thread_ts=" + thread_timestamp;
  }
  
  return url;
}