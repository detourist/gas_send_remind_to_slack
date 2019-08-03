//【Ref】https://qiita.com/KanaSakaguchi/items/f0b1bb1cf73f0ec5ec71
//【Ref】https://qiita.com/will_meaning/items/2714ada180f76650a92e
//【Ref】https://qiita.com/ykhirao/items/782e20ab0465533c48f6

function sendRemindToSlack() {
  //::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
  //日本の祝日カレンダー
  const calHoliday = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com')
  //Slack投稿時の名前
  const botName ="リマインドbot";
  //デバッグモード
  const debug = false;
  
  //::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
  //GASスクリプトプロパティ
  const scriptProperty = PropertiesService.getScriptProperties().getProperties();
  //Slack APIトークン
  const token = scriptProperty.SLACK_TOKEN;
    
  //::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
  //現在時刻
  const now = new Date();
  if(debug) Logger.log("now : "+ String(now));
  //マイナス10分
  const marginTime = new Date(now.getFullYear(), now.getMonth(), now.getDate(), now.getHours(), now.getMinutes()-10, now.getSeconds())
  if(debug) Logger.log("marginTime : "+ String(marginTime));
  //今日
  var date = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  if(debug) Logger.log("date : "+ String(date));
  //曜日
  const week = date.getDay()
  //曜日を漢字に変換
  const weekChars = ["日","月","火","水","木","金","土"];
  const weekChar = weekChars[week];
  
  //今月の1日
  const firstDate = new Date(date.getFullYear(), date.getMonth(), 1);
  if(debug) Logger.log("firstDate : "+ String(firstDate));
  //今月の何週目かを計算
  //7day*24hour*60min*60sec*1000ms = 604800000ミリ秒
  const weekNum = Math.floor((date - firstDate)/604800000) + 1;
  
  //来月の1日
  const nextFirstDate = new Date(date.getFullYear(), date.getMonth()+1, 1);
  if(debug) Logger.log("nextFirstDate : "+ String(nextFirstDate));
  //今月の月末からマイナス何週目かを計算（月末日 = -1週）
  //7day*24hour*60min*60sec*1000ms = 604800000ミリ秒
  const weekNumMinus = Math.floor((date - nextFirstDate)/604800000);

  //月初営業日
  //1日から辿って、祝日を除いた最初の営業日を、「月初営業日」と定義
  //今月1日を初期値とする
  const firstBizDate =new Date(date.getFullYear(), date.getMonth(), 1);
  //1日から今日まで繰返し
  for(var i = 1; i <= date.getDate(); i++){
    //土日でなく、祝日でない（祝日カレンダーで今日のEventが0件である）ならば、月初営業日となる
    if(firstBizDate.getDay()!=0 && firstBizDate.getDay()!=6 && calHoliday.getEventsForDay(firstBizDate).length==0){
      //条件を満たすとき、for処理を抜ける
      break;
    }
    //翌日にインクリメント
    firstBizDate.setDate(firstBizDate.getDate() + 1);
  }
  if(debug) Logger.log("firstBizDate : "+ String(firstBizDate));

  //月末営業日
  //月末から遡って、祝日を除いた最後の営業日を、「月末営業日」と定義
  //今月末日を初期値とする
  const lastBizDate = new Date(date.getFullYear(), date.getMonth()+1, 0);
  //末日から今日まで遡り繰返し
  for(var i = lastBizDate.getDate(); i >= date.getDate(); i--){
    //土日でなく、祝日でない（祝日カレンダーで今日のEventが0件である）ならば、週末営業日となる
    if(lastBizDate.getDay()!=0 && lastBizDate.getDay()!=6 && calHoliday.getEventsForDay(lastBizDate).length==0){
      //条件を満たすとき、for処理を抜ける
      break;
    }
    //前日にデクリメント
    lastBizDate.setDate(lastBizDate.getDate() - 1);
  }
  if(debug) Logger.log("lastBizDate : "+ String(lastBizDate));

  //::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
  //スプレッドシート
  var sheetName = "【Sheet】";
  if(debug) sheetName = "【設定例】";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const lastRow = sheet.getDataRange().getLastRow();

  //テーブルを取得
  const items = sheet.getRange(2, 1, lastRow -1, 5).getValues();
  if(debug) Logger.log(items);
    
  //時刻判定用の変数
  var targetTime = new Date(now.getFullYear(), now.getMonth(), now.getDate(), now.getHours(), now.getMinutes(), now.getSeconds());

  //spliceを使用するためにデクリメントで処理する。lengthの値は更新されないため。
  for(var i = items.length - 1; i >= 0; i--){
    //スプレッドシートに入力されている時刻を反映
    targetTime.setHours(items[i][2].getHours());
    targetTime.setMinutes(items[i][2].getMinutes());
    if(debug) Logger.log("targetTime : "+ String(targetTime));
    
    //weekCharを検索する正規表現
    var regChar = new RegExp(weekChar);
    //weekNumMinusを検索する正規表現
    var regMinus = new RegExp(weekNumMinus);
    //weekNumの直前にマイナスが付いていない条件で検索する正規表現
    var regNotMinus = new RegExp("(^|[^-])"+ weekNum);

    //条件判定
    //テーブルの配列のうち、条件に条件に一致するもののみを配列に残す（条件に一致しないものを配列から除外する）
    //【条件】
    //週が"0"(毎週)であり、曜日がweekCharを含む
    //または　週がweekNumMinusを含み、曜日がweekCharを含む
    //または　週がweekNumを含み（週が「-1」のときに「1」でヒットしないよう「weekNumの直前にマイナスが付いていない」条件で検索する）、曜日がweekCharを含む
    //または　週が"月"、曜日が"初"であり、今日の日付がfirstBizDateと一致する
    //または　週が"月"、曜日が"末"であり、今日の日付がlastBizDateと一致する
    //そのいずれにおいても　時刻が現在（Script実行）時刻からマイナス10分以内である
    if(
      (
        (items[i][0]=="0" && regChar.test(items[i][1]))
        || (regMinus.test(items[i][0]) && regChar.test(items[i][1]))
        || (regNotMinus.test(items[i][0]) && regChar.test(items[i][1]))
        || (items[i][0]=="月" && items[i][1]=="初" && date.getTime() == firstBizDate.getTime()) 
        || (items[i][0]=="月" && items[i][1]=="末" && date.getTime() == lastBizDate.getTime())
        //JavaScript の日付を比較するときに、演算子の両側の日付が同じオブジェクトを参照している場合にだけ == 演算子が true を返すことを念頭に置く必要があります。 
        //したがって、2 つの別々の Date オブジェクトが同じ日付に設定されている場合、date1 == date2 は false を返します。 
      )
      && (targetTime > marginTime && targetTime <= now)
    ){
      //条件に一致するので何もしない
    }else{
      //条件に一致しないレコードを配列から除外する
      items.splice(i, 1);
    }
  } 
  if(debug) Logger.log(items);
  
  //::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
  //Slackメッセージを送信  
  if(items.length > 0){
    for(var i = 0; i < items.length; i++){
      var options = 
        {
          "method" : "POST",
          "payload" : 
          {
            "token": token,
            "username" : botName,
            "icon_emoji" : ":mega:",
            "channel": items[i][3],
            "text": items[i][4]
          }
        }
      
      //投稿
      var response = UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", options);
      if(debug) Logger.log(response);
    }
  }
  //::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
}
