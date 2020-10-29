var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheets()[0];

//実行メニューを作成
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu("GAS実行");
  menu.addItem("Googleグループメンバー出力", "exportGoogleGroupMembers");
  menu.addToUi();
}

function exportGoogleGroupMembers() {
  //対象ドメイン
  var domainName = 'geniee.co.jp';
  
  //取得するグループの最大数（デフォルトは200）
  var maxResults = 500;

  //シートをクリア
  sheet.clear();
   
  var values = [];
  
  //ヘッダー追加
  values.push([
    "メールアドレス",
    "グループ名",
    "説明",
    "メンバー数",
    "メンバー",
  ]);
  
  //グループ一覧の取得
  var groupsList = AdminDirectory.Groups.list({domain: domainName, maxResults: maxResults});
   
  if(groupsList) {
    for(var i = 0; i < groupsList.groups.length; i++){
      var value = [];
      
      //グループの基本情報を取得
      value.push(groupsList.groups[i].email); //メールアドレス
      value.push(groupsList.groups[i].name); //グループ名
      value.push(groupsList.groups[i].description); //説明
      value.push(groupsList.groups[i].directMembersCount); //メンバー数

      var strMembers = '';

      //グループのメンバーを取得
      var members = AdminDirectory.Members.list(groupsList.groups[i].email).members;
      
      if(members != null)
      {
        for (var j = 0; j < members.length; j++){
          //社外ドメインのアドレスの場合、ドメイン部以外はマスクする
          if(members[j].email.match(/@geniee.co.jp/)){
            strMembers += members[j].email + '\r\n';
          }else{
            var splitEmail = members[j].email.split('@');
            strMembers += '*****@' + splitEmail[1] + '\r\n';
          }
        }
      }
      value.push(strMembers.trim());

      values.push(value); 
    }
    
    //取得したデータをスプレッドシートにセット
    sheet.getRange(1, 1, groupsList.groups.length + 1 , 5).setValues(values);
    sheet.getRange(1, 1, groupsList.groups.length + 1 , 5).setVerticalAlignment('top')
  } 
}
