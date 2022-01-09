//=================================================================|1
//							関数の実行確認
//========================================================|2021.12.31
// v1: 関数の作成
function confirmExecFunction() {
  let msg = Browser.msgBox('実行確認','関数を実行します。よろしいですか？',Browser.Buttons.OK_CANCEL);
  if (msg == 'ok') {
    return true;
  } else {
    return false;
  }
}


//=================================================================|3
//							スクリプトプロパティの取得
//========================================================|2022.01.05
// v1: 関数の作成
// v2: keysを配列で受け取り、それに対応するプロパティを取得する仕様に変更
// v3: 第二引数がtrueの場合のみ、値が見つからなければ警告を発生する仕様に変更
function getProps(keys, bool) {
  const env = PropertiesService.getScriptProperties();
  const properties = {};
  if (keys.includes('xApiKey')) {
    properties.xApiKey = env.getProperty('API_KEY');
    if(!properties.xApiKey && bool) {
      Browser.msgBox('APIキーが見つかりませんでした。プロパティストアにAPI_KEYを登録しているか確認して下さい。');
      return
    }
  }
  if (keys.includes('xToken')) {
    properties.xToken = env.getProperty('API_TOKEN');
    if(!properties.xToken && bool) {
      Browser.msgBox('APIトークンが見つかりませんでした。プロパティストアにAPI_TOKENを登録しているか確認して下さい。');
      return
    }
  }
  if (keys.includes('appKey')) {
    properties.appKey = env.getProperty('APP_KEY');
    if(!properties.appKey && bool) {
      Browser.msgBox('アプリキーが見つかりませんでした。プロパティストアにAPP_KEYを登録しているか確認して下さい。');
      return
    }
  }
  if (keys.includes('saveCode')) {
    properties.saveCode = env.getProperty('SAVE_CODE');
    if(!properties.saveCode && bool) {
      Browser.msgBox('識別キーが見つかりませんでした。プロパティストアにSAVE_CODEを登録しているか確認して下さい。');
      return
    }
  }
  if (keys.includes('folderId')) {
    properties.folderId = env.getProperty('FOLDER_ID');
    if(!properties.folderId && bool) {
      Browser.msgBox('フォルダIDの取得に失敗しました。プロパティストアにFOLDER_IDを登録しているか確認して下さい。');
      return
    }
  }
  return properties
}


//=================================================================|1
//							レコードデータのCSVバイナリの作成
//========================================================|2021.12.31
// v1: 関数の作成
function convertBlob(activeSheet) {
  let records = activeSheet.getDataRange().getValues();
  let csv = records.join('\n');
  let blob = Utilities.newBlob(csv, MimeType.CSV, 'records.csv');
  return blob;
}


//=================================================================|1
//							関数実行ログの出力
//========================================================|2021.12.31
// v1: 関数の作成
function logOutput(funcName) {
  //----------------------------------------
  // ログシートの取得
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('log');

  //----------------------------------------
  // ログシートがない場合作成する
  if(!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    sheet.setName('log');
    const columns = ['実行日時', '実行関数', '実行ユーザー'];
    sheet.getRange(1, 1, 1, columns.length).setValues([columns]);
  }

  //----------------------------------------
  // 実行日時の取得
  const date = new Date();
  const execDate = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd hh:mm:ss');

  //----------------------------------------
  // 実行ユーザーIDの取得
  const userName = getUserId();

  //----------------------------------------
  // ログの出力
  const logData = [execDate, funcName, userName];
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, logData.length).setValues([logData]);
}


//=================================================================|1
//							実行ユーザー名の取得
//========================================================|2021.12.31
// v1: 関数の作成
function getUserId() {
  const userId = Session.getActiveUser().getUserLoginId();
  return userId;
}
