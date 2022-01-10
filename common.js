//=================================================================|1
//							関数の実行確認
//========================================================|2021.12.31
// v1: 関数の作成
/**
 * 関数実行前の確認を行う
 * @return {boolean}
 */
function confirmExecFunction() {
  let msg = Browser.msgBox(confirmExecFunctionTitle, confirmExecFunctionMsg, Browser.Buttons.OK_CANCEL);
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
/**
 * 指定したスクリプトプロパティを返す。
 * 第二引数で値が見つからない場合の挙動を指定する。
 * @param {Array} keys 
 * @param {boolean} bool 
 * @returns 
 */
function getProps(keys, bool) {
  const env = PropertiesService.getScriptProperties();
  const properties = {};
  if (keys.includes('xApiKey')) {
    properties.xApiKey = env.getProperty('API_KEY');
    if(!properties.xApiKey && bool) {
      const notFoundKey = 'API_KEY';
      Browser.msgBox(notFoundKey + getPropsMsg1 + notFoundKey + getPropsMsg2);
      return
    }
  }
  if (keys.includes('xToken')) {
    properties.xToken = env.getProperty('API_TOKEN');
    if(!properties.xToken && bool) {
      const notFoundKey = 'API_TOKEN';
      Browser.msgBox(notFoundKey + getPropsMsg1 + notFoundKey + getPropsMsg2);
      return
    }
  }
  if (keys.includes('appKey')) {
    properties.appKey = env.getProperty('APP_KEY');
    if(!properties.appKey && bool) {
      const notFoundKey = 'APP_KEY';
      Browser.msgBox(notFoundKey + getPropsMsg1 + notFoundKey + getPropsMsg2);
      return
    }
  }
  if (keys.includes('saveCode')) {
    properties.saveCode = env.getProperty('SAVE_CODE');
    if(!properties.saveCode && bool) {
      const notFoundKey = 'SAVE_CODE';
      Browser.msgBox(notFoundKey + getPropsMsg1 + notFoundKey + getPropsMsg2);
      return
    }
  }
  if (keys.includes('folderId')) {
    properties.folderId = env.getProperty('FOLDER_ID');
    if(!properties.folderId && bool) {
      const notFoundKey = 'FOLDER_ID';
      Browser.msgBox(notFoundKey + getPropsMsg1 + notFoundKey + getPropsMsg2);
      return
    }
  }
  return properties
}


//=================================================================|1
//							作業中スプレッドシートの最初のシートを取得
//========================================================|2021.12.31
// v1: 関数の作成
/**
 * スプレッドシートの最初のシートを取得して返す。
 * @returns 
 */
 function setFirstSheet() {
  let firstSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  return firstSheet;
}


//=================================================================|1
//							レコードデータのCSVバイナリの作成
//========================================================|2021.12.31
// v1: 関数の作成
/**
 * 指定したシートのCSVバイナリを作成して返す
 * @param {*} sheet 
 * @returns 
 */
function convertBlob(sheet) {
  let records = sheet.getDataRange().getValues();
  let csv = records.join('\n');
  let blob = Utilities.newBlob(csv, MimeType.CSV, 'records.csv');
  return blob;
}


//=================================================================|1
//							関数実行ログの出力
//========================================================|2021.12.31
// v1: 関数の作成
/**
 * 関数実行した日時・関数名・実行者のログをlogシートに出力する
 * @param {string} funcName 
 */
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
/**
 * 実行ユーザーのIDを返す。
 * @returns 
 */
function getUserId() {
  const userId = Session.getActiveUser().getUserLoginId();
  return userId;
}
