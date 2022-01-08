//=================================================================|2
//							スプレッドシートを開いた際に実行
//========================================================|2022.01.03
// v1: 関数の作成
// v2: 初期設定登録フォームの実装
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('サスケWorks');
  menu.addSubMenu(ui.createMenu('関数')
    .addItem('CSV出力', 'csvSaveToDrive')
    .addItem('一括登録', 'apiImportPost')
  );
  menu.addSeparator();
  menu.addItem('API設定', 'showApiConf');
  menu.addToUi();
}


//=================================================================|1
//							APIの初期設定
//========================================================|2022.01.04
// v1: 関数の作成
function showApiConf() {
  const html = HtmlService.createHtmlOutputFromFile('form')
                .setSandboxMode(HtmlService.SandboxMode.IFRAME)
                .setWidth(800)
                .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'API設定');
    }



  SpreadsheetApp.getActiveSpreadsheet().addMenu("サスケWorks",menu); 
}


//=================================================================|1
//							一括登録API
//========================================================|2021.12.31
// v1: 関数の作成
function apiImportPost() {
  //----------------------------------------
  // 実行確認
  let bool = confirmExecFunction();
  if (!bool) { return }

  //----------------------------------------
  // プロパティの取得
  let propKeys = ['xApiKey', 'xToken', 'appKey', 'saveCode'];
  const properties = getProperties(propKeys);
  if(!properties) { return }

  //----------------------------------------
  // csvデータの取得
  let sheet = setActiveSheet();
  let blob = convertBlob(sheet);

  //----------------------------------------
  // API通信
  const endpoint = 'https://api.work-s.app/v1/' + properties.appKey + '/import';
  let formData = {
    'file': blob,
    'save_code': properties.saveCode
  };
  let options = {
    'method': 'post',
    'headers': {
      'x-api-key': properties.xApiKey,
      'x-token': properties.xToken
    },
    'contentType': 'multipart/form-data',
    'payload': formData
  };
  let response = UrlFetchApp.fetch(endpoint, options);
  let responseText = response.getContentText();

  //----------------------------------------
  // 結果の表示
  Browser.msgBox(responseText);

  //----------------------------------------
  // ログの出力
  const funcName = arguments.callee.name;
  logOutput(funcName);
}

//=================================================================|1
//							DriveフォルダにCSV出力
//========================================================|2021.12.31
// v1: 関数の作成
function csvSaveToDrive() {
  //----------------------------------------
  // 実行確認
  let bool = confirmExecFunction();
  if (!bool) { return }

  //----------------------------------------
  // プロパティの取得
  let propKeys = ['folderId'];
  const properties = getProperties(propKeys);
  if(!properties) { return }

  //----------------------------------------
  // csvデータの取得
  let sheet = setActiveSheet();
  let blob = convertBlob(sheet);

  //----------------------------------------
  // csvの作成
  let folder = DriveApp.getFolderById(properties.folderId);
  let csvFile = folder.createFile(blob);

  if(csvFile) {
    Browser.msgBox('下記フォルダにCSVファイルを作成しました。\n\nhttps://drive.google.com/drive/u/2/folders/' + properties.folderId);
  }

  //----------------------------------------
  // ログの出力
  const funcName = arguments.callee.name;
  logOutput(funcName);
}


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


//=================================================================|2
//							スクリプトプロパティの取得
//========================================================|2021.12.31
// v1: 関数の作成
// v2: keysを配列で受け取り、それに対応するプロパティを取得する仕様に変更
function getProperties(keys) {
  const env = PropertiesService.getScriptProperties();
  const properties = {};
  if (keys.includes('xApiKey')) {
    properties.xApiKey = env.getProperty('API_KEY');
    if(!properties.xApiKey) {
      Browser.msgBox('APIキーが見つかりませんでした。プロパティストアにAPI_KEYを登録しているか確認して下さい。');
      return
    }
  }
  if (keys.includes('xToken')) {
    properties.xToken = env.getProperty('API_TOKEN');
    if(!properties.xToken) {
      Browser.msgBox('APIトークンが見つかりませんでした。プロパティストアにAPI_TOKENを登録しているか確認して下さい。');
      return
    }
  }
  if (keys.includes('appKey')) {
    properties.appKey = env.getProperty('APP_KEY');
    if(!properties.appKey) {
      Browser.msgBox('アプリキーが見つかりませんでした。プロパティストアにAPP_KEYを登録しているか確認して下さい。');
      return
    }
  }
  if (keys.includes('saveCode')) {
    properties.saveCode = env.getProperty('SAVE_CODE');
    if(!properties.saveCode) {
      Browser.msgBox('識別キーが見つかりませんでした。プロパティストアにSAVE_CODEを登録しているか確認して下さい。');
      return
    }
  }
  if (keys.includes('folderId')) {
    properties.folderId = env.getProperty('FOLDER_ID');
    if(!properties.folderId) {
      Browser.msgBox('フォルダIDの取得に失敗しました。プロパティストアにFOLDER_IDを登録しているか確認して下さい。');
      return
    }
  }
  return properties
}


//=================================================================|1
//							作業中スプレッドシートの最初のシートを取得
//========================================================|2021.12.31
// v1: 関数の作成
function setActiveSheet() {
  let activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  return activeSheet;
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



