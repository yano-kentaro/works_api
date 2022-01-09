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
  const properties = getProps(propKeys, true);
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
  const properties = getProps(propKeys, true);
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

