//=================================================================|2
//							スプレッドシートを開いた際に実行
//========================================================|2022.01.03
// v1: 関数の作成
// v2: 初期設定登録フォームの実装
/**
 * スプレッドシートを開いた際に、メニューバーにカスタムメニューを追加する。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu(saaskeWorks);
  menu.addSubMenu(ui.createMenu(functionJp)
    .addItem(convertToCsv, 'csvSaveToDrive')
    .addItem(bulkRegistration, 'apiImportPost')
  );
  menu.addSeparator();
  menu.addItem(apiSetting, 'showApiConf');
  menu.addToUi();
}


//=================================================================|1
//							APIの初期設定
//========================================================|2022.01.04
// v1: 関数の作成
/**
 * 「API設定」ボタンが押された際、ダイアログにform.htmlを表示する。
 */
function showApiConf() {
  const html = HtmlService.createHtmlOutputFromFile('form')
                .setSandboxMode(HtmlService.SandboxMode.IFRAME)
                .setWidth(800)
                .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, apiSetting);
}


//=================================================================|1
//							一括登録API
//========================================================|2021.12.31
// v1: 関数の作成
/**
 * 先頭のシートからデータを抽出し、Worksアプリへ一括登録を行う。
 * その後、logシートへ操作ログを出力する。
 */
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
  let sheet = setFirstSheet();
  let blob = convertBlob(sheet);

  //----------------------------------------
  // API通信
  const endpoint = apiVersionOneUrl + properties.appKey + apiResourceUrl;
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
/**
 * 指定したDriveフォルダにWorksアプリデータをCSV出力する。
 */
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
  let sheet = setFirstSheet();
  let blob = convertBlob(sheet);

  //----------------------------------------
  // csvの作成
  let folder = DriveApp.getFolderById(properties.folderId);
  let csvFile = folder.createFile(blob);

  if(csvFile) {
    Browser.msgBox(csvSaveToDriveMsg + properties.folderId);
  }

  //----------------------------------------
  // ログの出力
  const funcName = arguments.callee.name;
  logOutput(funcName);
}

