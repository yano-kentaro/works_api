//=================================================================|1
//							HTML側でプロパティを取得する
//========================================================|2022.01.05
// v1: 関数の作成
function getPropsOnClient() {
  const properties = PropertiesService.getScriptProperties().getProperties();

  const propData = [];
  for (let key in properties) {
    let temProp = {};
    temProp.key = key;
    temProp.value = properties[key];
    propData.push(temProp);
  }

  return JSON.stringify(propData);
}


//=================================================================|1
//							HTML側でプロパティを登録・修正する
//========================================================|2022.01.08
// v1: 関数の作成
function setPropAndRemount(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
  return getPropsOnClient();
}


//=================================================================|1
//							HTML側でプロパティを削除する
//========================================================|2022.01.08
// v1: 関数の作成
function deletePropAndRemount(prop) {
  PropertiesService.getScriptProperties().deleteProperty(prop.key);
  return getPropsOnClient();
}


//=================================================================|1
//							作業中スプレッドシートの最初のシートを取得
//========================================================|2021.12.31
// v1: 関数の作成
function setActiveSheet() {
  let activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  return activeSheet;
}
