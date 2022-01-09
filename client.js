//=================================================================|1
//							HTML側でプロパティを取得する
//========================================================|2022.01.05
// v1: 関数の作成
/**
 * 登録されているスクリプトプロパティを返す。
 * @returns 
 */
function getPropsOnClient() {
  const properties = PropertiesService.getScriptProperties().getProperties();

  const propData = [];
  for (let key in properties) {
    let temProp = {};
    temProp.key = key;
    temProp.value = properties[key];
    propData.push(temProp);
  }
  propData.sort(function(top, bottom) {
    if (top.key > bottom.key) { return 1; }
    else { return -1; }
  });

  return JSON.stringify(propData);
}


//=================================================================|1
//							HTML側でプロパティを登録・修正する
//========================================================|2022.01.08
// v1: 関数の作成
/**
 * クライアント側から指定したスクリプトプロパティを登録・修正する。
 * @param {string} key 
 * @param {string} value 
 * @returns 
 */
function setPropAndRemount(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
  return getPropsOnClient();
}


//=================================================================|1
//							HTML側でプロパティを削除する
//========================================================|2022.01.08
// v1: 関数の作成
/**
 * クライアント側から指定したスクリプトプロパティを削除する。
 * @param {object} prop 
 * @returns 
 */
function deletePropAndRemount(prop) {
  PropertiesService.getScriptProperties().deleteProperty(prop.key);
  return getPropsOnClient();
}

