// スクリプトプロパティから取得
const ROOT_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID');
const SHEET_ID = PropertiesService.getScriptProperties().getProperty('SHEET_ID');

// スプレッドシート情報
const SHEET_NAME = '都道府県マスタ';
const LOG_SHEET_NAME = 'アップロードログ';
const COL_PREF = 1; // A列
const COL_CITY = 2; // B列
const COL_URL  = 6; // F列

function doGet() {
  return HtmlService.createHtmlOutputFromFile('form')
    .setTitle('ポスター掲示板写真アップロード')
    .addMetaTag('viewport', 'initial-scale=0.4, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getLocationOptions() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < values.length; i++) {
    const pref = values[i][0];
    const city = values[i][1];
    if (!pref || !city) continue;
    if (!map[pref]) {
      map[pref] = [];
    }
    map[pref].push(city);
  }
  // 重複削除
  for (const k in map) {
    map[k] = Array.from(new Set(map[k]));
  }
  return map;
}

function getUploadFolderId(pref, city) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === pref && values[i][1] === city) {
      const url = values[i][5]; // F列
      if (!url) return null;
      const m = String(url).match(/[-\w]{25,}/);
      return m ? m[0] : null;
    }
  }
  return null;
}

function ensureLogSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET_NAME);
    sh.appendRow([
      'Timestamp','Prefecture','City',
      'Number','Comment',                // ← 追加
      'Filename','File URL','File ID',
      'Latitude','Longitude','TakenAt'
    ]);
    return sh;
  }
  // 既存シートにも不足カラムがあれば末尾に追加
  const headers = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0];
  const need = ['Number','Comment'].filter(h => headers.indexOf(h) === -1);
  if (need.length) {
    sh.getRange(1, sh.getLastColumn()+1, 1, need.length).setValues([need]);
  }
  return sh;
}

function logUpload(pref, city, number, comment, filename, url, id, exif) {
  const sh = ensureLogSheet_();
  const headers = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0];
  const idx = {};
  headers.forEach((h,i)=> idx[h]=i); // 0-based

  const row = new Array(headers.length).fill('');
  row[idx['Timestamp']] = new Date();
  if ('Prefecture' in idx) row[idx['Prefecture']] = pref;
  if ('City' in idx)       row[idx['City']]       = city;
  if ('Number' in idx)     row[idx['Number']]     = number || '';
  if ('Comment' in idx)    row[idx['Comment']]    = comment || '';
  if ('Filename' in idx)   row[idx['Filename']]   = filename;
  if ('File URL' in idx)   row[idx['File URL']]   = url;
  if ('File ID' in idx)    row[idx['File ID']]    = id;
  if ('Latitude' in idx)   row[idx['Latitude']]   = exif && exif.lat != null ? exif.lat : '';
  if ('Longitude' in idx)  row[idx['Longitude']]  = exif && exif.lng != null ? exif.lng : '';
  if ('TakenAt' in idx)    row[idx['TakenAt']]    = exif && exif.takenAt ? exif.takenAt : '';

  const next = sh.getLastRow()+1;
  sh.getRange(next, 1, 1, headers.length).setValues([row]);
}

/**
 * ブラウザから base64 データを受け取って Drive に保存
 * @param {Object} exif { lat?:number, lng?:number, takenAt?:string }
 * @param {string} number 掲示板番号（全ファイル共通）
 * @param {string} comment コメント（全ファイル共通）
 */
function uploadSingleFileBase64(pref, city, fileObj, folderId, exif, number, comment) {
  if (!folderId) throw new Error('Folder ID not found.');
  const bytes = Utilities.base64Decode(fileObj.data);
  const blob = Utilities.newBlob(bytes, fileObj.mimeType || MimeType.BINARY, fileObj.name);
  const folder = DriveApp.getFolderById(folderId);
  const file = folder.createFile(blob);
  const res = { name: file.getName(), id: file.getId(), url: file.getUrl() };
  // ← ここで番号・コメントも一緒に記録
  logUpload(pref, city, number, comment, res.name, res.url, res.id, exif);
  return res;
}

/***** エントリーポイント：F列が空の行のみ処理 *****/
function createFoldersForEmptyF() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`シートが見つかりません: ${SHEET_NAME}`);

  const range = sh.getDataRange();
  const values = range.getValues(); // 2次元配列

  const rootFolder = DriveApp.getFolderById(extractFolderId(ROOT_FOLDER_ID));

  // パフォーマンス用キャッシュ（同名フォルダの重複作成を防止）
  const prefCache = new Map(); // key: prefecture name -> Folder
  const cityCache = new Map(); // key: prefecture name + '>' + city -> Folder

  // 見出し行をスキップ
  for (let r = 1; r < values.length; r++) {
    const pref = String(values[r][COL_PREF - 1] || '').trim();
    const city = String(values[r][COL_CITY - 1] || '').trim();
    const url  = String(values[r][COL_URL  - 1] || '').trim();

    // 条件：F列が空の行のみ／A,Bが入っていること
    if (!pref || !city || url) continue;

    try {
      // 都道府県フォルダを取得/作成（キャッシュあり）
      let prefFolder = prefCache.get(pref);
      if (!prefFolder) {
        prefFolder = getOrCreateSubFolder(rootFolder, pref);
        prefCache.set(pref, prefFolder);
      }

      // 市区町村フォルダを取得/作成（キャッシュあり）
      const cityKey = `${pref}>${city}`;
      let cityFolder = cityCache.get(cityKey);
      if (!cityFolder) {
        cityFolder = getOrCreateSubFolder(prefFolder, city);
        cityCache.set(cityKey, cityFolder);
      }

      // F列にURLを書き込み
      const folderUrl = cityFolder.getUrl();
      sh.getRange(r + 1, COL_URL).setValue(folderUrl); // rは0始まり、行は1始まり
    } catch (e) {
      // 失敗しても他行は続行。必要ならE列（ステータス）にエラーを書いてもOK
      // sh.getRange(r + 1, 5).setValue(`エラー: ${e.message || e}`);
      console.error(`Row ${r + 1}: ${e && e.message ? e.message : e}`);
    }
  }
}

/***** 補助：URLやIDからフォルダIDを抽出 *****/
function extractFolderId(idOrUrl) {
  if (!idOrUrl) throw new Error('ROOT_FOLDER_ID が未設定です。');
  const m = String(idOrUrl).match(/[-\w]{25,}/);
  return m ? m[0] : idOrUrl;
}

/***** 補助：親フォルダ直下に同名があればそれを返し、無ければ作成 *****/
function getOrCreateSubFolder(parentFolder, name) {
  // まず同名のサブフォルダがあるか確認
  const it = parentFolder.getFoldersByName(name);
  if (it.hasNext()) {
    return it.next();
  }
  // 無ければ作成
  return parentFolder.createFolder(name);
}
