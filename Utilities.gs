/**
 * 共通ユーティリティ (Utilities) v2
 * 修正点: getFileIdFromUrl を強化し、スプレッドシートのURLにも対応
 */

// 設定定数
const TIME_LIMIT_SECONDS = 280; 
const TARGET_SHEET_NAME = 'Form Responses 1';

/** Configシートから設定を連想配列で取得 */
function getConfig(sheet) {
  if (!sheet) return {};
  const values = sheet.getRange("A2:B20").getValues();
  const config = {};
  values.forEach(row => { if(row[0]) config[row[0]] = row[1]; });
  return config;
}

/** Configシートへ値を書き込み */
function setConfigValue(sheet, key, value) {
  if (!sheet) return;
  const data = sheet.getRange("A:A").getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

/** 実行時間チェック */
function isTimeUp(startTime) {
  return (new Date().getTime() - startTime) > (TIME_LIMIT_SECONDS * 1000);
}

/** 配列シャッフル */
function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}

/** * URLからファイルIDを抽出 (強化版)
 * 対応: /file/d/, /spreadsheets/d/, /forms/d/, id=xxx
 */
function getFileIdFromUrl(url) {
  if (!url || url === "") return null;
  let id = "";
  try {
    // パターン1: /d/xxxxx/ の形式
    const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (match) {
      id = match[1];
    } 
    // パターン2: id=xxxxx の形式
    else if (url.includes('id=')) {
      const matchId = url.match(/id=([a-zA-Z0-9_-]+)/);
      if (matchId) id = matchId[1];
    }
  } catch (e) {
    console.warn("ID抽出失敗: " + url);
  }
  return id;
}

/** URLから画像Blobを取得 */
function getBlobFromUrl(url) {
  const id = getFileIdFromUrl(url);
  if (!id) return null;
  try {
    return DriveApp.getFileById(id).getBlob();
  } catch (e) {
    console.warn(`画像取得エラー: ${url}`);
    return null;
  }
}

/** URLからBase64文字列を取得 */
function getBase64FromUrl(url) {
  const id = getFileIdFromUrl(url);
  if (!id) return null;
  try {
    const blob = DriveApp.getFileById(id).getBlob();
    return "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    return null;
  }
}