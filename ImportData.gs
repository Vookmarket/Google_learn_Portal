/**
 * 外部の「動物形態機能学」スプレッドシートからデータを読み込み、
 * Google Learn Portal用の「Master」シート形式（全17列・正答番号対応）に整形して追記するスクリプト
 */
function importQuizData() {
  // ==========================================
  // 設定エリア
  // ==========================================
  
  // 1. 読み込み元（ソース）のスプレッドシートURL
  const SOURCE_SS_URL = 'https://docs.google.com/spreadsheets/d/xxxxxxxxxxxxxxxxxxxxxxxxxxxxx/edit';
  
  // 2. 読み込み元のシート名
  const SOURCE_SHEET_NAME = '動物形態機能学・国家試験対策 - 問題集'; 

  // 3. 書き込み先のシート名
  const DEST_SHEET_NAME = 'Master'; 

  // ==========================================
  // 処理ロジック
  // ==========================================

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = ss.getSheetByName(DEST_SHEET_NAME);
  
  if (!destSheet) {
    Browser.msgBox("エラー: 書き込み先のシート '" + DEST_SHEET_NAME + "' が見つかりません。");
    return;
  }

  // ソースデータの取得
  let sourceSheet;
  try {
    const sourceSS = SpreadsheetApp.openByUrl(SOURCE_SS_URL);
    sourceSheet = sourceSS.getSheetByName(SOURCE_SHEET_NAME);
    if (!sourceSheet) {
      sourceSheet = sourceSS.getSheets()[0]; // 見つからない場合は1番目のシート
    }
  } catch (e) {
    Browser.msgBox("エラー: 読み込み元を開けませんでした。\\n" + e.message);
    return;
  }

  const values = sourceSheet.getDataRange().getValues();
  if (values.length < 2) {
    Browser.msgBox("データがありません。");
    return;
  }

  // ヘッダー行から列インデックスを特定
  const headers = values[0];
  const colMap = {
    question: headers.indexOf('question'),
    qImage: headers.indexOf('questionImageURL'),
    explanation: headers.indexOf('explanation'),
    eImage: headers.indexOf('explanationImageURL'), // 解説画像があれば取得
    link: headers.indexOf('Link'),
    trueOpt: headers.indexOf('trueOption'),
    falseOpt: headers.indexOf('falseOption')
  };

  // 必須カラムチェック
  if (colMap.question === -1 || colMap.trueOpt === -1 || colMap.falseOpt === -1) {
    Browser.msgBox("エラー: 必須列（question, trueOption, falseOption）が見つかりません。");
    return;
  }

  const newRows = [];

  // 2行目からデータ処理開始
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // データ取得
    const question = row[colMap.question];
    const qImage = colMap.qImage > -1 ? row[colMap.qImage] : '';
    const explanation = colMap.explanation > -1 ? row[colMap.explanation] : '';
    const eImage = colMap.eImage > -1 ? row[colMap.eImage] : '';
    const link = colMap.link > -1 ? row[colMap.link] : '';
    const trueOption = String(row[colMap.trueOpt]).trim();
    const falseOptionRaw = String(row[colMap.falseOpt]);

    // 不正解選択肢を分割
    let falseOptions = falseOptionRaw.split(/,\s*|、\s*/).map(s => s.trim()).filter(s => s !== "");

    // 選択肢をマージしてシャッフル
    let allOptions = [trueOption, ...falseOptions];
    allOptions = shuffleArray_(allOptions); // シャッフル実行

    // 正解が何番目（1〜5）に移動したかを探す
    // indexOfは0始まりなので+1する
    const correctIndex = allOptions.indexOf(trueOption);
    const correctNumber = correctIndex + 1;

    // UUID生成 (問題ID用)
    const uuid = Utilities.getUuid();

    // Masterシートの列順に合わせて配列を作成
    // 順序: [ID, テキスト, 画像, 選1, 選2, 選3, 選4, 選5, 正答番号, 解説, 解説画像, 参考URL, 参考タイトル, カテゴリ, サブカテ, FormID, ImageID]
    const formattedRow = [
      uuid,                 // A: 問題ID
      question,             // B: 問題テキスト
      qImage,               // C: 問題画像URl
      allOptions[0] || '',  // D: 選択肢1
      allOptions[1] || '',  // E: 選択肢2
      allOptions[2] || '',  // F: 選択肢3
      allOptions[3] || '',  // G: 選択肢4
      allOptions[4] || '',  // H: 選択肢5
      correctNumber,        // I: 正答番号 (数値)
      explanation,          // J: 解説テキスト
      eImage,               // K: 解説画像URl
      link,                 // L: 参考URL
      '',                   // M: 参考リンクタイトル (空欄)
      '',                   // N: カテゴリ (空欄)
      '',                   // O: サブカテゴリ (空欄)
      '',                   // P: FormItemId (自動処理のため空欄)
      ''                    // Q: ImageItemId (自動処理のため空欄)
    ];

    newRows.push(formattedRow);
  }

  // Masterシートへ書き込み
  if (newRows.length > 0) {
    const lastRow = destSheet.getLastRow();
    // 17列分のデータを書き込む
    destSheet.getRange(lastRow + 1, 1, newRows.length, 17).setValues(newRows);
    
    Browser.msgBox("完了: " + newRows.length + " 件のデータを追加しました。");
  } else {
    Browser.msgBox("データがありませんでした。");
  }
}

/**
 * 配列をランダムにシャッフルするヘルパー関数
 */
function shuffleArray_(array) {
  const newArray = array.slice(); 
  for (let i = newArray.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [newArray[i], newArray[j]] = [newArray[j], newArray[i]];
  }
  return newArray;
}
