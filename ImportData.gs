/**
 * 外部の「動物形態機能学」スプレッドシートからデータを読み込み、
 * Google Learn Portal用の「Master」シート形式に整形して追記するスクリプト
 */
function importQuizData() {
  // ==========================================
  // 設定エリア：ここを書き換えてください
  // ==========================================
  
  // 1. 読み込み元（ソース）のスプレッドシートURL
  const SOURCE_SS_URL = 'https://docs.google.com/spreadsheets/d/xxxxxxxxxxxxxxxxxxxxxxxxxxxxx/edit';
  
  // 2. 読み込み元のシート名（CSVをインポートしたシートの名前）
  const SOURCE_SHEET_NAME = '問題集の名前'; 

  // 3. 書き込み先（このGASがあるスプシ）のシート名
  // ※Google Learn Portalのデフォルトは 'Master' です
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
      // シート名が見つからない場合、1番目のシートを使用する（CSVインポート直後などを想定）
      sourceSheet = sourceSS.getSheets()[0];
    }
  } catch (e) {
    Browser.msgBox("エラー: 読み込み元スプレッドシートを開けませんでした。URLと権限を確認してください。\\n" + e.message);
    return;
  }

  // データの読み込み（ヘッダー含む全データ）
  const values = sourceSheet.getDataRange().getValues();
  if (values.length < 2) {
    Browser.msgBox("データがありません。");
    return;
  }

  // ヘッダー行から列インデックスを特定（列の並び順が変わっても対応できるようにする）
  const headers = values[0];
  const colMap = {
    question: headers.indexOf('question'),
    image: headers.indexOf('questionImageURL'),
    explanation: headers.indexOf('explanation'),
    link: headers.indexOf('Link'),
    trueOpt: headers.indexOf('trueOption'),
    falseOpt: headers.indexOf('falseOption')
  };

  // 必須カラムチェック
  if (colMap.question === -1 || colMap.trueOpt === -1 || colMap.falseOpt === -1) {
    Browser.msgBox("エラー: 必須列（question, trueOption, falseOption）が見つかりません。");
    return;
  }

  // 転記用データの作成
  const newRows = [];

  // 2行目からデータ処理開始
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // データを取得
    const question = row[colMap.question];
    const image = colMap.image > -1 ? row[colMap.image] : '';
    const explanation = colMap.explanation > -1 ? row[colMap.explanation] : '';
    const link = colMap.link > -1 ? row[colMap.link] : '';
    const trueOption = String(row[colMap.trueOpt]).trim();
    const falseOptionRaw = String(row[colMap.falseOpt]);

    // 不正解選択肢を分割（カンマ区切りに対応）
    // ※ CSVの状況に合わせて区切り文字を調整しています（カンマまたは読点）
    let falseOptions = falseOptionRaw.split(/,\s*|、\s*/).map(s => s.trim()).filter(s => s !== "");

    // 選択肢をマージしてシャッフル
    let allOptions = [trueOption, ...falseOptions];
    allOptions = shuffleArray_(allOptions);

    // Google Learn Portal (Master) の形式に合わせて配列を作成
    // 想定カラム順: [問題文, 画像URL, 選択肢1, 選択肢2, 選択肢3, 選択肢4, 正解, 解説, 参考URL]
    // ※プロジェクトによって「選択肢の数」のカラム数が違う場合があるため、最大5つまで対応するようにしています
    
    const formattedRow = [
      question,           // A列: 問題文
      image,              // B列: 画像URL
      allOptions[0] || '', // C列: 選択肢1
      allOptions[1] || '', // D列: 選択肢2
      allOptions[2] || '', // E列: 選択肢3
      allOptions[3] || '', // F列: 選択肢4
      allOptions[4] || '', // G列: 選択肢5（予備）
      trueOption,         // H列: 正解（テキスト一致判定用）
      explanation,        // I列: 解説
      link                // J列: 参考URL
    ];

    newRows.push(formattedRow);
  }

  // Masterシートの最終行に追加
  if (newRows.length > 0) {
    // Masterシートのカラム構造に合わせて書き込み
    // A列(1)からJ列(10)までと仮定
    const lastRow = destSheet.getLastRow();
    destSheet.getRange(lastRow + 1, 1, newRows.length, 10).setValues(newRows);
    
    Browser.msgBox("完了: " + newRows.length + " 件のデータをMasterシートに追加しました。");
  } else {
    Browser.msgBox("追加対象のデータがありませんでした。");
  }
}

/**
 * 配列をランダムにシャッフルするヘルパー関数 (Fisher-Yates)
 */
function shuffleArray_(array) {
  const newArray = array.slice(); 
  for (let i = newArray.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [newArray[i], newArray[j]] = [newArray[j], newArray[i]];
  }
  return newArray;
}
