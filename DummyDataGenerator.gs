/**
 * クイズシステム: ダミーデータ生成ツール
 * 機能: Masterシートを参照し、指定された件数のランダムな回答データを回答シートに生成する
 */

// ■ 設定エリア
const DUMMY_CONFIG = {
  NUM_RECORDS: 3,       // 生成するデータ数
  TARGET_ACCURACY: 0.6, // 目標正答率 (0.0〜1.0) ※nullにすると完全ランダム
  
  // ランダムに使用する名前リスト
  NAMES: [
    "ごんた", "スフレ", "つくね", "こむぎ", "こうめ", 
    "ミッキー", "ミーシャ", "こたろう", "おかか", "とろろ"
  ]
};

function generateDummyData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('Master');
  const responseSheet = ss.getSheetByName('Form Responses 1');

  if (!masterSheet || !responseSheet) {
    Browser.msgBox("エラー: 必須シートが見つかりません。");
    return;
  }

  // 1. Masterデータの読み込み
  // A:ID, B:Text, C:Img, D-H:Choices, I:Correct ...
  const masterValues = masterSheet.getDataRange().getValues();
  const questions = [];
  
  for (let i = 1; i < masterValues.length; i++) {
    const row = masterValues[i];
    if(!row[0]) continue; // IDなしはスキップ

    // 有効な選択肢だけ抽出 (D列〜H列)
    const validChoices = row.slice(3, 8).filter(c => String(c).trim() !== "");
    const correctIndex = Number(row[8]) - 1; // 0始まりのインデックスに変換
    
    questions.push({
      id: row[0],
      choices: validChoices,
      correctText: validChoices[correctIndex], // 正解のテキスト
      correctChoiceIndex: correctIndex
    });
  }

  // 2. 回答シートの列特定
  const headers = responseSheet.getDataRange().getValues()[0];
  const colMap = {
    timestamp: headers.indexOf('タイムスタンプ'),
    score: headers.indexOf('スコア'),
    name: headers.indexOf('ニックネーム (回答者名)'),
    rankAllow: headers.indexOf('ランキングへの掲載')
  };

  // 質問列の開始位置 (属性列の次からと仮定)
  const qStartIndex = Math.max(colMap.timestamp, colMap.score, colMap.name, colMap.rankAllow) + 1;

  // 3. ダミーデータ生成ループ
  const newRows = [];
  
  for (let i = 0; i < DUMMY_CONFIG.NUM_RECORDS; i++) {
    const rowData = new Array(headers.length).fill(""); // 空の行を作成

    // 基本情報
    const now = new Date();
    now.setMinutes(now.getMinutes() - Math.floor(Math.random() * 1000)); // 少し過去の時間にする
    rowData[colMap.timestamp] = now;
    
    // 名前 (ランダム + ID)
    const baseName = DUMMY_CONFIG.NAMES[Math.floor(Math.random() * DUMMY_CONFIG.NAMES.length)];
    rowData[colMap.name] = baseName /*+ "_dummy" + (i+1)*/;
    
    // ランキング掲載 (ランダム)
    rowData[colMap.rankAllow] = Math.random() > 0.5 ? "はい、掲載して構いません" : "いいえ、掲載しないでください";

    // 回答生成
    let correctCount = 0;
    
    questions.forEach((q, idx) => {
      const colIndex = qStartIndex + idx;
      if (colIndex >= headers.length) return; // 列が足りない場合は無視

      let selectedText = "";
      
      // 正答率に基づく選択ロジック
      const isCorrectTarget = (DUMMY_CONFIG.TARGET_ACCURACY !== null) 
        ? (Math.random() < DUMMY_CONFIG.TARGET_ACCURACY) 
        : (Math.random() < (1 / q.choices.length)); // 完全ランダム

      if (isCorrectTarget && q.correctText) {
        // 正解を選ぶ
        selectedText = q.correctText;
        correctCount++;
      } else {
        // 不正解を選ぶ (正解以外の選択肢からランダム)
        const wrongChoices = q.choices.filter(c => c !== q.correctText);
        if (wrongChoices.length > 0) {
          selectedText = wrongChoices[Math.floor(Math.random() * wrongChoices.length)];
        } else {
          // 選択肢が1つしかない等の異常時はそのまま正解を入れるか、適当に処理
          selectedText = q.choices[0]; 
        }
      }
      
      rowData[colIndex] = selectedText;
    });

    // スコア入力 ( "3 / 5" の形式 )
    rowData[colMap.score] = `${correctCount} / ${questions.length}`;

    newRows.push(rowData);
  }

  // 4. シートへ書き込み
  if (newRows.length > 0) {
    const lastRow = responseSheet.getLastRow();
    responseSheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    
    // 集計スクリプトを呼び出して反映させる
    // (DataAggregator.gs がある前提)
    try {
      runAggregation();
      Browser.msgBox(`ダミーデータ ${newRows.length} 件を作成し、集計を更新しました。`);
    } catch (e) {
      Browser.msgBox(`ダミーデータ作成完了。\n(集計更新エラー: ${e.message})`);
    }
  }
}