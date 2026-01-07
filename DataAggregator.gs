/**
 * 集計・分析エンジン (Utilities利用版 + 重複ネーム自動リネーム)
 * 修正点: 同じニックネームが2回目以降登場した場合、末尾に日時を付与して別名として登録する
 */

function runAggregation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // シート取得
  const configSheet = ss.getSheetByName('Config');
  const masterSheet = ss.getSheetByName('Master');
  const responseSheet = ss.getSheetByName('Form Responses 1');
  const userSheet = ss.getSheetByName('Analysis_Users');
  const questionSheet = ss.getSheetByName('Analysis_Questions');

  if (!responseSheet || responseSheet.getLastRow() < 2) {
    console.log("まだ回答データがありません。");
    return;
  }
  
  const config = getConfig(configSheet); // Utilities
  const passThreshold = Number(config['Threshold_Pass']) || 70;
  const hardThreshold = Number(config['Threshold_Diff_Hard']) || 30;
  const goodIndexThreshold = Number(config['Threshold_Index_Good']) || 0.2;

  // --- Masterデータ読み込み ---
  const masterValues = masterSheet.getDataRange().getValues();
  const questions = [];
  
  for (let i = 1; i < masterValues.length; i++) {
    const row = masterValues[i];
    if(!row[0]) continue;
    questions.push({
      id: row[0],
      text: row[1],
      correctVal: Number(row[8]), 
      correctText: row[3 + (Number(row[8]) - 1)], 
      category: row[13], 
      subCategory: row[14],
      index: i
    });
  }

  // --- 回答データ解析 ---
  const responseValues = responseSheet.getDataRange().getValues();
  const headers = responseValues[0];
  // ヘッダー検出を柔軟に変更（部分一致対応）
  const colMap = {
    timestamp: -1,
    score: -1, 
    name: -1,
    rankAllow: -1
  };

  headers.forEach((h, i) => {
    const str = String(h).trim();
    if (str === 'タイムスタンプ') colMap.timestamp = i;
    else if (str === 'スコア') colMap.score = i;
    else if (str.includes('ニックネーム') || str.includes('回答者名')) colMap.name = i;
    else if (str.includes('ランキング') && str.includes('掲載')) colMap.rankAllow = i;
  });

  // 必須列が見つからない場合のログ出力
  if (colMap.name === -1) console.warn("警告: 'ニックネーム'列が見つかりません。");
  if (colMap.rankAllow === -1) console.warn("警告: 'ランキングへの掲載'列が見つかりません。掲載フラグはfalseになります。");

  let questionStartIndex = Math.max(
    colMap.timestamp, colMap.score, colMap.name, colMap.rankAllow
  ) + 1;

  // --- ユーザーごとの採点・集計 ---
  const userRows = [];
  const responses = [];
  
  // ★重複チェック用のセット (登場した名前を記憶する)
  const seenNames = new Set();

  for (let i = 1; i < responseValues.length; i++) {
    const row = responseValues[i];
    const timestamp = row[colMap.timestamp];
    let displayName = row[colMap.name]; // 表示名(ニックネーム)
    const rankAllowVal = row[colMap.rankAllow];
    
    // ★重複ネームの処理ロジック
    // すでに同じ名前が存在していたら、日時を付与してリネームする
    if (seenNames.has(displayName)) {
      const dateSuffix = Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), "_yyyyMMddHHmm");
      displayName = displayName + dateSuffix;
    }
    // 名前を記憶セットに追加
    seenNames.add(displayName);

    // 採点処理
    let correctCount = 0;
    const answerResults = [];

    questions.forEach((q, idx) => {
      const colIndex = questionStartIndex + idx;
      const userAnsStr = row[colIndex];
      let isCorrect = false;
      if (userAnsStr && q.correctText) {
        if (String(userAnsStr).trim() === String(q.correctText).trim()) {
          isCorrect = true;
        }
      }
      answerResults.push(isCorrect);
      if (isCorrect) correctCount++;
    });

    const percentage = (correctCount / questions.length) * 100;
    const isPass = percentage >= passThreshold;
    const isRank = (String(rankAllowVal).includes('はい'));

    // ID生成 (ここもdisplayNameベースにする)
    const uniqueId = displayName; 

    userRows.push([
      uniqueId, 
      timestamp, 
      displayName, // ★加工後の名前を保存
      correctCount, 
      questions.length, 
      percentage, 
      isPass ? "合格" : "不合格", 
      isRank ? "掲載" : "非掲載"
    ]);
    
    responses.push({ isPass: isPass, answers: answerResults });
  }

  // --- シート書き込み (Analysis_Users) ---
  if (userRows.length > 0) {
    if(userSheet.getLastRow() > 1) {
      userSheet.getRange(2, 1, userSheet.getLastRow()-1, userSheet.getLastColumn()).clearContent();
    }
    userSheet.getRange(2, 1, userRows.length, userRows[0].length).setValues(userRows);
  }

  // --- 問題ごとの統計計算 ---
  const questionRows = [];
  const passGroup = responses.filter(r => r.isPass);
  const failGroup = responses.filter(r => !r.isPass);
  
  questions.forEach((q, idx) => {
    const totalCorrect = responses.filter(r => r.answers[idx]).length;
    const totalRate = responses.length > 0 ? (totalCorrect / responses.length) : 0;
    
    let discriminationIndex = "";
    if (passGroup.length > 0 && failGroup.length > 0) {
      const passRate = passGroup.filter(r => r.answers[idx]).length / passGroup.length;
      const failRate = failGroup.filter(r => r.answers[idx]).length / failGroup.length;
      discriminationIndex = passRate - failRate;
    }

    const ratePercent = totalRate * 100;
    const diffLabel = ratePercent <= hardThreshold ? "難" : (ratePercent >= 80 ? "易" : "普");
    const goodLabel = (typeof discriminationIndex === 'number' && discriminationIndex >= goodIndexThreshold) ? "★良問" : "-";

    questionRows.push([q.id, q.text, q.category, q.subCategory, totalCorrect, responses.length, totalRate, discriminationIndex, diffLabel, goodLabel]);
  });

  if (questionRows.length > 0) {
    if(questionSheet.getLastRow() > 1) questionSheet.getRange(2, 1, questionSheet.getLastRow()-1, questionSheet.getLastColumn()).clearContent();
    questionSheet.getRange(2, 1, questionRows.length, questionRows[0].length).setValues(questionRows);
  }
  
  console.log(`集計完了: ${userRows.length}件処理しました。`);
}
