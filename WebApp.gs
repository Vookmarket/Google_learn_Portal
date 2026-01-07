/**
 * ダッシュボード WebApp (Utilities利用版 + リンク対応)
 */

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pageTitle = ss.getName() + " 結果分析";
  const template = HtmlService.createTemplateFromFile('index');
  template.pageTitle = pageTitle; 
  return template.evaluate()
    .setTitle(pageTitle)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName('Analysis_Users');
    const questionSheet = ss.getSheetByName('Analysis_Questions');
    const masterSheet = ss.getSheetByName('Master');

    if (!userSheet || !questionSheet || !masterSheet) throw new Error("必要なシートが見つかりません。");
    
    // マスタデータ
    const masterValues = masterSheet.getDataRange().getValues();
    const masterMap = new Map();
    // A:ID(0), B:Text(1), C:Img(2), D-H:Choices(3-7), I:Correct(8), ..., L:RefURL(11), M:RefTitle(12)
    for(let i=1; i<masterValues.length; i++) {
      const row = masterValues[i];
      const choices = row.slice(3, 8).filter(c => String(c).trim() !== "");
      masterMap.set(row[0], { 
        qImgUrl: row[2], 
        choices: choices, 
        correctVal: row[8],
        refUrl: row[11],    // ★追加: 参考URL
        refTitle: row[12]   // ★追加: 参考タイトル
      }); 
    }

    // 問題分析データ
    const qLastRow = questionSheet.getLastRow();
    const questionRows = qLastRow > 1 ? questionSheet.getRange(2, 1, qLastRow - 1, 10).getValues() : [];
    
    const enrichedQuestions = questionRows.map(q => {
      const masterInfo = masterMap.get(q[0]) || { choices: [], correctVal: 0, qImgUrl: "", refUrl: "", refTitle: "" };
      return {
        id: q[0], text: q[1], category: q[2], rate: Number(q[6]),
        di: (q[7] === "" || q[7] === null || isNaN(Number(q[7]))) ? -999 : Number(q[7]),
        diffLabel: q[8], choices: masterInfo.choices, correctVal: masterInfo.correctVal,
        qImgUrl: masterInfo.qImgUrl, base64Img: null,
        refUrl: masterInfo.refUrl,      // ★データに追加
        refTitle: masterInfo.refTitle   // ★データに追加
      };
    });

    // 難問・良問選定
    const hardTop3 = [...enrichedQuestions].sort((a, b) => (a.rate !== b.rate) ? a.rate - b.rate : 0.5 - Math.random()).slice(0, 3);
    const goodTop3 = [...enrichedQuestions].filter(q => q.di !== -999).sort((a, b) => (a.di !== b.di) ? b.di - a.di : 0.5 - Math.random()).slice(0, 3);

    // 画像変換 (Utilities使用)
    [...hardTop3, ...goodTop3].forEach(q => {
      if (q.qImgUrl) q.base64Img = getBase64FromUrl(q.qImgUrl);
    });

    // ユーザーデータ
    const uLastRow = userSheet.getLastRow();
    const userRows = uLastRow > 1 ? userSheet.getRange(2, 1, uLastRow - 1, 8).getValues() : [];
    const safeUserRows = userRows.map(r => {
      if (r[1] instanceof Date) r[1] = Utilities.formatDate(r[1], Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
      return r;
    });

    const scores = safeUserRows.map(r => Number(r[3]) || 0);
    const totalScore = scores.reduce((a,b)=>a+b,0);
    const avgScore = scores.length ? (totalScore / scores.length).toFixed(1) : "0.0";
    const passCount = safeUserRows.filter(r => String(r[6]) === '合格').length;
    const passRate = safeUserRows.length ? ((passCount / safeUserRows.length) * 100).toFixed(1) : "0.0";

    return {
      status: "success",
      summary: { totalUsers: safeUserRows.length, avgScore: avgScore, passRate: passRate },
      users: safeUserRows,
      hardQuestions: hardTop3,
      goodQuestions: goodTop3
    };

  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}