/**
 * クイズシステム V2.9: 完了画面リンク対応版
 * 機能: ConfigシートのURL(分析/ポータル)を完了メッセージに反映する
 */

// --- onOpen関数を修正 ---
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡️クイズシステムV2')
    .addItem('フォーム生成・更新 (実行)', 'main')
    .addSeparator()
    .addItem('🔄 集計データを更新 (削除の反映)', 'forceAggregation') // ★追加
    .addSeparator()
    .addItem('🤖 ダミーデータ生成', 'generateDummyData')
    .addItem('🌐 ポータルへ登録', 'showRegisterDialog')
    .addSeparator()
    .addItem('⚠️ 設定リセット', 'resetConfig')
    .addToUi();
}

// ...(既存のmain関数などはそのまま)...

// --- ファイルの末尾に以下の関数を追加 ---

/** 手動で集計を実行するラッパー関数 */
function forceAggregation() {
  const ui = SpreadsheetApp.getUi();
  try {
    // DataAggregator.gs の関数を呼び出す
    runAggregation();
    ui.alert("集計完了", "最新の回答データに基づいて分析シートを更新しました。\nGoogleサイトをリロードすると反映されます。", ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("エラー", e.toString(), ui.ButtonSet.OK);
  }
}

function main() {
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const masterSheet = ss.getSheetByName('Master');
  const configSheet = ss.getSheetByName('Config');
  if (!masterSheet || !configSheet) {
    Browser.msgBox("エラー: 必須シート(Master, Config)が見つかりません。");
    return;
  }

  const config = getConfig(configSheet); // Utilities
  let formId = config['Form_ID'];
  let lastRow = parseInt(config['Last_Processed_Row'], 10);
  if (isNaN(lastRow)) lastRow = 1;

  const data = masterSheet.getDataRange().getValues();
  const totalRows = data.length;

  let form;

  try {
    // 1. フォーム準備
    if (!formId) {
      console.log("フェーズ: 新規フォーム作成");

      const oldSheet = ss.getSheetByName(TARGET_SHEET_NAME);
      if (oldSheet) ss.deleteSheet(oldSheet);
      const sheetsBefore = ss.getSheets();
      const idsBefore = sheetsBefore.map(s => s.getSheetId());

      form = FormApp.create(ss.getName());
      formId = form.getId();

      try {
        const ssFile = DriveApp.getFileById(ss.getId());
        const formFile = DriveApp.getFileById(formId);
        const parents = ssFile.getParents();
        if (parents.hasNext()) formFile.moveTo(parents.next());
      } catch (e) { console.warn("フォルダ移動警告: " + e.message); }

      form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

      SpreadsheetApp.flush();
      const sheetsAfter = ss.getSheets();
      let newSheet = null;
      for (const sheet of sheetsAfter) {
        if (!idsBefore.includes(sheet.getSheetId())) {
          newSheet = sheet;
          break;
        }
      }
      if (newSheet) newSheet.setName(TARGET_SHEET_NAME);

      setConfigValue(configSheet, 'Form_ID', formId);
      setConfigValue(configSheet, 'Form_Edit_Url', form.getEditUrl());

      createAttributeSection(form);

    } else {
      console.log(`フェーズ: 既存フォーム(ID: ${formId})を開く`);
      try {
        form = FormApp.openById(formId);
      } catch (e) {
        throw new Error("フォームが見つかりません。設定リセットしてください。");
      }
    }

    // --- ★設定の強制適用 & メッセージ生成 ---
    form.setIsQuiz(false);
    form.setIsQuiz(true);
    form.setCollectEmail(true);
    form.setCollectEmail(false);
    form.setProgressBar(true);
    form.setLimitOneResponsePerUser(false);

    // 完了メッセージの構築
    let confirmMsg = "回答送信が完了しました。\n\n⬇️ 下の【スコアを表示】ボタンを押すと、\n点数・正誤・詳しい解説を確認できます。";

    // ConfigからURLを取得して追記
    const dashboardUrl = config['Dashboard_Url'];
    const portalUrl = config['Portal_Url'];

    if (dashboardUrl) {
      confirmMsg += `\n\n📊 全体の分析結果を見る:\n${dashboardUrl}`;
    }
    if (portalUrl) {
      confirmMsg += `\n\n🏠 ポータルサイトに戻る:\n${portalUrl}`;
    }

    form.setConfirmationMessage(confirmMsg);
    // ----------------------------------------

    // 2. 問題生成・更新
    console.log(`フェーズ: 問題生成開始 (開始行: ${lastRow})`);

    let currentIndex = lastRow;
    let itemIdsToUpdate = [];

    if (totalRows < 2) {
      Browser.msgBox("Masterシートにデータがありません。");
      return;
    }

    for (let i = currentIndex; i < totalRows; i++) {
      if (isTimeUp(startTime)) {
        SpreadsheetApp.getUi().alert(`時間制限のため中断しました。\n現在: ${i}行目完了。\n再実行してください。`);
        break;
      }

      const row = data[i];
      const rowIndex = i + 1;

      const qText = row[1];
      const qImgUrl = row[2];
      const correctNum = Number(row[8]);
      const expText = row[9];
      const expImgUrl = row[10];
      const refUrl = row[11];
      const refTitle = row[12];

      const existingItemId = row[15];
      const existingImgId = row[16];

      if (!qText || isNaN(correctNum) || correctNum < 1 || correctNum > 5) {
        currentIndex = i + 1;
        continue;
      }

      // A. 質問アイテム
      let item;
      if (existingItemId) {
        try { item = form.getItemById(existingItemId); } catch (e) { item = null; }
      }
      let mcItem;
      if (item) {
        mcItem = item.asMultipleChoiceItem();
      } else {
        mcItem = form.addMultipleChoiceItem();
      }

      mcItem.setTitle(`[${row[0]}] ${qText}`).setRequired(true).setPoints(1);

      let rawChoices = [];
      for (let c = 0; c < 5; c++) {
        const choiceText = row[3 + c];
        if (choiceText && String(choiceText).trim() !== "") {
          const isCorrect = ((c + 1) === correctNum);
          rawChoices.push({ text: choiceText, isCorrect: isCorrect });
        }
      }
      rawChoices = shuffleArray(rawChoices);
      const finalChoices = rawChoices.map(c => mcItem.createChoice(c.text, c.isCorrect));
      mcItem.setChoices(finalChoices);

      const feedbackBuilder = FormApp.createFeedback();
      let feedbackMainText = expText || "";
      if (feedbackMainText) feedbackBuilder.setText(feedbackMainText);

      if (refUrl) {
        const title = refTitle || "詳細解説・参考資料はこちら";
        feedbackBuilder.addLink(refUrl, title);
      } else if (expImgUrl) {
        feedbackBuilder.addLink(expImgUrl, "解説図解を開く");
      }

      if (refUrl && expImgUrl) {
        feedbackBuilder.setText(`${feedbackMainText}\n\n▼解説図解:\n${expImgUrl}`);
      } else if (!refUrl && !expImgUrl && feedbackMainText === "") {
        feedbackBuilder.setText(" ");
      }
      const feedback = feedbackBuilder.build();
      mcItem.setFeedbackForCorrect(feedback).setFeedbackForIncorrect(feedback);

      itemIdsToUpdate.push({ row: rowIndex, col: 16, val: mcItem.getId() });

      // B. 画像アイテム
      if (qImgUrl) {
        let imgItem;
        if (existingImgId) {
          try { imgItem = form.getItemById(existingImgId); } catch (e) { imgItem = null; }
        }

        const blob = getBlobFromUrl(qImgUrl); // Utilities
        if (blob) {
          if (!imgItem) {
            imgItem = form.addImageItem();
          } else {
            imgItem = imgItem.asImageItem();
          }
          imgItem.setImage(blob);
          imgItem.setTitle(`[${row[0]}] 参考画像`);
          imgItem.setAlignment(FormApp.Alignment.CENTER);
          itemIdsToUpdate.push({ row: rowIndex, col: 17, val: imgItem.getId() });
        }
      }

      // 改ページ
      if (i % 5 === 0 && i < totalRows - 1 && !existingItemId) {
        form.addPageBreakItem();
      }

      currentIndex = i + 1;
    }

    if (itemIdsToUpdate.length > 0) {
      itemIdsToUpdate.forEach(obj => {
        masterSheet.getRange(obj.row, obj.col).setValue(obj.val);
      });
    }

    setConfigValue(configSheet, 'Last_Processed_Row', currentIndex);

    if (currentIndex >= totalRows) {
      setConfigValue(configSheet, 'Process_Status', 'COMPLETED');
      const publishedUrl = form.getPublishedUrl();
      const editUrl = form.getEditUrl();
      showUrlDialog(publishedUrl, editUrl);
    } else {
      setConfigValue(configSheet, 'Process_Status', 'SUSPENDED');
    }

  } catch (e) {
    console.error(e.stack);
    Browser.msgBox("エラー", e.toString(), Browser.Buttons.OK);
  }
}

function showUrlDialog(pubUrl, editUrl) {
  const htmlOutput = HtmlService
    .createHtmlOutput(`<div style="font-family:sans-serif; padding:10px;">
      <h3>🎉 フォーム生成完了！</h3>
      <p>以下のURLから動作確認してください。</p>
      <p><strong>回答用URL:</strong><br><a href="${pubUrl}" target="_blank">${pubUrl}</a></p>
      <p><strong>編集用URL:</strong><br><a href="${editUrl}" target="_blank">${editUrl}</a></p>
    </div>`)
    .setWidth(450)
    .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '処理完了');
}

function createAttributeSection(form) {
  form.setDescription("回答データを集計するため、ニックネームを入力してください。");

  // ▼▼▼ 追加: 文字数制限の設定 ▼▼▼
  const textValidation = FormApp.createTextValidation()
    .requireTextLengthLessThanOrEqualTo(10) // 10文字以下
    .setHelpText("ニックネームは10文字以内で入力してください。") // エラー時のメッセージ
    .build();

  form.addTextItem()
    .setTitle('ニックネーム (回答者名)')
    .setRequired(true)
    .setValidation(textValidation); // バリデーションを適用
  // ▲▲▲ 追加ここまで ▲▲▲

  const rankItem = form.addMultipleChoiceItem();
  rankItem.setTitle('ランキングへの掲載')
    .setChoices([
      rankItem.createChoice('はい、掲載して構いません'),
      rankItem.createChoice('いいえ、掲載しないでください')
    ])
    .setRequired(true);
  form.addPageBreakItem();
}

function resetConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const masterSheet = ss.getSheetByName('Master');
  setConfigValue(configSheet, 'Form_ID', '');
  setConfigValue(configSheet, 'Form_Edit_Url', '');
  setConfigValue(configSheet, 'Last_Processed_Row', '1');
  setConfigValue(configSheet, 'Process_Status', '');
  if (masterSheet.getLastRow() > 1) {
    masterSheet.getRange(2, 16, masterSheet.getLastRow() - 1, 2).clearContent();
  }
  Browser.msgBox("設定をリセットしました。");
}