/**
 * Hub連携機能 (Config自動参照版)
 * 修正点: Configシートの Portal_Url と Dashboard_Url を読み込み、初期値としてセットする
 */

function showRegisterDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig(ss.getSheetByName('Config'));

  // 1. Hub ID の自動取得
  // Configの Portal_Url からIDを抽出する (なければ Hub_Sheet_Id を使う)
  let defaultHubId = "";
  if (config['Portal_Url']) {
    defaultHubId = getFileIdFromUrl(config['Portal_Url']) || "";
  }
  if (!defaultHubId && config['Hub_Sheet_Id']) {
    defaultHubId = config['Hub_Sheet_Id'];
  }

  // 2. Dashboard URL の自動取得
  let defaultDashUrl = config['Dashboard_Url'] || "";
  
  // もしConfigになくても、スクリプトから取得できる場合は試みる
  if (!defaultDashUrl) {
    try { defaultDashUrl = ScriptApp.getService().getUrl(); } catch(e){}
  }

  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:sans-serif; padding:10px;">
      <p>この問題集を「ポータルサイト」に登録します。<br>
      <small style="color:#666;">Configシートの設定値が自動入力されています。</small></p>
      
      <form id="regForm">
        <label style="font-weight:bold;">1. HubスプシのID:</label><br>
        <input type="text" id="hubId" style="width:100%; margin-bottom:10px;" 
               value="${defaultHubId}" placeholder="スプレッドシートID"><br>
        
        <label style="font-weight:bold;">2. ダッシュボードURL:</label><br>
        <input type="text" id="dashUrl" style="width:100%; margin-bottom:10px;" 
               value="${defaultDashUrl}" placeholder=".../exec"><br>
        
        <label style="font-weight:bold;">3. 説明文:</label><br>
        <textarea id="desc" rows="3" style="width:100%; margin-bottom:10px;"></textarea><br>
        
        <label style="font-weight:bold;">4. カテゴリ:</label><br>
        <input type="text" id="cat" style="width:100%; margin-bottom:10px;" value="未分類"><br>
        
        <label style="font-weight:bold;">5. サムネイル画像URL:</label><br>
        <input type="text" id="thumb" style="width:100%; margin-bottom:20px;"><br>
        
        <button type="button" onclick="runRegister()" style="background:#3b82f6; color:white; border:none; padding:10px; width:100%; border-radius:4px; cursor:pointer;">登録・更新する</button>
      </form>
      <div id="status" style="margin-top:10px; color:#666; text-align:center;"></div>
      
      <script>
        function runRegister() {
          const status = document.getElementById('status');
          const hubId = document.getElementById('hubId').value;
          const dashUrl = document.getElementById('dashUrl').value;
          const desc = document.getElementById('desc').value;
          const cat = document.getElementById('cat').value;
          const thumb = document.getElementById('thumb').value;
          
          if(!hubId || !dashUrl) { 
            status.textContent = "❌ Hub IDとダッシュボードURLは必須です"; 
            return; 
          }
          
          const btn = document.querySelector('button');
          btn.disabled = true;
          status.textContent = "送信中...";
          
          google.script.run
            .withSuccessHandler(function(res) { 
              status.textContent = "✅ 登録完了！"; 
              setTimeout(()=>google.script.host.close(), 2000); 
            })
            .withFailureHandler(function(err) { 
              status.textContent = "❌ エラー: " + err; 
              btn.disabled = false;
            })
            .registerToHub(hubId, dashUrl, desc, cat, thumb);
        }
      </script>
    </div>
  `).setWidth(450).setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'ポータル登録');
}

function registerToHub(hubId, dashboardUrl, description, category, thumbUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const title = ss.getName();
  const config = getConfig(configSheet); // Utilities

  let formUrl = "";
  if (config['Form_ID']) {
    try { formUrl = FormApp.openById(config['Form_ID']).getPublishedUrl(); } catch(e) {}
  }

  let hubSs;
  try { hubSs = SpreadsheetApp.openById(hubId); } catch(e) { throw new Error("Hubスプシが見つかりません。IDを確認してください。"); }

  const dirSheet = hubSs.getSheetByName('Directory');
  if (!dirSheet) throw new Error("HubにDirectoryシートがありません");

  const data = dirSheet.getDataRange().getValues();
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === title || data[i][4] === dashboardUrl) {
      foundRow = i + 1;
      break;
    }
  }

  const record = [title, description, category, thumbUrl, dashboardUrl, formUrl, '公開'];
  if (foundRow > 0) dirSheet.getRange(foundRow, 1, 1, record.length).setValues([record]);
  else dirSheet.appendRow(record);
}