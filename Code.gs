// ============================================================
// 機体返却管理システム - Google Apps Script 完全版
// Code.gs
//
// 【スクリプトプロパティに設定が必要なもの】
//   MAIN_SHEET_ID   : このスプレッドシートのID
//   KANRI_SHEET_ID  : 管理表ファイルのスプレッドシートID
//   API_TOKEN       : PWAアプリからの認証トークン（任意の文字列）
//   DRIVE_FOLDER_ID : 伝票画像の保存先DriveフォルダID
//   VISION_API_KEY  : Google Cloud Vision APIキー
// ============================================================


// ============================================================
// 定数
// ============================================================

const STATUS_HEADER_ROWS    = 4;  // statusシート：1タイトル 2注釈 3凡例 4ヘッダー
const UNRETURNED_HEADER_ROWS = 4; // unreturnedシート：同上
const MASTER_HEADER_ROWS    = 3;  // masterシート：1タイトル 2注釈 3ヘッダー

// statusシートの列インデックス（1始まり）
const ST_COL = {
  ITEM_ID:   1,  // A: 機体ID
  STATUS:    2,  // B: 最新状態（出荷/返却）
  TIMESTAMP: 3,  // C: 最終操作日時
  TRACKING:  4,  // D: 追跡番号
  WORKER:    5,  // E: 作業者
  CUSTOMER:  6,  // F: ★顧客名（管理表から取得）
  PRODUCT:   7,  // G: ★商品名・プラン名（管理表から取得）
  DEADLINE:  8,  // H: ★返却期限（管理表から取得）
  MODEL:     9,  // I: ★機種名（masterから取得）
  OVERDUE:   10, // J: 延滞フラグ
};

// unreturnedシートの列インデックス（1始まり）
const UR_COL = {
  ITEM_ID:   1,  // A: 機体ID
  MODEL:     2,  // B: 機種名
  SHIP_TIME: 3,  // C: 出荷日時
  CUSTOMER:  4,  // D: ★顧客名
  PRODUCT:   5,  // E: ★商品名・プラン名
  DEADLINE:  6,  // F: ★返却期限
  DAYS:      7,  // G: 経過日数
  URGENCY:   8,  // H: 緊急度
  NOTES:     9,  // I: 備考（手動入力欄）
};


// ============================================================
// 1. PWAアプリからのPOSTを受け取るエントリーポイント
// ============================================================

function doPost(e) {
  try {
    const props = PropertiesService.getScriptProperties().getProperties();
    const data  = JSON.parse(e.postData.contents);

    // 認証チェック
    if (data.token !== props.API_TOKEN) {
      return jsonRes({ success: false, error: 'Unauthorized' });
    }

    const result = processRecord(data, props);
    return jsonRes({ success: true, ...result });

  } catch (err) {
    console.error(err.stack);
    return jsonRes({ success: false, error: err.message });
  }
}


// ============================================================
// 2. メイン処理
//    ① OCR → ② log記録 → ③ 管理表照合 → ④ status更新 → ⑤ unreturned再生成
// ============================================================

function processRecord(data, props) {
  const ss       = SpreadsheetApp.openById(props.MAIN_SHEET_ID);
  const logSheet = ss.getSheetByName('log');

  // ① 伝票画像のOCR処理
  let trackingNumber = data.trackingNumber || '';
  let imageUrl       = '';
  let ocrRawText     = '';

  if (data.imageBase64) {
    const ocr  = runOCR(data.imageBase64, props);
    ocrRawText = ocr.rawText;
    imageUrl   = saveImageToDrive(data.imageBase64, props);
    // OCRで取得できた追跡番号を優先、取れなければアプリ送信値を使用
    if (ocr.trackingNumber) trackingNumber = ocr.trackingNumber;
  }

  // ② logシートに記録（機体IDごとに1行ずつ追記）
  const timestamp = new Date();
  const logRows   = (data.itemIds || []).map(itemId => [
    timestamp,             // A: タイムスタンプ
    data.workerName    || '', // B: 作業者名
    data.operationType || '', // C: 操作種別（出荷/返却）
    itemId,                // D: 機体ID
    trackingNumber,        // E: 追跡番号（OCR取得）
    data.notes         || '', // F: 備考
    imageUrl,              // G: 画像URL
    ocrRawText,            // H: OCR生テキスト
  ]);

  if (logRows.length > 0) {
    logSheet
      .getRange(logSheet.getLastRow() + 1, 1, logRows.length, logRows[0].length)
      .setValues(logRows);
  }

  // ③ 管理表から顧客情報を照合
  //    追跡番号が出荷追跡番号・返送追跡番号のどちらに一致するかも両方チェック
  const kanriInfo = lookupKanri(trackingNumber, props);

  // ④ statusシートを更新
  updateStatus(
    ss, data.itemIds, data.operationType,
    timestamp, trackingNumber, data.workerName, kanriInfo
  );

  // ⑤ unreturnedシートを再生成
  refreshUnreturned(ss);

  return {
    trackingNumber,
    customerName:   kanriInfo?.customerName   || '',
    productName:    kanriInfo?.productName    || '',
    returnDeadline: kanriInfo?.returnDeadline || '',
    recordCount:    logRows.length,
  };
}


// ============================================================
// 3. Google Vision API による OCR
// ============================================================

function runOCR(base64Image, props) {
  const url  = `https://vision.googleapis.com/v1/images:annotate?key=${props.VISION_API_KEY}`;
  const body = {
    requests: [{
      image:    { content: base64Image },
      features: [{ type: 'TEXT_DETECTION', maxResults: 1 }],
    }],
  };

  try {
    const res     = UrlFetchApp.fetch(url, {
      method:      'post',
      contentType: 'application/json',
      payload:     JSON.stringify(body),
      muteHttpExceptions: true,
    });
    const result  = JSON.parse(res.getContentText());
    const rawText = result.responses?.[0]?.fullTextAnnotation?.text || '';

    // ヤマト運輸の追跡番号：12桁数字
    const match   = rawText.match(/\b(\d{12})\b/);

    return {
      rawText,
      trackingNumber: match ? match[1] : '',
    };
  } catch (e) {
    console.error('OCRエラー:', e.message);
    return { rawText: '', trackingNumber: '' };
  }
}


// ============================================================
// 4. 伝票画像を Google Drive に保存してURLを返す
// ============================================================

function saveImageToDrive(base64Image, props) {
  try {
    const blob   = Utilities.newBlob(
      Utilities.base64Decode(base64Image),
      'image/jpeg',
      `slip_${Date.now()}.jpg`
    );
    const folder = DriveApp.getFolderById(props.DRIVE_FOLDER_ID);
    const file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    console.error('Drive保存エラー:', e.message);
    return '';
  }
}


// ============================================================
// 5. 管理表を追跡番号で照合して顧客情報を返す
//
//  管理表の構造：
//    1行目 : 日付ヘッダー（B列以降、右方向に日付が並ぶ）
//    2行目 : 件数（=COUNTA(X4:X6873) ← その列自身の4行目以降を参照）
//    3行目〜: 1セル＝1件の出荷情報テキスト（空セルはスキップ）
//
//  テキスト例（1セルに改行区切りで全情報が入っている）：
//    "※備考テキスト
//     5.41 ：商品：GoPro HERO8 2泊3日プラン [格安レンタル]
//     ：名前：山田　太郎
//     ：期間：2025年12月10日 / 2025年12月12日
//     ：時間帯：14-16時
//     ：送付先：763-0022　香川県丸亀市...
//     ：出荷追跡番号：194703491496
//     ：返送追跡番号：500072462105
//     ：オプション：なし
//     ：電話番号：09012345678"
// ============================================================

function lookupKanri(trackingNumber, props) {
  if (!trackingNumber) return null;

  let kanriSS;
  try {
    kanriSS = SpreadsheetApp.openById(props.KANRI_SHEET_ID);
  } catch (e) {
    console.warn('管理表ファイルを開けません:', e.message);
    return null;
  }

  const sheet  = kanriSS.getSheets()[0];
  const values = sheet.getDataRange().getValues();

  // B列（index 1）以降を日付列としてスキャン
  for (let col = 1; col < values[0].length; col++) {

    // 3行目（index 2）から下に向かって空セル以外をループ
    for (let row = 2; row < values.length; row++) {
      const cellText = String(values[row][col] || '').trim();
      if (!cellText) continue; // 空セルはスキップ

      const parsed = parseKanriText(cellText);

      // 出荷追跡番号 OR 返送追跡番号のどちらかが一致したら返す
      if (
        parsed.shipTracking   === trackingNumber ||
        parsed.returnTracking === trackingNumber
      ) {
        console.log(`管理表で一致: 列${col + 1} 行${row + 1} 顧客:${parsed.customerName}`);
        return parsed;
      }
    }
  }

  console.warn(`管理表で追跡番号[${trackingNumber}]が見つかりませんでした`);
  return null;
}


// ============================================================
// 6. 管理表テキストを正規表現で解析
// ============================================================

function parseKanriText(text) {
  // 「：キー：値」形式から値を取得するヘルパー
  // 値の終端は「：」「改行」のいずれか
  const get = (key) => {
    const m = text.match(new RegExp('：' + key + '：([^：\n\r]+)'));
    return m ? m[1].trim() : '';
  };

  // 期間：「2025年12月10日 / 2025年12月12日」→ 2つ目が返却期限
  const periodMatch = text.match(
    /：期間：(\d{4}年\d{2}月\d{2}日)\s*\/\s*(\d{4}年\d{2}月\d{2}日)/
  );

  return {
    productName:    get('商品'),
    customerName:   get('名前'),
    returnDeadline: periodMatch ? jpDateToISO(periodMatch[2]) : '', // 例: "2025-12-12"
    rentalStart:    periodMatch ? jpDateToISO(periodMatch[1]) : '', // 例: "2025-12-10"
    shipTracking:   get('出荷追跡番号'),
    returnTracking: get('返送追跡番号'),
    options:        get('オプション'),
    phone:          get('電話番号'),
    address:        get('送付先'),
  };
}

// "2025年12月12日" → "2025-12-12"
function jpDateToISO(jpDate) {
  const m = jpDate.match(/(\d{4})年(\d{2})月(\d{2})日/);
  return m ? `${m[1]}-${m[2]}-${m[3]}` : '';
}


// ============================================================
// 7. statusシートを更新（同一機体IDは上書き、新規は末尾に追加）
// ============================================================

function updateStatus(ss, itemIds, operationType, timestamp,
                      trackingNumber, workerName, kanriInfo) {
  const sheet     = ss.getSheetByName('status');
  const allValues = sheet.getDataRange().getValues();

  // 機体ID（A列）→ Excelの行番号 のマップを作成
  const idToRow = {};
  for (let i = STATUS_HEADER_ROWS; i < allValues.length; i++) {
    const id = allValues[i][ST_COL.ITEM_ID - 1];
    if (id) idToRow[id] = i + 1; // 1-indexed
  }

  // masterシートから {機体ID: 機種名} マップを取得
  const modelMap = getMasterModelMap(ss);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  (itemIds || []).forEach(itemId => {
    const model    = modelMap[itemId] || '';
    const deadline = kanriInfo?.returnDeadline || '';

    // 延滞フラグ（出荷中かつ返却期限を超過している場合のみ）
    let overdueFlag = '';
    if (operationType === '出荷' && deadline) {
      const deadDate = new Date(deadline);
      deadDate.setHours(0, 0, 0, 0);
      if (today > deadDate) overdueFlag = '⚠️ 延滞';
    }

    const newRow = [
      itemId,                           // A: 機体ID
      operationType,                    // B: 最新状態
      timestamp,                        // C: 最終操作日時
      trackingNumber,                   // D: 追跡番号
      workerName,                       // E: 作業者
      kanriInfo?.customerName   || '',  // F: ★顧客名
      kanriInfo?.productName    || '',  // G: ★商品名
      deadline,                         // H: ★返却期限
      model,                            // I: ★機種名
      overdueFlag,                      // J: 延滞フラグ
    ];

    if (idToRow[itemId]) {
      // 既存行を上書き
      sheet
        .getRange(idToRow[itemId], 1, 1, newRow.length)
        .setValues([newRow]);
    } else {
      // 新規行を末尾に追加
      sheet.appendRow(newRow);
    }
  });
}


// ============================================================
// 8. masterシートから {機体ID: 機種名} のマップを生成
// ============================================================

function getMasterModelMap(ss) {
  const sheet = ss.getSheetByName('master');
  if (!sheet) return {};

  const values = sheet.getDataRange().getValues();
  const map    = {};

  // MASTER_HEADER_ROWS行目までがヘッダー、それ以降がデータ
  for (let i = MASTER_HEADER_ROWS; i < values.length; i++) {
    const id    = values[i][0]; // A列: 機体ID
    const model = values[i][1]; // B列: 機種名
    if (id) map[id] = model;
  }
  return map;
}


// ============================================================
// 9. unreturnedシートを再生成
//    statusシートの「最新状態＝出荷」の行のみ抽出して書き込む
// ============================================================

function refreshUnreturned(ss) {
  const statusSheet     = ss.getSheetByName('status');
  const unreturnedSheet = ss.getSheetByName('unreturned');
  const statusData      = statusSheet.getDataRange().getValues();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // 出荷中のデータを抽出・緊急度を計算
  const newRows = [];
  for (let i = STATUS_HEADER_ROWS; i < statusData.length; i++) {
    const row = statusData[i];
    if (!row[ST_COL.ITEM_ID - 1])        continue; // 機体IDが空はスキップ
    if (row[ST_COL.STATUS - 1] !== '出荷') continue; // 返却済みはスキップ

    const deadline = row[ST_COL.DEADLINE - 1];
    let daysStr    = '';
    let urgency    = '⚪ 余裕あり';

    if (deadline) {
      const deadDate = new Date(deadline);
      deadDate.setHours(0, 0, 0, 0);
      const delta = Math.floor((today - deadDate) / 86400000);

      if      (delta > 0)  { daysStr = `+${delta}日 超過`; urgency = '🔴 超過'; }
      else if (delta === 0) { daysStr = '本日期限';         urgency = '🟠 本日期限'; }
      else if (delta >= -3) { daysStr = `あと${-delta}日`;  urgency = '🟡 3日以内'; }
      else                  { daysStr = `あと${-delta}日`;  urgency = '⚪ 余裕あり'; }
    }

    newRows.push([
      row[ST_COL.ITEM_ID   - 1], // A: 機体ID
      row[ST_COL.MODEL     - 1], // B: 機種名
      row[ST_COL.TIMESTAMP - 1], // C: 出荷日時
      row[ST_COL.CUSTOMER  - 1], // D: ★顧客名
      row[ST_COL.PRODUCT   - 1], // E: ★商品名
      deadline,                   // F: ★返却期限
      daysStr,                    // G: 経過日数
      urgency,                    // H: 緊急度
      '',                         // I: 備考（手動入力欄のため空で上書きしない工夫は下記）
    ]);
  }

  // 既存のデータ行をクリア（ヘッダー行は保持、備考列Iは保持）
  const lastRow = unreturnedSheet.getLastRow();
  if (lastRow > UNRETURNED_HEADER_ROWS) {
    // A〜H列のみクリア（I列の備考は手動入力のため保持）
    unreturnedSheet
      .getRange(UNRETURNED_HEADER_ROWS + 1, 1, lastRow - UNRETURNED_HEADER_ROWS, 8)
      .clearContent();
  }

  // 新データを書き込み（A〜H列のみ、I列備考は書き込まない）
  if (newRows.length > 0) {
    // I列を除いた8列分だけ書き込む
    const writeData = newRows.map(r => r.slice(0, 8));
    unreturnedSheet
      .getRange(UNRETURNED_HEADER_ROWS + 1, 1, writeData.length, 8)
      .setValues(writeData);
  }
}


// ============================================================
// 10. 毎日AM6:00に自動実行：延滞フラグの一括更新
//     → setup()でトリガーを登録すること
// ============================================================

function dailyOverdueCheck() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const ss    = SpreadsheetApp.openById(props.MAIN_SHEET_ID);

  updateAllOverdueFlags(ss);
  refreshUnreturned(ss);
  console.log('✅ 日次延滞チェック完了:', new Date());
}

function updateAllOverdueFlags(ss) {
  const sheet  = ss.getSheetByName('status');
  const values = sheet.getDataRange().getValues();
  const today  = new Date();
  today.setHours(0, 0, 0, 0);

  for (let i = STATUS_HEADER_ROWS; i < values.length; i++) {
    if (!values[i][ST_COL.ITEM_ID - 1])          continue; // IDなし
    if (values[i][ST_COL.STATUS - 1] !== '出荷') continue; // 返却済みはスキップ

    const deadline = values[i][ST_COL.DEADLINE - 1];
    if (!deadline) continue;

    const deadDate = new Date(deadline);
    deadDate.setHours(0, 0, 0, 0);

    const flag = today > deadDate ? '⚠️ 延滞' : '';
    sheet.getRange(i + 1, ST_COL.OVERDUE).setValue(flag);
  }
}


// ============================================================
// 11. 初期セットアップ（初回のみ手動で1回実行する）
// ============================================================

function setup() {
  // 既存トリガーを削除してから再登録（重複防止）
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'dailyOverdueCheck')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('dailyOverdueCheck')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  console.log('✅ セットアップ完了：毎日AM6:00に延滞チェックが自動実行されます');
}


// ============================================================
// 12. ユーティリティ
// ============================================================

function jsonRes(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================================
// 13. 手動テスト用（GASエディタから実行して動作確認できる）
// ============================================================

function testLookupKanri() {
  const props = PropertiesService.getScriptProperties().getProperties();
  // 実際に管理表に存在する追跡番号に書き換えてテスト
  const result = lookupKanri('194703491496', props);
  console.log('照合結果:', JSON.stringify(result, null, 2));
}

function testParseKanriText() {
  const sample = `※2台 ※コントローラー2つセットも一緒
5.41 ：商品：SONY PlayStation4 本体 500GB 3日間～ ソニー[格安レンタル] ：名前：大廣　太郎 ：期間：2025年12月12日 / 2025年12月14日 ：時間帯：14-16時 ：送付先：763-0022　香川県丸亀市浜町１０−１　ホテル・アルファーワン丸亀　フロント気付 ：出荷追跡番号：194703491496 ：返送追跡番号：500072462105 ：オプション：コントローラーdualshock 2つセット ：電話番号：08029111022`;

  const result = parseKanriText(sample);
  console.log('解析結果:', JSON.stringify(result, null, 2));
  /*
  期待値：
  {
    "productName":    "SONY PlayStation4 本体 500GB 3日間～ ソニー[格安レンタル] ",
    "customerName":   "大廣　太郎 ",
    "returnDeadline": "2025-12-14",
    "rentalStart":    "2025-12-12",
    "shipTracking":   "194703491496",
    "returnTracking": "500072462105",
    "options":        "コントローラーdualshock 2つセット ",
    "phone":          "08029111022",
    "address":        "763-0022　香川県丸亀市浜町１０−１　..."
  }
  */
}
