/**
 * ==========================================================================
 * 仕入カウント確認報告 - Google Apps Script (GAS) バックエンド完全コード
 * 構成: 商品マスタ / PL取込 / 仕入カウント / 仕入カウント差異確認
 * 機能: 
 *   - 未カウントPL抽出 ＆ カウント済みPL一括削除（軽量化）
 *   - 【仕入カウント差異確認】自動生成・更新（PL番号昇順 / 行ズレ防止 / 商品違い混入新行 / 赤字強調）
 *   - LockServiceによる安全な排他制御
 * ==========================================================================
 * 
 * 【デプロイ＆利用手順】
 * 1. スプレッドシートのメニュー「拡張機能」 > 「Apps Script」を開く
 * 2. 本コードをすべて貼り付けて保存 (Ctrl+S / Cmd+S)
 * 3. スプレッドシートを再読込すると、メニューバーに「【仕入検品】」が追加されます
 * 4. メキシコからExcelが届いたら「PL取込」シートにそのまま貼り付けるだけで即時連携！
 * 5. 右上の「デプロイ」 > 「新しいデプロイ」 > 「ウェブアプリ」
 *    - 次のユーザーとして実行: 「自分」
 *    - アクセスできるユーザー: 「全員」
 * 6. 発行された「ウェブアプリのURL」を検品アプリの「設定」に入力
 */

const SHEET_NAMES = {
  MASTER: '商品マスタ',
  PL_IMPORT: 'PL取込',
  COUNT: '仕入カウント',
  DIFF_CHECK: '仕入カウント差異確認'
};

/**
 * スプレッドシート起動時にカスタムメニューを追加
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('【仕入検品】')
      .addItem('📊「仕入カウント差異確認」シートを再生成/更新', 'menuUpdateDifferenceSheet')
      .addItem('⚡ カウント済みPLデータを一括削除（軽量化）', 'menuPurgeCountedPL')
      .addSeparator()
      .addItem('ℹ️ 全シートの初期セットアップ（ヘッダー作成）', 'setupAllSheets')
      .addToUi();
  } catch (e) {
    // Web App context might not have UI
  }
}

/**
 * メニュー実行用ラッパー（差異確認シート再生成）
 */
function menuUpdateDifferenceSheet() {
  const result = updateDifferenceVerificationSheet();
  try {
    SpreadsheetApp.getUi().alert('【差異確認シート更新完了】\n\n' + result.message);
  } catch (e) {}
}

/**
 * メニュー実行用ラッパー（カウント済みPL一括削除）
 */
function menuPurgeCountedPL() {
  purgeCountedPLFromImportSheet(false);
}

/**
 * Web App GET リクエスト処理
 */
function doGet(e) {
  try {
    const params = e ? e.parameter : {};
    const action = params.action || 'ping';

    // 1. 疎通確認（Ping）
    if (action === 'ping') {
      return createJsonResponse({
        status: 'success',
        message: '仕入カウント確認報告 API は正常に稼働しています',
        timestamp: new Date().toISOString()
      });
    }

    // 2. 商品マスタ取得（コピペ不要・スプレッドシート直結）
    if (action === 'getMaster') {
      const masterList = getMasterData();
      return createJsonResponse({
        status: 'success',
        count: masterList.length,
        data: masterList
      });
    }

    // 3. PL進捗・未カウントPLデータ一覧取得（PLデータ反映）
    if (action === 'getPLList' || action === 'getUncountedPL' || action === 'refreshPLData') {
      const summary = getUncountedPLData();
      return createJsonResponse({
        status: 'success',
        count: summary.plList.length,
        data: summary.plList,
        uncountedCount: summary.uncountedCount,
        countedCount: summary.countedCount
      });
    }

    // 4. 特定PLの明細商品一覧取得（PL番号で抽出 & マスタ結合）
    if (action === 'getPL' || action === 'getPLDetails') {
      const plNo = params.plNo;
      if (!plNo) {
        return createJsonResponse({ status: 'error', message: 'plNo パラメータが必要です' }, 400);
      }
      const plDetails = getPLDetails(plNo);
      return createJsonResponse({
        status: 'success',
        plNo: plNo,
        count: plDetails.length,
        data: plDetails
      });
    }

    // 5. カウント済みPLデータの一括削除（軽量化API）
    if (action === 'purgeCountedPL') {
      const result = purgeCountedPLFromImportSheet(true);
      return createJsonResponse(result);
    }

    // 6. 「仕入カウント差異確認」シートの明示的更新API
    if (action === 'updateDiffSheet' || action === 'generateDiffReport') {
      const result = updateDifferenceVerificationSheet();
      return createJsonResponse(result);
    }

    return createJsonResponse({ status: 'error', message: '不明なアクション: ' + action }, 400);
  } catch (error) {
    return createJsonResponse({
      status: 'error',
      message: error.toString(),
      stack: error.stack
    }, 500);
  }
}

/**
 * Web App POST リクエスト処理
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  const hasLock = lock.tryLock(30000);

  if (!hasLock) {
    return createJsonResponse({
      status: 'error',
      message: '他の端末が書き込み中です。少し待ってから再試行してください。(Lock Timeout)'
    }, 503);
  }

  try {
    let postData = {};
    if (e && e.postData && e.postData.contents) {
      try {
        postData = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        postData = e.parameter || {};
      }
    } else if (e && e.parameter) {
      postData = e.parameter;
    }

    const action = postData.action || '';

    // 検品実績の登録（仕入カウントシートへ保存 ＆ 差異確認シート自動更新）
    if (action === 'submitInspection') {
      const result = saveInspectionResults(postData);
      return createJsonResponse(result);
    }

    // カウント済みPLデータの一括削除（軽量化）
    if (action === 'purgeCountedPL') {
      const result = purgeCountedPLFromImportSheet(true);
      return createJsonResponse(result);
    }

    // 差異確認シートの手動更新トリガー
    if (action === 'updateDiffSheet') {
      const result = updateDifferenceVerificationSheet();
      return createJsonResponse(result);
    }

    return createJsonResponse({ status: 'error', message: '不明なアクション: ' + action }, 400);
  } catch (error) {
    return createJsonResponse({
      status: 'error',
      message: error.toString()
    }, 500);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 商品マスタシートの読み込み
 */
function getMasterData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.MASTER);
  if (!sheet) {
    setupAllSheets();
    sheet = ss.getSheetByName(SHEET_NAMES.MASTER);
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const list = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const itemCode = String(row[0] || '').trim();
    if (!itemCode) continue;

    list.push({
      itemCode: itemCode,
      mexDesign: String(row[1] || '').trim(),
      mexColor: String(row[2] || '').trim(),
      mexSize: String(row[3] || '').trim(),
      jpDesign: String(row[4] || '').trim(),
      jpColor: String(row[5] || '').trim(),
      jpSize: String(row[6] || '').trim()
    });
  }

  return list;
}

/**
 * 「PL取込」シートから未カウントのPL一覧および進捗サマリーを抽出
 */
function getUncountedPLData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let importSheet = ss.getSheetByName(SHEET_NAMES.PL_IMPORT);
  let countSheet = ss.getSheetByName(SHEET_NAMES.COUNT);

  if (!importSheet || !countSheet) {
    setupAllSheets();
    importSheet = ss.getSheetByName(SHEET_NAMES.PL_IMPORT);
    countSheet = ss.getSheetByName(SHEET_NAMES.COUNT);
  }

  // 1. 「仕入カウント」シートからカウント完了済みのPL番号を収集
  const countedPLMap = {};
  if (countSheet && countSheet.getLastRow() > 1) {
    const countValues = countSheet.getDataRange().getValues();
    for (let i = 1; i < countValues.length; i++) {
      const row = countValues[i];
      const plNo = String(row[2] || '').trim(); // C列: PL番号
      const inspector = String(row[1] || '').trim(); // B列: 作業者氏名
      const dateVal = row[0]; // A列: 検品日時
      if (plNo) {
        countedPLMap[plNo] = {
          isCompleted: true,
          inspectorName: inspector,
          completedAt: dateVal instanceof Date ? formatDate(dateVal) : String(dateVal || '')
        };
      }
    }
  }

  // 2. 「PL取込」シートからデータを集計
  const importValues = importSheet.getDataRange().getValues();
  if (importValues.length <= 1) {
    return { plList: [], uncountedCount: 0, countedCount: 0 };
  }

  const plSummaryMap = {};
  let uncountedCount = 0;
  let countedCount = 0;

  for (let i = 1; i < importValues.length; i++) {
    const row = importValues[i];
    const itemCode = String(row[0] || '').trim();
    const plNo = String(row[1] || '').trim();
    if (!itemCode || !plNo) continue;

    const qty = Number(row[8]) || Number(String(row[8] || '').replace(/[^0-9]/g, '')) || 0;

    if (!plSummaryMap[plNo]) {
      const completedInfo = countedPLMap[plNo];
      const isDone = !!completedInfo;
      if (isDone) {
        countedCount++;
      } else {
        uncountedCount++;
      }

      plSummaryMap[plNo] = {
        plNo: plNo,
        totalLines: 0,
        totalPlanQty: 0,
        isCompleted: isDone,
        inspectorName: isDone ? completedInfo.inspectorName : '',
        completedAt: isDone ? completedInfo.completedAt : ''
      };
    }

    plSummaryMap[plNo].totalLines += 1;
    plSummaryMap[plNo].totalPlanQty += qty;
  }

  const plList = [];
  for (const pl in plSummaryMap) {
    plList.push(plSummaryMap[pl]);
  }

  return {
    plList: plList,
    uncountedCount: uncountedCount,
    countedCount: countedCount
  };
}

/**
 * 特定のPL番号に紐づく予定商品明細を「PL取込」シートから抽出し商品マスタと結合
 */
function getPLDetails(targetPlNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let importSheet = ss.getSheetByName(SHEET_NAMES.PL_IMPORT);
  if (!importSheet) {
    setupAllSheets();
    importSheet = ss.getSheetByName(SHEET_NAMES.PL_IMPORT);
  }

  const masterList = getMasterData();
  const masterMap = {};
  for (let i = 0; i < masterList.length; i++) {
    masterMap[masterList[i].itemCode] = masterList[i];
  }

  const values = importSheet.getDataRange().getValues();
  const results = [];
  const cleanTargetPl = String(targetPlNo).trim();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const plNo = String(row[1] || '').trim();
    if (plNo !== cleanTargetPl) continue;

    const itemCode = String(row[0] || '').trim();
    const item = String(row[2] || '').trim();
    const color = String(row[4] || '').trim();
    const size = String(row[5] || '').trim();
    const pallet = String(row[6] || '').trim();
    const qty = Number(row[8]) || Number(String(row[8] || '').replace(/[^0-9]/g, '')) || 0;

    const master = masterMap[itemCode] || {
      itemCode: itemCode,
      mexDesign: item,
      mexColor: color,
      mexSize: size,
      jpDesign: '未登録商品 (' + itemCode + ')',
      jpColor: color || '',
      jpSize: size || ''
    };

    results.push({
      itemCode: itemCode,
      plNo: plNo,
      mexDesign: master.mexDesign || item,
      mexColor: master.mexColor || color,
      mexSize: master.mexSize || size,
      jpDesign: master.jpDesign,
      jpColor: master.jpColor,
      jpSize: master.jpSize,
      planQty: qty,
      palletNo: pallet
    });
  }

  return results;
}

/**
 * 【軽量化・クリア機能】
 * 「仕入カウント」に送信済みのPL行だけを「PL取込」シートから自動一括削除し、
 * 未カウントデータのみを残して動作を軽快に保つ
 */
function purgeCountedPLFromImportSheet(isApiCall) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    const errMsg = '現在他の端末が書き込み処理中です。少し待ってから再試行してください。';
    if (!isApiCall) SpreadsheetApp.getUi().alert(errMsg);
    return { status: 'error', message: errMsg };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let importSheet = ss.getSheetByName(SHEET_NAMES.PL_IMPORT);
    let countSheet = ss.getSheetByName(SHEET_NAMES.COUNT);

    if (!importSheet || !countSheet) {
      throw new Error('必要なシート（PL取込 または 仕入カウント）が見つかりません。');
    }

    // 1. 「仕入カウント」シートに記録済みのPL番号を取得
    const countedPLSet = {};
    if (countSheet.getLastRow() > 1) {
      const countValues = countSheet.getDataRange().getValues();
      for (let i = 1; i < countValues.length; i++) {
        const pl = String(countValues[i][2] || '').trim(); // C列
        if (pl) countedPLSet[pl] = true;
      }
    }

    const lastRow = importSheet.getLastRow();
    const lastCol = Math.max(importSheet.getLastColumn(), 10);
    if (lastRow <= 1) {
      const emptyMsg = '「PL取込」シートにデータが存在しません。';
      if (!isApiCall) SpreadsheetApp.getUi().alert(emptyMsg);
      return { status: 'success', deletedRows: 0, remainingRows: 0, message: emptyMsg };
    }

    // 2. 「PL取込」シートの全行を読み込み、未カウント行だけを抽出
    const allData = importSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const remainingRows = [];
    let deletedCount = 0;

    for (let i = 0; i < allData.length; i++) {
      const row = allData[i];
      const plNo = String(row[1] || '').trim(); // B列
      if (countedPLSet[plNo]) {
        deletedCount++;
      } else {
        remainingRows.push(row);
      }
    }

    if (deletedCount === 0) {
      const noDoneMsg = 'カウント済みのPLデータはありませんでした（すべて未カウントです）。';
      if (!isApiCall) SpreadsheetApp.getUi().alert(noDoneMsg);
      return { status: 'success', deletedRows: 0, remainingRows: remainingRows.length, message: noDoneMsg };
    }

    // 3. データ領域をクリアして、未カウント行だけを再書き込み
    importSheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();

    if (remainingRows.length > 0) {
      importSheet.getRange(2, 1, remainingRows.length, lastCol).setValues(remainingRows);
    }

    const successMsg = 'カウント完了済みのPLデータ ' + deletedCount + ' 行を「PL取込」シートから削除しました。\n（未カウント残数: ' + remainingRows.length + ' 行）';
    if (!isApiCall) {
      SpreadsheetApp.getUi().alert('【軽量化完了】\n\n' + successMsg);
    }

    return {
      status: 'success',
      deletedRows: deletedCount,
      remainingRows: remainingRows.length,
      message: successMsg
    };
  } catch (err) {
    if (!isApiCall) SpreadsheetApp.getUi().alert('エラー: ' + err.message);
    return { status: 'error', message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 検品結果を「仕入カウント」シートへ保存 ＆ 「仕入カウント差異確認」シートを自動更新
 */
function saveInspectionResults(postData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.COUNT);

  if (!sheet) {
    setupAllSheets();
    sheet = ss.getSheetByName(SHEET_NAMES.COUNT);
  }

  const items = postData.items || [];
  if (!Array.isArray(items) || items.length === 0) {
    return { status: 'error', message: '検品アイテムリストが空です。' };
  }

  const inspectorName = String(postData.inspectorName || '未記入').trim();
  const plNo = String(postData.plNo || '').trim();
  const sessionId = String(postData.sessionId || ('SES-' + new Date().getTime()));
  const timestamp = formatDate(new Date());

  const rowsToAppend = [];

  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    const planQty = Number(it.planQty) || 0;
    const actualQty = Number(it.actualQty !== null ? it.actualQty : planQty);
    const diffQty = actualQty - planQty;

    let status = '一致';
    let exceptionNoteFormatted = '';

    if (it.hasException && Array.isArray(it.discrepancies) && it.discrepancies.length > 0) {
      const reasons = [];
      for (let d = 0; d < it.discrepancies.length; d++) {
        const r = it.discrepancies[d].reason;
        if (reasons.indexOf(r) === -1) reasons.push(r);
      }
      status = '差異: ' + reasons.join('/');

      const noteParts = [];
      for (let d = 0; d < it.discrepancies.length; d++) {
        const disc = it.discrepancies[d];
        let part = '【' + disc.reason + '】';
        if (disc.targetItemCode) {
          part += disc.targetItemCode + ' (' + (disc.targetItemName || '') + ' ' + (disc.targetSize || '') + ') ';
        }
        part += (disc.reason === '数量過不足' ? '実数: ' + disc.qty : disc.qty + '点');
        if (disc.note) {
          part += ' [' + disc.note + ']';
        }
        noteParts.push(part);
      }
      exceptionNoteFormatted = noteParts.join(' / ');
    } else if (diffQty !== 0) {
      status = diffQty > 0 ? '超過 (+' + diffQty + ')' : '不足 (' + diffQty + ')';
    }

    const mexNotation = (it.mexDesign || '') + ' / ' + (it.mexColor || '') + ' / ' + (it.mexSize || '');
    const jpNotation = (it.jpDesign || '') + ' / ' + (it.jpColor || '') + ' / ' + (it.jpSize || '');

    // M列(13列目)にパース可能な完全JSON内訳を安全に格納
    const rawDiscrepanciesJson = JSON.stringify(it.discrepancies || []);

    rowsToAppend.push([
      timestamp,
      inspectorName,
      plNo,
      String(it.itemCode || ''),
      mexNotation,
      jpNotation,
      planQty,
      actualQty,
      diffQty,
      status,
      exceptionNoteFormatted,
      sessionId,
      rawDiscrepanciesJson
    ]);
  }

  if (rowsToAppend.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  }

  // 検品結果保存後、「仕入カウント差異確認」シートを自動更新（ハイブリッド構造・PL昇順・赤字強調）
  try {
    updateDifferenceVerificationSheet();
  } catch (err) {
    console.warn('差異確認シート自動更新エラー: ' + err);
  }

  return {
    status: 'success',
    count: rowsToAppend.length,
    plNo: plNo,
    message: 'PL「' + plNo + '」の検品結果 ' + rowsToAppend.length + ' 件を「仕入カウント」に記録し、「仕入カウント差異確認」シートを最新化しました。'
  };
}

/**
 * ==========================================================================
 * 【仕入カウント差異確認】シートの出力 ＆ ハイブリッド構造構築
 * 
 * 仕様要件:
 * 1. PL番号順（昇順・自然順）のソート
 * 2. 行ズレ防止 ＆ 混入ハイブリッド構造:
 *    - 基本は「PL取込（予定）」の1行 ＝ 「差異確認シート」の1行
 *    - 現場で「商品違い（予期せぬ別商品の混入・追加）」が発生した場合は新しい行を追加し、
 *      PL番号を維持したまま実績側にその商品コード・アイテム名・実数を記載（予定数の列は空欄）
 * 3. 見逃し防止の赤字強調:
 *    - 差異・例外が発生した行、または差異内容・例外混入のテキストは文字色を「赤字（#dc2626）」で出力
 * ==========================================================================
 */
function updateDifferenceVerificationSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let diffSheet = ss.getSheetByName(SHEET_NAMES.DIFF_CHECK);
  if (!diffSheet) {
    diffSheet = ss.insertSheet(SHEET_NAMES.DIFF_CHECK);
  }

  const importSheet = ss.getSheetByName(SHEET_NAMES.PL_IMPORT);
  const countSheet = ss.getSheetByName(SHEET_NAMES.COUNT);
  const masterList = getMasterData();
  const masterMap = {};
  for (let i = 0; i < masterList.length; i++) {
    masterMap[masterList[i].itemCode] = masterList[i];
  }

  // 1. 「仕入カウント」シートから最新の検品実績および商品違い（混入）データを収集
  const inspectionMap = {};
  const mixedItemsByPL = {}; // plNo -> array of mixed items

  if (countSheet && countSheet.getLastRow() > 1) {
    const countValues = countSheet.getDataRange().getValues();
    for (let i = 1; i < countValues.length; i++) {
      const row = countValues[i];
      const timestamp = row[0];
      const inspector = String(row[1] || '').trim();
      const plNo = String(row[2] || '').trim();
      const itemCode = String(row[3] || '').trim();
      const planQty = Number(row[6]) || 0;
      const actualQty = Number(row[7]) || 0;
      const diffQty = Number(row[8]) || 0;
      const status = String(row[9] || '').trim();
      const note = String(row[10] || '').trim();
      const jsonStr = String(row[12] || '').trim();

      if (!plNo) continue;

      const key = plNo + '__' + itemCode;
      let parsedDiscrepancies = [];
      if (jsonStr) {
        try { parsedDiscrepancies = JSON.parse(jsonStr); } catch (e) {}
      }

      inspectionMap[key] = {
        timestamp: timestamp instanceof Date ? formatDate(timestamp) : String(timestamp || ''),
        inspector: inspector,
        plNo: plNo,
        itemCode: itemCode,
        planQty: planQty,
        actualQty: actualQty,
        diffQty: diffQty,
        status: status,
        note: note,
        discrepancies: parsedDiscrepancies
      };

      // 「商品違い（混入）」の検出
      if (Array.isArray(parsedDiscrepancies) && parsedDiscrepancies.length > 0) {
        for (let d = 0; d < parsedDiscrepancies.length; d++) {
          const disc = parsedDiscrepancies[d];
          if (disc.reason === '商品違い' && disc.targetItemCode) {
            if (!mixedItemsByPL[plNo]) mixedItemsByPL[plNo] = [];
            mixedItemsByPL[plNo].push({
              plNo: plNo,
              targetItemCode: disc.targetItemCode,
              targetItemName: disc.targetItemName || '',
              targetSize: disc.targetSize || '',
              qty: Number(disc.qty) || 0,
              note: disc.note || '',
              inspector: inspector,
              timestamp: timestamp instanceof Date ? formatDate(timestamp) : String(timestamp || '')
            });
          }
        }
      } else if (note.indexOf('【商品違い】') !== -1) {
        if (!mixedItemsByPL[plNo]) mixedItemsByPL[plNo] = [];
        mixedItemsByPL[plNo].push({
          plNo: plNo,
          targetItemCode: '混入品',
          targetItemName: '別商品混入',
          targetSize: '',
          qty: Math.abs(diffQty) || 1,
          note: note,
          inspector: inspector,
          timestamp: timestamp instanceof Date ? formatDate(timestamp) : String(timestamp || '')
        });
      }
    }
  }

  // 2. 「PL取込」シートから予定データを収集（PL番号ごとにグループ化）
  const plGroups = {};
  const allPLSet = {};

  if (importSheet && importSheet.getLastRow() > 1) {
    const importValues = importSheet.getDataRange().getValues();
    for (let i = 1; i < importValues.length; i++) {
      const row = importValues[i];
      const itemCode = String(row[0] || '').trim();
      const plNo = String(row[1] || '').trim();
      if (!itemCode || !plNo) continue;

      const item = String(row[2] || '').trim();
      const color = String(row[4] || '').trim();
      const size = String(row[5] || '').trim();
      const pallet = String(row[6] || '').trim();
      const qty = Number(row[8]) || Number(String(row[8] || '').replace(/[^0-9]/g, '')) || 0;

      if (!plGroups[plNo]) plGroups[plNo] = [];
      plGroups[plNo].push({
        itemCode: itemCode,
        plNo: plNo,
        item: item,
        color: color,
        size: size,
        pallet: pallet,
        planQty: qty
      });
      allPLSet[plNo] = true;
    }
  }

  // 「仕入カウント」にあるが「PL取込」で既に軽量化削除されたPLも漏れなく包含
  for (const key in inspectionMap) {
    const pl = inspectionMap[key].plNo;
    allPLSet[pl] = true;
    if (!plGroups[pl]) plGroups[pl] = [];
  }

  // 3. PL番号順（昇順・自然順）にソート
  const sortedPLs = Object.keys(allPLSet).sort(function(a, b) {
    const numA = parseFloat(String(a).replace(/[^0-9.]/g, '')) || 0;
    const numB = parseFloat(String(b).replace(/[^0-9.]/g, '')) || 0;
    if (numA !== numB) return numA - numB;
    return String(a).localeCompare(String(b));
  });

  // 4. 行ズレ防止 ＆ 混入ハイブリッド構造の組み立て
  const outputRows = [];
  const fontColors = [];
  const bgColors = [];
  const fontWeights = [];

  const RED_COLOR = '#dc2626';     // 差異・例外・混入強調赤字
  const RED_BG = '#fef2f2';        // 差異行背景
  const DEFAULT_TEXT = '#1e293b';  // 通常文字色
  const DEFAULT_BG = '#ffffff';    // 通常背景

  for (let p = 0; p < sortedPLs.length; p++) {
    const plNo = sortedPLs[p];
    const plannedLines = plGroups[plNo] || [];

    // (A) 基本: 「PL取込（予定）」の1行 ＝ 「差異確認シート」の1行
    for (let l = 0; l < plannedLines.length; l++) {
      const plLine = plannedLines[l];
      const master = masterMap[plLine.itemCode] || {
        jpDesign: plLine.item || plLine.itemCode,
        jpColor: plLine.color || '',
        jpSize: plLine.size || ''
      };

      const mexNotation = (plLine.item || '') + ' / ' + (plLine.color || '') + ' / ' + (plLine.size || '');
      const jpNotation = (master.jpDesign || '') + ' / ' + (master.jpColor || '') + ' / ' + (master.jpSize || '');

      const key = plNo + '__' + plLine.itemCode;
      const insp = inspectionMap[key];

      let actualQtyVal = '';
      let diffQtyVal = '';
      let statusVal = '未検品';
      let actualItemCodeVal = plLine.itemCode;
      let actualItemNameVal = master.jpDesign;
      let noteVal = '';
      let workerVal = '';
      let dateVal = '';
      let isDiff = false;

      if (insp) {
        actualQtyVal = insp.actualQty;
        diffQtyVal = insp.diffQty;
        statusVal = insp.status;
        noteVal = insp.note;
        workerVal = insp.inspector;
        dateVal = insp.timestamp;

        // サイズ違いがあれば実績商品コード/品名を反映
        if (insp.discrepancies && insp.discrepancies.length > 0) {
          const sizeDisc = insp.discrepancies.find(function(d) { return d.reason === 'サイズ違い'; });
          if (sizeDisc && sizeDisc.targetItemCode) {
            actualItemCodeVal = sizeDisc.targetItemCode;
            actualItemNameVal = (sizeDisc.targetItemName || '') + ' (' + (sizeDisc.targetSize || '') + ')';
          }
        }

        // 差異判定 (予定数と異なる、または差異ステータス)
        if (diffQtyVal !== 0 || statusVal.indexOf('差異') !== -1 || statusVal.indexOf('超過') !== -1 || statusVal.indexOf('不足') !== -1) {
          isDiff = true;
        }
      }

      const row = [
        plNo,
        plLine.itemCode,
        master.jpDesign,
        (master.jpColor || '') + ' ' + (master.jpSize || ''),
        mexNotation,
        plLine.planQty,
        actualQtyVal,
        diffQtyVal,
        statusVal,
        actualItemCodeVal,
        actualItemNameVal,
        noteVal,
        workerVal,
        dateVal
      ];
      outputRows.push(row);

      // 見逃し防止の赤字強調
      const rowColors = [];
      const rowBgs = [];
      const rowWeights = [];
      for (let c = 0; c < row.length; c++) {
        if (isDiff) {
          rowColors.push(RED_COLOR);
          rowBgs.push(RED_BG);
          rowWeights.push(c === 0 || c === 7 || c === 8 ? 'bold' : 'normal');
        } else {
          rowColors.push(DEFAULT_TEXT);
          rowBgs.push(DEFAULT_BG);
          rowWeights.push('normal');
        }
      }
      fontColors.push(rowColors);
      bgColors.push(rowBgs);
      fontWeights.push(rowWeights);
    }

    // (B) 混入ハイブリッド構造: 現場で「商品違い（予期せぬ別商品の混入・追加）」が発生した場合は新しい行を追加
    // PL番号を維持したまま実績側にその商品コード・アイテム名・実数を記載（予定数の列は空欄）
    const mixedList = mixedItemsByPL[plNo] || [];
    for (let m = 0; m < mixedList.length; m++) {
      const mixed = mixedList[m];
      const mixedRow = [
        plNo,                           // A: PL番号（同一PLを維持）
        '',                             // B: 商品コード(予定) ➔ 空欄
        '',                             // C: 日本品名(予定) ➔ 空欄
        '',                             // D: カラー/サイズ(予定) ➔ 空欄
        '',                             // E: メキシコ表記(予定) ➔ 空欄
        '',                             // F: 予定数量 ➔ 空欄
        mixed.qty,                      // G: 検品実績数
        '+' + mixed.qty,                // H: 差異数量
        '差異: 商品違い(混入)',          // I: 判定ステータス
        mixed.targetItemCode,           // J: 商品コード(実績/混入)
        mixed.targetItemName + (mixed.targetSize ? ' (' + mixed.targetSize + ')' : ''), // K: 品名・サイズ(実績/混入)
        '【商品違い混入】' + (mixed.note ? ' [' + mixed.note + ']' : ''), // L: 差異内訳・理由メモ
        mixed.inspector,                // M: 作業者氏名
        mixed.timestamp                 // N: 検品日時
      ];
      outputRows.push(mixedRow);

      // 混入行は常に目立つ赤字 ＆ 強調背景
      const mixedColors = [];
      const mixedBgs = [];
      const mixedWeights = [];
      for (let c = 0; c < mixedRow.length; c++) {
        mixedColors.push(RED_COLOR);
        mixedBgs.push('#fee2e2');
        mixedWeights.push('bold');
      }
      fontColors.push(mixedColors);
      bgColors.push(mixedBgs);
      fontWeights.push(mixedWeights);
    }
  }

  // 5. シートへ書き出し
  diffSheet.clear();

  const headers = [
    'PL番号',
    '商品コード(予定)',
    '日本品名(予定)',
    'カラー/サイズ(予定)',
    'メキシコ表記(予定)',
    '予定数量',
    '検品実績数',
    '差異数量',
    '判定ステータス',
    '商品コード(実績/混入)',
    '品名・サイズ(実績/混入)',
    '差異内訳・理由メモ',
    '作業者氏名',
    '検品日時'
  ];

  diffSheet.appendRow(headers);
  diffSheet.getRange(1, 1, 1, headers.length)
    .setBackground('#0f172a')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  diffSheet.setFrozenRows(1);

  if (outputRows.length > 0) {
    const dataRange = diffSheet.getRange(2, 1, outputRows.length, headers.length);
    dataRange.setValues(outputRows);
    dataRange.setFontColors(fontColors);
    dataRange.setBackgrounds(bgColors);
    dataRange.setFontWeights(fontWeights);

    // 予定数(F列), 実績数(G列), 差異数(H列) の数値セル配置を右揃え
    diffSheet.getRange(2, 6, outputRows.length, 3).setHorizontalAlignment('right');
    // PL番号(A列), 判定ステータス(I列) を中央揃え
    diffSheet.getRange(2, 1, outputRows.length, 1).setHorizontalAlignment('center');
    diffSheet.getRange(2, 9, outputRows.length, 1).setHorizontalAlignment('center');
  }

  return {
    status: 'success',
    rowCount: outputRows.length,
    message: '「仕入カウント差異確認」シートを更新しました（全 ' + outputRows.length + ' 行）。'
  };
}

/**
 * 全シートの自動初期セットアップ（ヘッダー作成）
 */
function setupAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 商品マスタ
  let mSheet = ss.getSheetByName(SHEET_NAMES.MASTER);
  if (!mSheet) {
    mSheet = ss.insertSheet(SHEET_NAMES.MASTER);
    mSheet.appendRow(['商品コード', 'メキシコデザイン名', 'メキシコカラー', 'メキシコサイズ', '日本デザイン名', '日本カラー', '日本サイズ']);
    mSheet.getRange(1, 1, 1, 7).setBackground('#0f172a').setFontColor('#ffffff').setFontWeight('bold');
    mSheet.setFrozenRows(1);
  }

  // 2. PL取込（Excelを直接貼り付けるシート）
  let impSheet = ss.getSheetByName(SHEET_NAMES.PL_IMPORT);
  if (!impSheet) {
    impSheet = ss.insertSheet(SHEET_NAMES.PL_IMPORT);
    impSheet.appendRow(['商品コード', 'PL番号/ID', 'Item', 'description', 'color', 'S', 'pallet', 'carton', 'qty', '取込ステータス']);
    impSheet.getRange(1, 1, 1, 10).setBackground('#0f172a').setFontColor('#ffffff').setFontWeight('bold');
    impSheet.setFrozenRows(1);
  }

  // 3. 仕入カウント（検品実績蓄積）
  let countSheet = ss.getSheetByName(SHEET_NAMES.COUNT);
  if (!countSheet) {
    countSheet = ss.insertSheet(SHEET_NAMES.COUNT);
    countSheet.appendRow([
      '検品日時',
      '作業者氏名',
      'PL番号',
      '商品コード',
      'メキシコ表記(Item / color / S)',
      '日本表記(品名 / 色 / サイズ)',
      '予定数量',
      '検品実績数',
      '差異数量',
      '判定ステータス',
      '例外報告内容・メモ',
      'セッションID',
      '差異詳細JSON'
    ]);
    countSheet.getRange(1, 1, 1, 13).setBackground('#0f172a').setFontColor('#ffffff').setFontWeight('bold');
    countSheet.setFrozenRows(1);
  }

  // 4. 仕入カウント差異確認（ハイブリッド対比表）
  let diffSheet = ss.getSheetByName(SHEET_NAMES.DIFF_CHECK);
  if (!diffSheet) {
    diffSheet = ss.insertSheet(SHEET_NAMES.DIFF_CHECK);
    diffSheet.appendRow([
      'PL番号',
      '商品コード(予定)',
      '日本品名(予定)',
      'カラー/サイズ(予定)',
      'メキシコ表記(予定)',
      '予定数量',
      '検品実績数',
      '差異数量',
      '判定ステータス',
      '商品コード(実績/混入)',
      '品名・サイズ(実績/混入)',
      '差異内訳・理由メモ',
      '作業者氏名',
      '検品日時'
    ]);
    diffSheet.getRange(1, 1, 1, 14).setBackground('#0f172a').setFontColor('#ffffff').setFontWeight('bold');
    diffSheet.setFrozenRows(1);
  }

  try {
    SpreadsheetApp.getUi().alert('全シート（商品マスタ・PL取込・仕入カウント・仕入カウント差異確認）の初期セットアップが完了しました。');
  } catch (e) {}
}

function formatDate(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}

function createJsonResponse(data, statusCode) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
