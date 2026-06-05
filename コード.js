const SPREADSHEET_ID = '16hYqL1m52cvrfH66A1k5YmwymuzvFtS70BWJtCYNyTw';

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('材料検索システム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function getMaterialData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('シート2');
    if (!sheet) throw new Error("シート『シート2』が見つかりません。");

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return JSON.stringify({ headers: [], rows: [] });

    const headers = values[0];
    const validRows = values.slice(1).filter(row => {
      const pattern  = String(row[0] || '').trim();
      const wireSize = String(row[1] || '').trim();
      const category = String(row[2] || '').trim();
      return pattern !== '' || wireSize !== '' || category !== '';
    });

    return JSON.stringify({
      headers: headers.map(h => String(h).trim()),
      rows: validRows.map(row => row.map(cell => String(cell).trim()))
    });
  } catch (err) {
    throw new Error('データ取得エラー: ' + err.message);
  }
}

function logSearch(searchCount, user) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logSheet = ss.getSheetByName('Log');
    if (logSheet) {
      const ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
      logSheet.appendRow([ts, String(searchCount), String(user || '')]);
    }
  } catch (err) {
    console.error('ログ記録エラー:', err.message);
  }
}
