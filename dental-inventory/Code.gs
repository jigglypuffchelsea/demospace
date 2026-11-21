/**
 * 牙材庫存登記 — 手機登記表單（Google Apps Script Web App）
 * 綁定在「牙材庫存登記」試算表：擴充功能 → Apps Script，貼上本檔與 Index.html。
 */

const SHEET_ITEMS = '材料主檔';
const SHEET_LOG = '異動紀錄';
const SHEET_BATCH = '批號效期';
const SHEET_SETTINGS = '設定';
const TYPES = ['進貨', '領用', '報廢', '盤點調整'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('牙材登記')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 表單初始資料：材料、經手人、各材料在庫批號（依效期由近到遠） */
function getFormData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const itemRows = ss.getSheetByName(SHEET_ITEMS).getDataRange().getValues().slice(1);
  const items = itemRows
    .filter(r => r[0] !== '')
    .map(r => ({ id: String(r[0]), name: String(r[1]), unit: String(r[4]), stock: Number(r[7]) || 0 }));

  const staff = ss.getSheetByName(SHEET_SETTINGS).getRange('C2:C').getValues()
    .map(r => String(r[0]).trim())
    .filter(Boolean);

  const batches = {};
  ss.getSheetByName(SHEET_BATCH).getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0] || !(r[3] instanceof Date)) return;
    const id = String(r[0]);
    (batches[id] = batches[id] || []).push({
      batch: String(r[2]),
      expiry: Utilities.formatDate(r[3], 'Asia/Taipei', 'yyyy-MM-dd'),
      remaining: Number(r[4]) || 0,
    });
  });

  return { items, staff, batches };
}

/** 寫入一筆異動 */
function submitEntry(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const { items, batches } = getFormData();
    const item = items.find(i => i.id === d.id);
    const qty = Number(d.qty);

    if (!TYPES.includes(d.type)) throw new Error('請選擇異動類型');
    if (!item) throw new Error('找不到這個材料，請重新整理頁面');
    if (!d.staff) throw new Error('請選擇經手人');
    if (!Number.isFinite(qty) || qty === 0) throw new Error('請輸入數量');
    if (d.type !== '盤點調整' && qty < 0) throw new Error('數量請填正數');
    if (!d.batch) throw new Error('請填寫或選擇批號');

    let expiry = '';
    if (d.type === '進貨') {
      if (!d.expiry) throw new Error('進貨請填寫效期');
      expiry = new Date(d.expiry + 'T00:00:00');
    } else {
      const b = (batches[d.id] || []).find(x => x.batch === d.batch);
      if (!b) throw new Error('這個批號目前沒有庫存');
      if ((d.type === '領用' || d.type === '報廢') && qty > b.remaining) {
        throw new Error('批號 ' + d.batch + ' 只剩 ' + b.remaining + ' ' + item.unit);
      }
    }

    const date = d.date ? new Date(d.date + 'T00:00:00') : new Date();
    ss.getSheetByName(SHEET_LOG).appendRow([
      new Date(), date, d.type, item.id, item.name, d.batch, expiry, qty, d.staff, d.note || '',
    ]);
    SpreadsheetApp.flush();

    const after = getFormData().items.find(i => i.id === d.id);
    return '已登記：' + d.type + ' ' + item.name + ' ' + qty + ' ' + item.unit +
      '｜目前庫存 ' + after.stock + ' ' + item.unit;
  } finally {
    lock.releaseLock();
  }
}
