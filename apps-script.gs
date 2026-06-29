const SPREADSHEET_ID = '1jNneUVnANePLWLBf40EyNyldeIDMHhYQt6z_h4Zq5OY';
const ORDER_SHEET    = '訂單記錄';
const ACCOUNT_SHEET  = '帳號管理';
const PRODUCT_SHEET  = '商品管理';
const NOTIFY_EMAIL   = 'samuel040789@gmail.com';

function initSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let orderSheet = ss.getSheetByName(ORDER_SHEET);
  if (!orderSheet) {
    orderSheet = ss.insertSheet(ORDER_SHEET);
    const headers = ['訂單編號','下單時間','門市名稱','聯絡人','送貨日期','品項數','品項明細','備註','狀態','最後更新'];
    orderSheet.appendRow(headers);
    orderSheet.setFrozenRows(1);
    orderSheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#2e7d32').setFontColor('#fff');
  }

  let acctSheet = ss.getSheetByName(ACCOUNT_SHEET);
  if (!acctSheet) {
    acctSheet = ss.insertSheet(ACCOUNT_SHEET);
    const headers = ['帳號','密碼','門市名稱','啟用（YES/NO）','權限'];
    acctSheet.appendRow(headers);
    acctSheet.setFrozenRows(1);
    acctSheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#1565c0').setFontColor('#fff');
    acctSheet.appendRow(['guanyin',  '1234', '觀音店', 'YES', '']);
    acctSheet.appendRow(['taichung', '5678', '台中店', 'YES', '']);
    acctSheet.appendRow(['admin',    'admin888', '總部', 'YES', '管理員']);
    acctSheet.appendRow(['delivery', 'del123',   '出貨組', 'YES', '出貨']);
    acctSheet.autoResizeColumns(1, 5);
  }
}

function initProductSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(PRODUCT_SHEET);
  if (sheet) { Logger.log('商品管理已存在，略過'); return; }

  sheet = ss.insertSheet(PRODUCT_SHEET);
  const headers = ['分類', '品名', '規格', '價格', '啟用(YES/NO)'];
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#e65100').setFontColor('#fff');

  const data = [
    ['雞肉類', '清肉丁',               '18公斤/件',  '$155/kg｜總價 $2,790',  'YES'],
    ['雞肉類', 'T7(帶)雞腿',           '18公斤/件',  '$200/kg｜總價 $3,600',  'YES'],
    ['雞肉類', 'T6骨腿',               '18公斤/件',  '$110/kg｜總價 $1,980',  'YES'],
    ['豬肉類', '松阪豬',               '6公斤/件',   '$550/kg｜總價 $3,300',  'YES'],
    ['豬肉類', '豬肉絲(醃)',            '2.5公斤/包', '$225/kg｜總價 $562.5',  'YES'],
    ['豬肉類', '辣味噌豬排',            '20片/包',    '總價 $550',             'YES'],
    ['魚類',   '鯛魚(腹)',              '10公斤/件',  '總價 $2,650',           'YES'],
    ['魚類',   '鮭魚',                 '4公斤/件',   '總價 $1,350',           'YES'],
    ['牛肉類', '巴東牛腱',              '6份/包',     '總價 $230',             'YES'],
    ['蔬菜/紅藜', '毛豆仁',            '15公斤/件',  '$125/kg｜總價 $1,875',  'YES'],
    ['蔬菜/紅藜', '紅藜',              '25公斤/件',  '$150/kg｜總價 $3,750',  'YES'],
    ['醬料/粉', '特調醋',              '瓶',         '$120/瓶',               'YES'],
    ['醬料/粉', '菇菇醬',              '瓶',         '$750/瓶',               'YES'],
    ['醬料/粉', '烤雞粉',              '包',         '$130/包',               'YES'],
    ['醬料/粉', '口水粉',              '包',         '$30/包',                'YES'],
    ['耗材/包材', '木頭便當盒',         '6000個/批',  '$6.41/個｜總價 $38,460','YES'],
    ['耗材/包材', '加厚真空袋 20*30',   '1000個/件',  '$2.06/個｜總價 $2,060', 'YES'],
    ['耗材/包材', '裁切烤紙',           '批',         '總價 $6,200',           'YES'],
    ['耗材/包材', '塑膠袋大',           '27公斤/件',  '$135｜總價 $3,645',     'YES'],
    ['耗材/包材', '塑膠袋中',           '27公斤/件',  '$135｜總價 $3,645',     'YES'],
    ['耗材/包材', '塑膠袋小',           '25公斤/件',  '$135｜總價 $3,375',     'YES'],
    ['耗材/包材', '真空幫浦油',         '瓶',         '$220/瓶',               'YES'],
    ['耗材/包材', '高溫膠帶 2m',        '捲',         '$150/捲',               'YES'],
    ['耗材/包材', '制服(2XL/XL/L/M/S)', '件',         '$300/件',               'YES'],
    ['耗材/包材', '茶葉(綠/紅/清/烏龍)', '包',         '$300/包',               'YES'],
  ];

  for (var i = 0; i < data.length; i++) {
    sheet.appendRow(data[i]);
  }
  sheet.autoResizeColumns(1, 5);
  Logger.log('商品管理建立完成，共 ' + data.length + ' 筆');
}

// ════════════════════════════════════════════════════
//  doGet
// ════════════════════════════════════════════════════
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'list';

  // ── 加盟主登入 ──
  if (action === 'login') {
    const u = e.parameter.u || '';
    const p = e.parameter.p || '';
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(ACCOUNT_SHEET);
    if (!sheet) return json({ success: false, error: '帳號表不存在' });
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (String(row[0]).trim() === u &&
          String(row[1]).trim() === p &&
          String(row[3]).trim().toUpperCase() === 'YES') {
        return json({ success: true, store: String(row[2]).trim() });
      }
    }
    return json({ success: false });
  }

  // ── 後台管理員／出貨人員登入 ──
  if (action === 'adminLogin') {
    const account  = e.parameter.account  || '';
    const password = e.parameter.password || '';
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(ACCOUNT_SHEET);
    if (!sheet) return json({ success: false, error: '找不到帳號表' });
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (String(row[0]).trim() === account &&
          String(row[1]).trim() === password &&
          String(row[3]).trim().toUpperCase() === 'YES') {
        const role = String(row[4] || '').trim(); // E欄：權限
        if (role === '管理員' || role === '出貨') {
          return json({ success: true, role: role, name: String(row[2]).trim() });
        }
        return json({ success: false, error: '此帳號無後台權限' });
      }
    }
    return json({ success: false, error: '帳號或密碼錯誤' });
  }

  // ── 取得商品（加盟主叫貨用）──
  if (action === 'getProducts') {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(PRODUCT_SHEET);
    if (!sheet) return json({ products: [] });
    const vals = sheet.getDataRange().getValues();
    const catMap = {};
    const catOrder = [];
    for (var i = 1; i < vals.length; i++) {
      var row    = vals[i];
      var cat    = String(row[0]).trim();
      var name   = String(row[1]).trim();
      var unit   = String(row[2]).trim();
      var price  = String(row[3]).trim();
      var active = String(row[4]).trim().toUpperCase();
      if (!cat || !name || active !== 'YES') continue;
      if (!catMap[cat]) { catMap[cat] = []; catOrder.push(cat); }
      catMap[cat].push({ name: name, unit: unit, price: price });
    }
    var result = catOrder.map(function(cat) {
      return { cat: cat, items: catMap[cat] };
    });
    return json({ products: result });
  }

  // ── 取得商品列表（管理員 CRUD 用）──
  if (action === 'getProductList') {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(PRODUCT_SHEET);
    if (!sheet || sheet.getLastRow() <= 1) return json({ products: [] });
    const vals = sheet.getDataRange().getValues();
    const products = vals.slice(1).map(function(r, i) {
      return {
        rowIndex: i + 2,
        category: String(r[0] || '').trim(),
        name:     String(r[1] || '').trim(),
        unit:     String(r[2] || '').trim(),
        price:    String(r[3] || '').trim(),
        active:   String(r[4] || '').trim()
      };
    }).filter(function(p) { return p.name; });
    return json({ products: products });
  }

  // ── 更新訂單狀態 ──
  if (action === 'updateStatus') {
    const rowIdx = parseInt(e.parameter.row);
    const status = e.parameter.status || '';
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(ORDER_SHEET);
    const row = sheet.getRange(rowIdx, 1, 1, 10).getValues()[0];
    const currentStatus = String(row[8]).trim();
    if (status === '已取消' && currentStatus !== '待確認') {
      return json({ success: false, error: '此訂單已確認，無法取消，請聯絡總部' });
    }
    sheet.getRange(rowIdx, 9).setValue(status);
    sheet.getRange(rowIdx, 10).setValue(fmt(new Date()));
    if (status === '已取消') {
      try {
        MailApp.sendEmail({
          to: NOTIFY_EMAIL,
          subject: '【樂健屋取消訂單】' + row[2] + ' - ' + row[0],
          body: '訂單取消通知\n\n訂單編號：' + row[0] + '\n門市：' + row[2] +
                '\n品項：' + row[6] + '\n\n此訂單已由分店取消。'
        });
      } catch(e) {}
    }
    return json({ success: true });
  }

  // ── 讀取所有訂單 ──
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ORDER_SHEET);
  if (!sheet || sheet.getLastRow() <= 1) return json({ orders: [] });
  const vals = sheet.getDataRange().getValues();
  const orders = vals.slice(1).map(function(r, i) {
    return {
      rowIndex:    i + 2,
      orderId:     r[0], time:        r[1].toString(),
      store:       r[2], contact:     r[3],
      deliverDate: r[4], itemCount:   r[5],
      items:       r[6], note:        r[7],
      status:      r[8], lastUpdate:  r[9].toString()
    };
  }).reverse();
  return json({ orders: orders });
}

// ════════════════════════════════════════════════════
//  doPost
// ════════════════════════════════════════════════════
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ── 新增商品 ──
    if (data.action === 'addProduct') {
      const s = ss.getSheetByName(PRODUCT_SHEET);
      s.appendRow([data.category || '', data.name, data.unit, data.price, 'YES']);
      return json({ success: true });
    }

    // ── 修改商品 ──
    if (data.action === 'updateProduct') {
      const s = ss.getSheetByName(PRODUCT_SHEET);
      const r = parseInt(data.row);
      s.getRange(r, 1).setValue(data.category || '');
      s.getRange(r, 2).setValue(data.name);
      s.getRange(r, 3).setValue(data.unit);
      s.getRange(r, 4).setValue(data.price);
      return json({ success: true });
    }

    // ── 刪除商品 ──
    if (data.action === 'deleteProduct') {
      ss.getSheetByName(PRODUCT_SHEET).deleteRow(parseInt(data.row));
      return json({ success: true });
    }

    // ── 新增訂單 ──
    let sheet = ss.getSheetByName(ORDER_SHEET);
    if (!sheet) { initSheets(); sheet = ss.getSheetByName(ORDER_SHEET); }
    const orderId    = 'ORD-' + new Date().getTime();
    const itemDetail = data.items.map(function(i) {
      return i.name + ' x' + i.qty + i.unit;
    }).join(' | ');
    const now = new Date();
    sheet.appendRow([
      orderId, fmt(now), data.store, data.contact || '',
      data.deliverDate || '', data.items.length + ' 項',
      itemDetail, data.note || '', '待確認', fmt(now)
    ]);
    sheet.autoResizeColumns(1, 10);
    try {
      const itemLines = data.items.map(function(i, idx) {
        return (idx + 1) + '. ' + i.name + ' x ' + i.qty + ' ' + i.unit;
      }).join('\n');
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: '【樂健屋新訂單】' + data.store + ' (' + data.items.length + '項)',
        body: '=== 樂健屋新訂單通知 ===\n\n訂單編號：' + orderId +
              '\n門市：' + data.store +
              '\n聯絡人：' + (data.contact || '未填') +
              '\n預計寄出：' + (data.deliverDate || '未指定') +
              '\n下單時間：' + fmt(now) +
              '\n\n叫貨品項：\n' + itemLines +
              (data.note ? '\n\n備註：' + data.note : '') +
              '\n\n請至後台確認訂單。'
      });
    } catch (mailErr) {}
    return json({ success: true, orderId: orderId });

  } catch (err) {
    return json({ success: false, error: err.toString() });
  }
}

// ════════════════════════════════════════════════════
//  工具函數
// ════════════════════════════════════════════════════
function fmt(d) {
  return Utilities.formatDate(d, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
}
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
