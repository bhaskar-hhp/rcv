/* Jio Order Manager — Google Apps Script proxy
 * Deploy as Web App → use the URL in rcv/index.html
 * Config stored in "Config" sheet tab (key | value rows)
 */

const JIO_BASE = 'https://onejio-all.jioconnect.com';
const CONFIG_SHEET = 'Config';

function getConfig(key) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
    if (!sheet) return '';
    const data = sheet.getDataRange().getValues();
    for (const row of data) {
      if (row[0] === key) return String(row[1] || '');
    }
  } catch (e) {}
  return '';
}

function setConfig(key, value) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG_SHEET);
      sheet.appendRow(['key', 'value']);
      sheet.setFrozenRows(1);
    }
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        return;
      }
    }
    sheet.appendRow([key, value]);
  } catch (e) {}
}

function getProps() {
  return {
    cookies: getConfig('JIO_COOKIES') || '',
    hmacKey: getConfig('JIO_HMAC_KEY') || 'T1MtRFNNLUtFWSRAKiUmIQ==',
    userId: getConfig('JIO_USER_ID') || '',
    userName: getConfig('JIO_USERNAME') || '',
    addFundMult: getConfig('ADD_FUND_MULT') || '',
    otherMult: getConfig('OTHER_MULT') || '',
    balance: getConfig('BALANCE') || '',
  };
}

// ── HMAC Signature ────────────────────────────────────────────────────
function computeSignature(fullName, userId, path, bodyJson) {
  const raw = fullName + ':' + userId + ':' + path + ':' + (bodyJson || '{}');
  const key = Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_256,
    raw,
    getProps().hmacKey
  );
  return key.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// ── Jio API call ──────────────────────────────────────────────────────
function jioApi(method, path, body, userName, userId) {
  const props = getProps();
  const bodyStr = body ? JSON.stringify(body) : null;
  const name = userName || props.userName;
  const id = userId || props.userId;

  const sig = computeSignature(name, id, path, bodyStr || '{}');

  const headers = {
    'Accept': 'application/json, text/plain, */*',
    'User-Agent': 'Mozilla/5.0',
    'Referer': JIO_BASE + '/dsm-orders/',
    'UserDetails': id,
    'X-Signature': sig,
    'Content-Type': 'application/json',
    'Cookie': props.cookies,
  };

  const params = {
    method: method,
    headers: headers,
    muteHttpExceptions: true,
  };

  if (bodyStr) params.payload = bodyStr;

  const resp = UrlFetchApp.fetch(JIO_BASE + path, params);
  const text = resp.getContentText();

  try {
    return { status: resp.getResponseCode(), data: JSON.parse(text) };
  } catch (e) {
    return { status: resp.getResponseCode(), data: text };
  }
}

// ── Get user info from Jio (no X-Signature needed) ────────────────────
function getUserInfo() {
  const props = getProps();
  const headers = {
    'Accept': 'application/json, text/plain, */*',
    'Cookie': props.cookies,
  };
  try {
    const resp = UrlFetchApp.fetch(JIO_BASE + '/api/dsm-authorize/get-logged-user', {
      method: 'GET', headers: headers, muteHttpExceptions: true,
    });
    if (resp.getResponseCode() === 200) {
      let data = JSON.parse(resp.getContentText());
      if (Array.isArray(data)) data = data[0];
      return data;
    }
  } catch (e) {}
  return null;
}

// ── Web App Handlers ─────────────────────────────────────────────────

function doGet(e) {
  // Simple GET: show status
  const html = HtmlService.createHtmlOutput(`
    <h2>Jio Order Manager — Apps Script Proxy</h2>
    <p>Status: Running</p>
    <p>Use POST with JSON body to call actions.</p>
    <hr>
    <h3>Actions:</h3>
    <ul>
      <li><code>setup</code> — Set JIO_COOKIES, JIO_USER_ID, JIO_USERNAME in Script Properties</li>
      <li><code>saveOrder</code> — Create push order <code>{customerNum, amount}</code></li>
      <li><code>getOrders</code> — Fetch orders <code>{type: "primary"|"secondary"}</code></li>
      <li><code>getUser</code> — Get logged-in user info</li>
    </ul>
  `);
  return html.setTitle('Jio Proxy');
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    let result;
    switch (action) {

      case 'setup':
        if (data.cookies) setConfig('JIO_COOKIES', data.cookies);
        if (data.userId) setConfig('JIO_USER_ID', data.userId);
        if (data.userName) setConfig('JIO_USERNAME', data.userName);
        if (data.addFundMult) setConfig('ADD_FUND_MULT', data.addFundMult);
        if (data.otherMult) setConfig('OTHER_MULT', data.otherMult);
        if (data.balance) setConfig('BALANCE', data.balance);
        result = { success: true, message: 'Config saved' };
        break;

      case 'getConfig':
        const cfg = getProps();
        result = { success: true, userName: cfg.userName, userId: cfg.userId, hasCookies: !!cfg.cookies, addFundMult: cfg.addFundMult, otherMult: cfg.otherMult, balance: cfg.balance };
        break;

      case 'getUser':
        const user = getUserInfo();
        result = { success: !!user, data: user };
        break;

      case 'saveOrder':
        result = createPushOrder(data.customerNum, data.amount);
        break;

      case 'getOrders':
        result = fetchOrders(data.type || 'secondary');
        break;

      case 'getGstCalc':
        result = getGstCalcData();
        break;

      case 'getRetailers':
        result = getRetailersList();
        break;

      case 'addCreditRow':
        result = addGstCalcRow(data);
        break;

      case 'getClosingBalance':
        result = getClosingBalance();
        break;

      case 'completeRow':
        result = completeRowFromSheet(data.rowIndex);
        break;

      case 'pushRow':
        result = pushRowFromSheet(data.rowIndex);
        break;

      case 'updateOrderId':
        result = updateOrderIdInSheet(data.rowIndex, data.orderId);
        break;

      case 'saveTallyData':
        result = saveTallyDataToSheet(data.ledgers);
        break;

      case 'updateRemark':
        result = updateRemarkInSheet(data.rowIndex, data.remark);
        break;

      case 'getTallyData':
        result = getTallyDataFromSheet();
        break;

      case 'deleteRow':
        result = deleteRowFromSheet(data);
        break;

      case 'getPrimary':
        result = getPrimaryValue();
        break;

      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: e.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Save Order ────────────────────────────────────────────────────────
function createPushOrder(customerNum, amount) {
  const userInfo = getUserInfo();
  if (!userInfo) return { success: false, error: 'Auth failed — cookies may be expired' };

  const props = getProps();
  const startup = (userInfo.StartUp || [{}])[0];
  const fullName = startup.fullName || props.userName;
  const userId = startup.id || props.userId;
  const priceGrp = userInfo.PriceGrp || 'DI';

  const path = '/api/dsm-orders/create-push-etopup-order';
  const body = {
    userType: 'ZD',
    payload: {
      DocType: 'ZETP',
      PriceGroup: priceGrp,
      ETPPUSHHDRITMNAV: [
        { CustomerNum: String(customerNum), OrderAmt: String(Number(amount).toFixed(2)) }
      ],
    },
  };

  const result = jioApi('POST', path, body, fullName, userId);

  if (result.status === 200) {
    const data = result.data;
    const messages = (Array.isArray(data) && data[0]?.messages) || [];
    const orderMsg = messages.find(m => m.message?.includes('Created Successfully'));
    const orderId = orderMsg ? orderMsg.message.match(/Order\s+(\d+)/)?.[1] : null;

    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Jio_secondary');
      if (sheet) {
  const today = Utilities.formatDate(new Date(), 'IST', 'dd-MM-yyyy');
        sheet.appendRow([today, orderId || '', customerNum, amount, 'Pending']);
      }
    } catch (e) {}

    return { success: true, orderId, messages };
  } else {
    return { success: false, error: 'HTTP ' + result.status, data: result.data };
  }
}

// ── GST Calc Table (Save Order Page) ─────────────────────────────────

function loadRdsMaster() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('rds_master');
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] && row[1]) {
      const key = String(row[0]).replace(/\(RCV\)/gi, '').replace(/[,\-]/g, ' ').trim().toLowerCase();
      map[key] = String(row[1]).trim();
    }
  }
  return map;
}

function getRetailersList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('rds_master');
  if (!sheet) return { success: false, error: 'rds_master not found' };
  const data = sheet.getDataRange().getValues();
  const retailers = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      const name = String(row[0]).replace(/\s+$/, '');
      const customerNum = String(row[1] || '').trim();
      retailers.push({ name, customerNum });
    }
  }
  return { success: true, count: retailers.length, data: retailers };
}

function matchCustomer(rdsName, masterMap) {
  if (!rdsName) return '';
  const clean = String(rdsName).replace(/\(RCV\)/gi, '').replace(/[,\-]/g, ' ').trim().toLowerCase();
  // Direct match
  if (masterMap[clean]) return masterMap[clean];
  // Partial match: check each key
  for (const key in masterMap) {
    if (clean.includes(key) || key.includes(clean)) return masterMap[key];
  }
  return '';
}

function getGstCalcData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const gst = ss.getSheetByName('gst_calc');
    if (!gst) return { success: false, error: 'gst_calc sheet not found' };

    const masterMap = loadRdsMaster();
    const rows = gst.getDataRange().getValues();
    const result = [];

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      const partner = String(r[2] || '').replace(/\s+$/, '');
      const customer = matchCustomer(partner, masterMap) || String(r[3] || '');
      result.push({
        rowIndex: i + 1,
        date: r[0] instanceof Date ? Utilities.formatDate(r[0], 'IST', 'dd-MM-yyyy') : String(r[0] || ''),
        orderId: String(r[1] || '').replace(/[^\d]/g, ''),
        partner: partner,
        customerNum: customer,
        basic: String(r[4] || ''),
        amount: String(r[5] || ''),
        status: String(r[8] || 'Pending'),
        closingBal: String(r[6] || ''),
        remark: String(r[7] || ''),
        hasOrderId: !!String(r[1] || '').replace(/[^\d]/g, ''),
      });
    }
    return { success: true, count: result.length, data: result };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function completeRowFromSheet(rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const gst = ss.getSheetByName('gst_calc');
    if (!gst) return { success: false, error: 'gst_calc not found' };
    gst.getRange(rowIndex, 9).setValue('Completed');
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function pushRowFromSheet(rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const gst = ss.getSheetByName('gst_calc');
    if (!gst) return { success: false, error: 'gst_calc not found' };

    const row = gst.getRange(rowIndex, 1, 1, 10).getValues()[0];
    const partner = String(row[2] || '').replace(/\s+$/, '');
    const masterMap = loadRdsMaster();
    const customerNum = matchCustomer(partner, masterMap) || String(row[3] || '');
    const amount = String(row[5] || '').replace(/[^\d.]/g, '');

    if (!customerNum) return { success: false, error: 'Customer number not found for ' + partner };
    if (!amount || parseFloat(amount) <= 0) return { success: false, error: 'Invalid amount: ' + row[5] };

    // Create push order via Jio API
    const orderResult = createPushOrder(customerNum, amount);
    if (orderResult.success && orderResult.orderId) {
      // Update sheet: Order ID (B), Status (I)
      gst.getRange(rowIndex, 2).setValue(orderResult.orderId);
      gst.getRange(rowIndex, 9).setValue('Completed');
      return { success: true, orderId: orderResult.orderId };
    }
    return orderResult;
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ── Fetch Orders ──────────────────────────────────────────────────────
function fetchOrders(type) {
  const userInfo = getUserInfo();
  if (!userInfo) return { success: false, error: 'Auth failed' };

  const props = getProps();
  const startup = (userInfo.StartUp || [{}])[0];
  const fullName = startup.fullName || props.userName;
  const userId = startup.id || props.userId;
  const custNum = userInfo.CustomerNum || '660002825';
  const pushInd = type === 'primary' ? 'L' : 'S';

  const today = new Date();
  const twoDaysAgo = new Date(today.getTime() - 2 * 86400000);
  const fmt = d => Utilities.formatDate(d, 'IST', "yyyy-MM-dd'T'HH:mm:ss");

  const params = {
    soldToParty: "'" + custNum + "'",
    statusType: "'A'",
    shipToParty: type === 'secondary' ? "'" + custNum + "'" : "''",
    userInd: "''",
    returnInd: "''",
    fromDate: "'" + fmt(twoDaysAgo) + "'",
    toDate: "'" + fmt(today) + "'",
    pushOrderInd: "'" + pushInd + "'",
    userType: 'ZD',
  };

  const qs = Object.entries(params).map(([k, v]) => k + '=' + v).join('&');
  const path = '/api/dsm-orders/e-topup-orders?' + qs;

  const result = jioApi('GET', path, null, fullName, userId);
  if (result.status === 200) {
    return { success: true, count: (result.data || []).length, data: result.data };
  }
  return { success: false, error: 'HTTP ' + result.status, data: result.data };
}

// ── Add Credit Row to gst_calc ────────────────────────────────────
function getLastClosingBalance() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('gst_calc');
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const val = sheet.getRange(lastRow, 7).getValue();
  return parseFloat(val) || 0;
}

function recalculateBalances() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('gst_calc');
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const balances = [];
  let balance = 0;

  for (const row of data) {
    const amount = parseFloat(row[5]) || 0;
    const partner = String(row[2] || '').trim().toUpperCase();
    balance = partner === 'SELF' ? balance + amount : balance - amount;
    balances.push([balance]);
  }

  sheet.getRange(2, 7, balances.length, 1).setValues(balances);
  setConfig('BALANCE', String(balance));
  return balance;
}

function addGstCalcRow(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('gst_calc');
  if (!sheet) return { success: false, error: 'gst_calc not found' };

  const prevBalance = getLastClosingBalance();
  const amount = parseFloat(data.amount) || 0;
  const newBalance = data.isAddFund
    ? prevBalance + amount
    : prevBalance - amount;

  const today = Utilities.formatDate(new Date(), 'IST', 'dd-MMM-yyyy');
  sheet.appendRow([
    today,            // Date
    data.orderId || '', // Order ID
    data.partner,     // Partner
    data.customerNum, // Customer #
    data.basic,       // Basic
    data.amount,      // Amount
    newBalance,       // Closing Bal
    '',               // Remark
    'In Credit',      // Status
  ]);

  setConfig('BALANCE', String(newBalance));
  return { success: true, closingBal: newBalance };
}

function getClosingBalance() {
  return { success: true, closingBal: getLastClosingBalance() };
}

function getPrimaryValue() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('rds_master');
  if (!sheet) return { success: false, error: 'rds_master not found' };
  const val = sheet.getRange('F1').getValue();
  return { success: true, primary: String(val) };
}

function updateOrderIdInSheet(rowIndex, orderId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('gst_calc');
  if (!sheet) return { success: false, error: 'gst_calc not found' };
  sheet.getRange(rowIndex, 2).setValue(orderId);
  return { success: true };
}

function saveTallyDataToSheet(ledgers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetGid = 638048736;
  let sheet = ss.getSheets().filter(s => s.getSheetId() === targetGid)[0];
  if (!sheet) return { success: false, error: 'Sheet with gid=638048736 not found' };

  sheet.clearContents();
  const now = Utilities.formatDate(new Date(), 'IST', 'dd-MM-yyyy HH:mm');
  const rows = ledgers.map(l => [l.name, l.openingBalance, l.closingBalance, now]);
  rows.unshift(['Ledger Name', 'Opening Balance', 'Closing Balance', 'Timestamp']);
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  return { success: true, count: ledgers.length };
}

function getTallyDataFromSheet() {
  const targetGid = 638048736;
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(s => s.getSheetId() === targetGid)[0];
  if (!sheet) return { success: false, error: 'Tally sheet not found' };
  const rows = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    const name = String(rows[i][0] || '').trim();
    const closing = String(rows[i][2] || '').trim();
    if (name) map[name] = closing;
  }
  return { success: true, data: map };
}

function updateRemarkInSheet(rowIndex, remark) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('gst_calc');
  if (!sheet) return { success: false, error: 'gst_calc not found' };
  sheet.getRange(rowIndex, 8).setValue(remark);
  return { success: true };
}

function deleteRowFromSheet(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('gst_calc');
  if (!sheet) return { success: false, error: 'gst_calc not found' };

  sheet.deleteRow(data.rowIndex);
  const closingBal = recalculateBalances();
  return { success: true, closingBal };
}
