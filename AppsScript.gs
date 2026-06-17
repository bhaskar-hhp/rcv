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

      case 'getJioBalance':
        result = getJioBalance();
        break;
      case 'getCustomerCredit':
        result = getCustomerCredit();
        break;

      case 'approveOrder':
        result = approveOrderOnJio(data.orderNum, data.amount);
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

      case 'getRecoData':
        result = getRecoDataFromSheet();
        break;

      case 'deleteDuplicateVoucher':
        result = deleteDuplicateVoucherFromSheet();
        break;

      case 'deleteRow':
        result = deleteRowFromSheet(data);
        break;

      case 'getPrimary':
        result = getPrimaryValue();
        break;

      case 'getRdsData':
        result = getRdsMasterData();
        break;

      case 'addRdsRow':
        result = addRdsRowToSheet();
        break;

      case 'savePrimaryOrder':
        result = createPrimaryOrder(data.customerNum, data.basicAmount, data.amount);
        break;

      case 'deleteRdsRow':
        result = deleteRdsRowFromSheet(data.rowIndex);
        break;

      case 'getDeviceProducts':
        result = fetchDeviceProducts();
        break;
      case 'getChannelPartners':
        result = fetchChannelPartners();
        break;
      case 'getARDpartners':
        result = fetchARDpartners();
        break;
      case 'createDeviceOrder':
        result = createDeviceOrderApi(data);
        break;
      case 'approveDeviceOrder':
        result = approveDeviceOrderApi(data.orderNum);
        break;
      case 'getPendingDeviceOrders':
        result = fetchMyDeviceOrdersList(data.from, data.to);
        break;
      case 'saveDeviceOrder':
        result = saveDeviceOrderToSheet(data);
        break;
      case 'getSavedDeviceOrders':
        result = fetchSavedDeviceOrders();
        break;
      case 'fetchDeviceSheetData':
        result = fetchDeviceSheetData(data);
        break;
      case 'syncDeviceSheetStatus':
        result = syncDeviceSheetStatus(data);
        break;
      case 'updateDeviceOrderStatus':
        result = updateDeviceOrderStatus(data.orderId, data.status);
        break;
      case 'createSimOrder':
        result = createSimOrderApi(data);
        break;
      case 'approveSimOrder':
        result = approveSimOrderApi(data.orderNum);
        break;
      case 'saveSimOrder':
        result = saveSimOrderToSheet(data);
        break;
      case 'getSavedSimOrders':
        result = fetchSavedSimOrders();
        break;
      case 'updateSimOrderStatus':
        result = updateSimOrderStatus(data.orderId, data.status);
        break;
      case 'getPendingSimOrders':
        result = fetchMySimOrdersList(data.from, data.to);
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

// ── Create Primary Order (SELF) ─────────────────────────────────────────
function createPrimaryOrder(customerNum, basicAmount, amount) {
  const userInfo = getUserInfo();
  if (!userInfo) return { success: false, error: 'Auth failed — cookies may be expired' };

  const props = getProps();
  const startup = (userInfo.StartUp || [{}])[0];
  const fullName = startup.fullName || props.userName;
  const userId = startup.id || props.userId;

  const article = '491216866';

  // Step 1: quantity-header-set
  const qtyPath = '/api/dsm-orders/quantity-header-set';
  const qtyBody = {
    userType: 'ZD',
    payload: {
      BlockInd: '', ContactPerson: '', ShipToParty: '660002825',
      SoldToParty: '660002825', DocType: 'ZETP',
      DraftOrderNum: '', UserId: '', ParentPatner: '',
      QTYDETHEADERNAV: [{ ArticleNum: article, TargetQty: String(basicAmount), UoM: 'EA' }],
    },
  };
  const qtyResult = jioApi('POST', qtyPath, qtyBody, fullName, userId);
  if (qtyResult.status !== 200) {
    return { success: false, error: 'quantity-header-set failed: HTTP ' + qtyResult.status };
  }

  // Step 2: order-create-set
  const orderPath = '/api/dsm-orders/order-create-set';
  const orderBody = {
    userType: 'ZD',
    payload: {
      BlockInd: '', ContactPerson: '', ShipToParty: '660002825',
      SoldToParty: '660002825', DocType: 'ZETP',
      DraftOrderNum: '', UserId: '', ParentPatner: '',
      ETPORDERCREATENAV: [{ ArticleNum: article, TargetQty: String(amount), UoM: 'EA' }],
    },
  };
  const orderResult = jioApi('POST', orderPath, orderBody, fullName, userId);

  const item = Array.isArray(orderResult.data) ? orderResult.data[0] : orderResult.data;
  const bizStatus = item?.statusCode || orderResult.status;
  if (bizStatus === 200 || bizStatus === 201) {
    const sapMsg = item?.headers?.['sap-message'];
    let orderNum = '';
    if (sapMsg) {
      try {
        const parsed = JSON.parse(sapMsg);
        const match = (parsed.message || '').match(/Order\s+(\d+)/);
        if (match) orderNum = match[1];
      } catch (e) {}
    }
    if (!orderNum) orderNum = item?.body?.d?.OrderNum || item?.body?.d?.orderNum || '';
    return { success: true, orderId: String(orderNum) };
  }
  return { success: false, error: 'order-create-set failed: HTTP ' + bizStatus, data: orderResult.data };
}

// ── Approve Order ─────────────────────────────────────────────────────
function approveOrderOnJio(orderNum, amount) {
  const userInfo = getUserInfo();
  if (!userInfo) return { success: false, error: 'Auth failed' };
  const props = getProps();
  const startup = (userInfo.StartUp || [{}])[0];
  const fullName = startup.fullName || props.userName;
  const userId = startup.id || props.userId;
  const path = '/api/dsm-orders/order-rejection-set?userType=ZD';
  const body = { GrantedValue: String(Number(amount).toFixed(2)), OrderNum: String(orderNum), ReasonRej: '', RDCustNum: '', ZDCustNum: '660002825' };
  const result = jioApi('PUT', path, body, fullName, userId);
  if (result.status === 200) {
    return { success: true };
  }
  return { success: false, error: 'HTTP ' + result.status, data: result.data };
}

// ── Get Jio Balance ───────────────────────────────────────────────────
function getJioBalance() {
  const userInfo = getUserInfo();
  if (!userInfo) return { success: false, error: 'Auth failed' };
  const props = getProps();
  const startup = (userInfo.StartUp || [{}])[0];
  const fullName = startup.fullName || props.userName;
  const userId = startup.id || props.userId;
  const path = "/api/dsm-orders/details-of-balance?userType=ZD&customerNumber='660002825'";
  const result = jioApi('GET', path, null, fullName, userId);
  if (result.status === 200) {
    const arr = Array.isArray(result.data) ? result.data : (result.data?.d?.results || []);
    const amt = arr[0]?.AvailableAmt || arr[0]?.availableAmt || '0';
    const credit = arr[0]?.AvlCreditLimit || arr[0]?.avlCreditLimit || '';
    return { success: true, jioBalance: parseFloat(amt), avlCreditLimit: String(credit) };
  }
  return { success: false, error: 'HTTP ' + result.status };
}

function getCustomerCredit() {
  const userInfo = getUserInfo();
  if (!userInfo) return { success: false, error: 'Auth failed' };
  const props = getProps();
  const startup = (userInfo.StartUp || [{}])[0];
  const fullName = startup.fullName || props.userName;
  const userId = startup.id || props.userId;
  const custNum = userInfo.CustomerNum || '660002825';
  const path = "/api/dsm-orders/call-customer-credit-set?userType=ZD&customerNumber='" + custNum + "'";
  const result = jioApi('GET', path, null, fullName, userId);
  if (result.status === 200) {
    const arr = Array.isArray(result.data) ? result.data : [];
    const connectivity = arr.find(r => r.ControlArea === 'RREL');
    const credit = connectivity?.AvlCreditLimit || '';
    return { success: true, avlCreditLimit: String(credit), all: arr };
  }
  return { success: false, error: 'HTTP ' + result.status };
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
        date: r[0] && typeof r[0] === 'object' && typeof r[0].getMonth === 'function' ? Utilities.formatDate(r[0], 'IST', 'dd-MM-yyyy') : String(r[0] || ''),
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
      // Update sheet: Order ID (B)
      gst.getRange(rowIndex, 2).setValue(orderResult.orderId);
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

  const prevBalance = parseFloat(getConfig('BALANCE')) || getLastClosingBalance();
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
    data.remark || '', // Remark
    'In Credit',      // Status
  ]);

  setConfig('BALANCE', String(newBalance));
  return { success: true, closingBal: newBalance };
}

function getClosingBalance() {
  return { success: true, closingBal: parseFloat(getConfig('BALANCE')) || getLastClosingBalance() };
}

function getPrimaryValue() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('rds_master');
  if (!sheet) return { success: false, error: 'rds_master not found' };
  const val = sheet.getRange('F1').getValue();
  return { success: true, primary: String(val) };
}

function getRdsMasterData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('rds_master');
  if (!sheet) return { success: false, error: 'rds_master not found' };
  const rows = sheet.getDataRange().getValues();
  const data = [];
  for (let i = 1; i < rows.length; i++) {
    data.push({ rowIndex: i + 1, a: String(rows[i][0] || ''), b: String(rows[i][1] || ''), c: String(rows[i][2] || '') });
  }
  return { success: true, data };
}

function addRdsRowToSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('rds_master');
  if (!sheet) return { success: false, error: 'rds_master not found' };
  const lastRow = sheet.getLastRow();
  let formula = '';
  if (lastRow >= 2) {
    formula = sheet.getRange(lastRow, 3).getFormula();
  }
  sheet.appendRow(['', '', formula || '']);
  return { success: true };
}

function deleteRdsRowFromSheet(rowIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('rds_master');
  if (!sheet) return { success: false, error: 'rds_master not found' };
  sheet.deleteRow(rowIndex);
  return { success: true };
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

function getRecoDataFromSheet() {
  const targetGid = 605373015;
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(s => s.getSheetId() === targetGid)[0];
  if (!sheet) return { success: false, error: 'Reco sheet not found' };
  const rows = sheet.getDataRange().getValues();
  const groups = {};
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const rawDate = r[0];
    let date = '';
    if (rawDate && typeof rawDate === 'object' && typeof rawDate.getMonth === 'function') {
      date = Utilities.formatDate(rawDate, 'IST', 'dd-MM-yyyy');
    } else {
      const s = String(rawDate || '').trim();
      const d = new Date(s);
      date = (!isNaN(d.getTime())) ? Utilities.formatDate(d, 'IST', 'dd-MM-yyyy') : s;
    }
    const partner = String(r[6] || '').trim();
    const drAmt = Math.abs(parseFloat(String(r[4] || '0').replace(/,/g, ''))) || 0;
    const crAmt = Math.abs(parseFloat(String(r[5] || '0').replace(/,/g, ''))) || 0;
    if (!partner && !date) continue;
    const key = date + '|' + partner;
    if (!groups[key]) groups[key] = { date, partner, crAmt: 0, drAmt: 0 };
    groups[key].crAmt += crAmt;
    groups[key].drAmt += drAmt;
  }
  return { success: true, data: Object.values(groups) };
}

function deleteRowFromSheet(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('gst_calc');
  if (!sheet) return { success: false, error: 'gst_calc not found' };

  sheet.deleteRow(data.rowIndex);

  // Sync balance from Jio
  const userInfo = getUserInfo();
  if (userInfo) {
    const props = getProps();
    const startup = (userInfo.StartUp || [{}])[0];
    const fullName = startup.fullName || props.userName;
    const userId = startup.id || props.userId;
    const path = "/api/dsm-orders/details-of-balance?userType=ZD&customerNumber='660002825'";
    const jioResult = jioApi('GET', path, null, fullName, userId);
    if (jioResult.status === 200) {
      const arr = Array.isArray(jioResult.data) ? jioResult.data : (jioResult.data?.d?.results || []);
      const amt = arr[0]?.AvailableAmt || arr[0]?.availableAmt || '0';
      const jioBal = parseFloat(amt);
      setConfig('BALANCE', String(jioBal));
      return { success: true, jioBalance: jioBal };
    }
  }

  const closingBal = recalculateBalances();
  return { success: true, closingBal };
}

function deleteDuplicateVoucherFromSheet() {
  const targetGid = 605373015;
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(s => s.getSheetId() === targetGid)[0];
  if (!sheet) return { success: false, error: 'Sheet gid=605373015 not found' };

  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { success: true, deleted: 0 };

  // Group by VoucherNo (col D, index 3)
  const groups = {};
  for (let i = 1; i < rows.length; i++) {
    const voucher = String(rows[i][3] || '').trim();
    if (!voucher) continue;
    if (!groups[voucher]) groups[voucher] = [];
    groups[voucher].push({ row: i + 1, data: rows[i] });
  }

  const details = [];
  const toDelete = [];

  for (const voucher in groups) {
    const items = groups[voucher];
    if (items.length < 2) continue;

    // Compare all columns — keep first, delete rest if identical
    const first = items[0].data;
    const dupRows = [];
    for (let j = 1; j < items.length; j++) {
      const cur = items[j].data;
      let identical = true;
      for (let c = 0; c < cur.length; c++) {
        if (String(cur[c] || '') !== String(first[c] || '')) { identical = false; break; }
      }
      if (identical) dupRows.push(items[j].row);
    }
    if (dupRows.length) {
      toDelete.push(...dupRows);
      details.push({ voucher, rows: dupRows.length });
    }
  }

  if (!toDelete.length) return { success: true, deleted: 0 };

  // Delete from bottom to top to preserve indices
  toDelete.sort((a, b) => b - a);
  for (const rowNum of toDelete) {
    sheet.deleteRow(rowNum);
  }

  return { success: true, deleted: toDelete.length, details };
}

// ── Device Orders ──────────────────────────────────────────────────────

function fetchChannelPartners() {
  const userInfo = getUserInfo();
  if (!userInfo) return { success: false, error: 'Auth failed' };
  const props = getProps();
  const startup = (userInfo.StartUp || [{}])[0];
  const fullName = startup.fullName || props.userName;
  const userId = startup.id || props.userId;
  const path = "/api/dsm-orders/get-channel-partners?userType=ZD&customerNumber='660002825'&skip=0&top=9999";
  const result = jioApi('GET', path, null, fullName, userId);
  if (result.status === 200) {
    return { success: true, data: Array.isArray(result.data) ? result.data : [] };
  }
  const errDetail = typeof result.data === 'string' ? result.data : JSON.stringify(result.data);
  return { success: false, error: 'HTTP ' + result.status + ' — ' + errDetail, data: result.data };
}

function fetchDeviceProducts() {
  const userInfo = getUserInfo();
  if (!userInfo) return { success: false, error: 'Auth failed' };
  const props = getProps();
  const startup = (userInfo.StartUp || [{}])[0];
  const fullName = startup.fullName || props.userName;
  const userId = startup.id || props.userId;
  const custNum = userInfo.CustomerNum || '660002825';
  const path = "/api/dsm-orders/ProductListOSSet?UserInd=ZD&ProductGrp=76&SoldToParty=" + custNum + "&ProductDiv=BA&PriceGrp=DI&OrderType=ZBHA&ProcInd=L";
  const result = jioApi('GET', path, null, fullName, userId);
  if (result.status === 200) {
    return { success: true, data: Array.isArray(result.data) ? result.data : [] };
  }
  const errDetail = typeof result.data === 'string' ? result.data : JSON.stringify(result.data);
  return { success: false, error: 'HTTP ' + result.status + ' — ' + errDetail, data: result.data };
}

function fetchARDpartners() {
  const userInfo = getUserInfo();
  if (!userInfo) return { success: false, error: 'Auth failed' };
  const props = getProps();
  const startup = (userInfo.StartUp || [{}])[0];
  const fullName = startup.fullName || props.userName;
  const userId = startup.id || props.userId;
  const custNum = userInfo.CustomerNum || '660002825';
  const path = "/api/dsm-orders/get-ARD-partners?userType=ZD&customerNumber='" + custNum + "'";
  const result = jioApi('GET', path, null, fullName, userId);
  if (result.status === 200) {
    return { success: true, data: Array.isArray(result.data) ? result.data : [] };
  }
  const errDetail = typeof result.data === 'string' ? result.data : JSON.stringify(result.data);
  return { success: false, error: 'HTTP ' + result.status + ' — ' + errDetail, data: result.data };
}

function createDeviceOrderApi(data) {
  const userInfo = getUserInfo();
  if (!userInfo) return { success: false, error: 'Auth failed' };
  const props = getProps();
  const startup = (userInfo.StartUp || [{}])[0];
  const fullName = startup.fullName || props.userName;
  const userId = startup.id || props.userId;
  const custNum = userInfo.CustomerNum || '660002825';
  const body = {
    BlockInd: 'Z5',
    ContactPerson: '',
    OrderType: data.orderType,
    ParentPatner: custNum,
    UserID: custNum,
    ShipToParty: data.shipToParty,
    SoldToParty: custNum,
    DraftOrderNum: '',
    CREATEHEADNAV: [{ ArticleNum: data.articleNum, UoM: 'EA', TargetQty: String(data.qty) }],
  };
  const result = jioApi('POST', '/api/dsm-orders/post-Order-Create?userType=ZD', body, fullName, userId);
  const item = Array.isArray(result.data) ? result.data[0] : result.data;
  const bizStatus = item?.statusCode || result.status;
  if (bizStatus === 200 || bizStatus === 201) {
    const sapMsg = item?.headers?.['sap-message'];
    let orderNum = '';
    if (sapMsg) {
      try {
        const parsed = JSON.parse(sapMsg);
        const match = (parsed.message || '').match(/Order\s+(\d+)/);
        if (match) orderNum = match[1];
      } catch (e) {}
    }
    if (!orderNum) orderNum = item?.body?.d?.OrderNum || item?.body?.d?.orderNum || '';
    return { success: true, orderNum: String(orderNum), data: result.data };
  }
  return { success: false, error: 'HTTP ' + bizStatus, data: result.data };
}

function approveDeviceOrderApi(orderNum) {
  const userInfo = getUserInfo();
  if (!userInfo) return { success: false, error: 'Auth failed' };
  const props = getProps();
  const startup = (userInfo.StartUp || [{}])[0];
  const fullName = startup.fullName || props.userName;
  const userId = startup.id || props.userId;
  const body = { userType: 'ZD', OrderNum: String(orderNum), ReleaseInd: 'Z5' };
  const result = jioApi('PUT', '/api/dsm-orders/put-approval-set?userType=ZD', body, fullName, userId);
  const item = Array.isArray(result.data) ? result.data[0] : result.data;
  const bizStatus = item?.statusCode || result.status;
  if (bizStatus === 200 || bizStatus === 201) {
    return { success: true, data: result.data };
  }
  return { success: false, error: 'HTTP ' + bizStatus, data: result.data };
}

function fetchPendingDeviceOrders(from, to) {
  const userInfo = getUserInfo();
  if (!userInfo) return { success: false, error: 'Auth failed' };
  const props = getProps();
  const startup = (userInfo.StartUp || [{}])[0];
  const fullName = startup.fullName || props.userName;
  const userId = startup.id || props.userId;
  const custNum = userInfo.CustomerNum || '660002825';

  const now = new Date();
  const fromDate = from || Utilities.formatDate(new Date(now.getTime() - 2 * 86400000), 'IST', "yyyy-MM-dd'T'HH:mm:ss");
  const toDate = to || Utilities.formatDate(now, 'IST', "yyyy-MM-dd'T'HH:mm:ss");

  const path = "/api/dsm-orders/order-displayList-set?userType=ZD&statusCode=P&fromDate=" + encodeURIComponent(fromDate) + "&toDate=" + encodeURIComponent(toDate) + "&soldToParty=" + custNum + "&shipToParty=";
  const result = jioApi('GET', path, null, fullName, userId);

  if (Array.isArray(result.data)) {
    return { success: true, data: result.data };
  }

  const item = Array.isArray(result.data) ? result.data[0] : null;
  const results = item?.body?.d?.results;
  if (Array.isArray(results)) {
    return { success: true, data: results };
  }

  return { success: true, data: [] };
}

function fetchMyDeviceOrdersList(from, to) {
  const userInfo = getUserInfo();
  if (!userInfo) return { success: false, error: 'Auth failed', data: [] };
  const props = getProps();
  const startup = (userInfo.StartUp || [{}])[0];
  const fullName = startup.fullName || props.userName;
  const userId = startup.id || props.userId;
  const custNum = userInfo.CustomerNum || '660002825';

  const now = new Date();
  const fromDate = from || Utilities.formatDate(new Date(now.getTime() - 30 * 86400000), 'IST', "yyyy-MM-dd'T'HH:mm:ss");
  const toDate = to || Utilities.formatDate(now, 'IST', "yyyy-MM-dd'T'HH:mm:ss");

  const path = "/api/dsm-orders/MyOrder-List-set?UserInd=ZD&CustomerNum=" + custNum + "&DateFrom=" + encodeURIComponent(fromDate) + "&DateTo=" + encodeURIComponent(toDate) + "&StatusTyp=L&RRLInd=";
  const result = jioApi('GET', path, null, fullName, userId);

  let orders = Array.isArray(result.data) ? result.data : [];
  if (!orders.length && result.data && typeof result.data === 'object') {
    const item = Array.isArray(result.data) ? result.data[0] : result.data;
    const results = item?.body?.d?.results;
    if (Array.isArray(results)) orders = results;
  }

  const deviceOrders = orders.filter(r => r.OrdTypDesc === 'JIO Bharat' || r.OrderType === 'ZBHA');
  return { success: true, data: deviceOrders };
}

function saveDeviceOrderToSheet(data) {
  const targetGid = 320908957;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheets().filter(s => s.getSheetId() === targetGid)[0];
  if (!sheet) return { success: false, error: 'Device sheet not found' };
  const today = Utilities.formatDate(new Date(), 'IST', 'dd-MM-yyyy');
  sheet.appendRow([
    today,
    data.orderId || '',
    data.partnerNum || '',
    data.partnerName || '',
    data.articleNum || '',
    data.productName || '',
    data.qty || '',
    data.dealerPrice || '',
    data.totalAmount || '',
    data.status || 'Pending',
    '',
  ]);
  return { success: true };
}

function updateDeviceJioStatusInSheet(orderId, jioStatus) {
  const targetGid = 320908957;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheets().filter(s => s.getSheetId() === targetGid)[0];
  if (!sheet) return;
  const rows = sheet.getDataRange().getValues();
  const startRow = String(rows[0][0]).toLowerCase().includes('date') ? 1 : 0;
  for (let i = startRow; i < rows.length; i++) {
    if (String(rows[i][1] || '') === String(orderId)) {
      sheet.getRange(i + 1, 11).setValue(jioStatus || '');
      return;
    }
  }
}

function fetchSavedDeviceOrders() {
  const targetGid = 320908957;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheets().filter(s => s.getSheetId() === targetGid)[0];
  if (!sheet) return { success: false, error: 'Device sheet not found' };
  const rows = sheet.getDataRange().getValues();

  const jioStatus = {};
  try {
    const jioRes = fetchMyDeviceOrdersList('2020-01-01T00:00:00', '2030-12-31T00:00:00');
    if (jioRes.success && Array.isArray(jioRes.data)) {
      for (const o of jioRes.data) {
        if (o.OrderNum) {
          const sid = String(o.OrderNum);
          jioStatus[sid] = o.StatusDesc || '';
        }
      }
    }
  } catch (e) {}

  const data = [];
  const startRow = String(rows[0][0]).toLowerCase().includes('date') ? 1 : 0;
  for (let i = startRow; i < rows.length; i++) {
    const r = rows[i];
    let dateStr = '';
    if (r[0] && typeof r[0] === 'object' && typeof r[0].getMonth === 'function') {
      dateStr = Utilities.formatDate(r[0], 'IST', 'dd-MM-yyyy');
    } else {
      dateStr = String(r[0] || '');
    }
    const orderId = String(r[1] || '');
    const sheetStat = String(r[9] || 'Pending');
    const liveStat = jioStatus[orderId] || '';
    const colK = String(r[10] || '');
    if (liveStat && sheetStat !== 'Completely Dispatched') {
      sheet.getRange(i + 1, 10).setValue(liveStat);
      r[9] = liveStat;
    }
    const finalStat = sheetStat === 'Completely Dispatched' ? sheetStat : (liveStat || colK || sheetStat);
    data.push({
      date: dateStr,
      orderId: orderId,
      partnerNum: String(r[2] || ''),
      partnerName: String(r[3] || ''),
      articleNum: String(r[4] || ''),
      productName: String(r[5] || ''),
      qty: String(r[6] || ''),
      dealerPrice: String(r[7] || ''),
      totalAmount: String(r[8] || ''),
      status: finalStat,
      sheetStatus: liveStat ? liveStat : sheetStat,
      jioStatus: liveStat || colK || '',
    });
  }

  return { success: true, data: data };
}

function fetchDeviceSheetData(data) {
  try {
    const targetGid = 320908957;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheets().filter(s => s.getSheetId() === targetGid)[0];
    if (!sheet) return { success: false, error: 'Device sheet not found' };

    const rows = sheet.getDataRange().getValues();
    const startRow = String(rows[0][0]).toLowerCase().includes('date') ? 1 : 0;

    const filterDateFrom = data?.dateFrom || '';
    const filterDateTo = data?.dateTo || '';
    const filterPartner = (data?.partner || '').toLowerCase();
    const filterStatus = (data?.status || '').toLowerCase();

    const result = [];
    for (let i = startRow; i < rows.length; i++) {
      const r = rows[i];

      const partnerName = String(r[3] || '').toLowerCase();
      if (filterPartner && !partnerName.includes(filterPartner)) continue;
      const rowStatus = String(r[9] || 'Pending').toLowerCase();
      if (filterStatus && rowStatus !== filterStatus) continue;

      result.push({
        date: String(r[0] || ''),
        orderId: String(r[1] || ''),
        partnerNum: String(r[2] || ''),
        partnerName: String(r[3] || ''),
        articleNum: String(r[4] || ''),
        productName: String(r[5] || ''),
        qty: String(r[6] || ''),
        dealerPrice: String(r[7] || ''),
        totalAmount: String(r[8] || ''),
        status: String(r[9] || 'Pending'),
      });
    }
    return { success: true, data: result, total: result.length };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function syncDeviceSheetStatus(data) {
  const targetGid = 320908957;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheets().filter(s => s.getSheetId() === targetGid)[0];
  if (!sheet) return { success: false, error: 'Device sheet not found' };

  const tableOrders = (data?.orders || []).filter(o => o.orderId);
  const tableMap = {};
  for (const o of tableOrders) {
    tableMap[o.orderId] = o.statusDesc || '';
  }

  const rows = sheet.getDataRange().getValues();
  const startRow = String(rows[0][0]).toLowerCase().includes('date') ? 1 : 0;
  const sheetOrderIds = [];
  let updated = 0;
  for (let i = startRow; i < rows.length; i++) {
    const orderId = String(rows[i][1] || '');
    if (orderId) sheetOrderIds.push(orderId);
    const sheetStat = String(rows[i][9] || 'Pending');
    const liveStat = tableMap[orderId] || '';
    if (liveStat && sheetStat !== 'Completely Dispatched' && liveStat !== sheetStat) {
      sheet.getRange(i + 1, 10).setValue(liveStat);
      updated++;
    }
  }
  const matched = tableOrders.filter(o => sheetOrderIds.includes(o.orderId));
  return {
    success: true, updated: updated,
    totalOrders: tableOrders.length,
    sheetCount: sheetOrderIds.length,
    matchedCount: matched.length,
  };
}

function updateDeviceOrderStatus(orderId, newStatus) {
  const targetGid = 320908957;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheets().filter(s => s.getSheetId() === targetGid)[0];
  if (!sheet) return { success: false, error: 'Device sheet not found' };
  const rows = sheet.getDataRange().getValues();
  const startRow = String(rows[0][0]).toLowerCase().includes('date') ? 1 : 0;
  for (let i = startRow; i < rows.length; i++) {
    if (String(rows[i][1] || '') === String(orderId)) {
      sheet.getRange(i + 1, 10).setValue(newStatus || 'Approved');
      return { success: true };
    }
  }
  return { success: false, error: 'Order not found: ' + orderId };
}

// ── SIM Order ──────────────────────────────────────────────────────────

const SIM_GID = 1261221666;

function createSimOrderApi(data) {
  const userInfo = getUserInfo();
  if (!userInfo) return { success: false, error: 'Auth failed' };
  const props = getProps();
  const startup = (userInfo.StartUp || [{}])[0];
  const fullName = startup.fullName || props.userName;
  const userId = startup.id || props.userId;
  const custNum = userInfo.CustomerNum || '660002825';
  const body = {
    BlockInd: 'Z5',
    ContactPerson: '',
    OrderType: 'ZJTP',
    ParentPatner: custNum,
    UserID: custNum,
    ShipToParty: data.shipToParty,
    SoldToParty: custNum,
    DraftOrderNum: '',
    CREATEHEADNAV: [{ ArticleNum: '920001280', UoM: 'EA', TargetQty: String(data.qty) }],
  };
  const result = jioApi('POST', '/api/dsm-orders/post-Order-Create?userType=ZD', body, fullName, userId);
  const item = Array.isArray(result.data) ? result.data[0] : result.data;
  const bizStatus = item?.statusCode || result.status;
  if (bizStatus === 200 || bizStatus === 201) {
    const sapMsg = item?.headers?.['sap-message'];
    let orderNum = '';
    if (sapMsg) {
      try {
        const parsed = JSON.parse(sapMsg);
        const match = (parsed.message || '').match(/Order\s+(\d+)/);
        if (match) orderNum = match[1];
      } catch (e) {}
    }
    if (!orderNum) orderNum = item?.body?.d?.OrderNum || item?.body?.d?.orderNum || '';
    return { success: true, orderNum: String(orderNum), data: result.data };
  }
  return { success: false, error: 'HTTP ' + bizStatus, data: result.data };
}

function approveSimOrderApi(orderNum) {
  const userInfo = getUserInfo();
  if (!userInfo) return { success: false, error: 'Auth failed' };
  const props = getProps();
  const startup = (userInfo.StartUp || [{}])[0];
  const fullName = startup.fullName || props.userName;
  const userId = startup.id || props.userId;
  const body = { userType: 'ZD', OrderNum: String(orderNum), ReleaseInd: 'Z5' };
  const result = jioApi('PUT', '/api/dsm-orders/put-approval-set', body, fullName, userId);
  const item = Array.isArray(result.data) ? result.data[0] : result.data;
  const bizStatus = item?.statusCode || result.status;
  if (bizStatus === 200 || bizStatus === 201) {
    return { success: true, data: result.data };
  }
  return { success: false, error: 'HTTP ' + bizStatus, data: result.data };
}

function fetchMySimOrdersList(from, to) {
  const userInfo = getUserInfo();
  if (!userInfo) return { success: false, error: 'Auth failed' };
  const props = getProps();
  const startup = (userInfo.StartUp || [{}])[0];
  const fullName = startup.fullName || props.userName;
  const userId = startup.id || props.userId;
  const custNum = userInfo.CustomerNum || '660002825';

  const now = new Date();
  const fromDate = from || Utilities.formatDate(new Date(now.getTime() - 30 * 86400000), 'IST', "yyyy-MM-dd'T'HH:mm:ss");
  const toDate = to || Utilities.formatDate(now, 'IST', "yyyy-MM-dd'T'HH:mm:ss");

  const path = "/api/dsm-orders/MyOrder-List-set?UserInd=ZD&CustomerNum=" + custNum + "&DateFrom=" + encodeURIComponent(fromDate) + "&DateTo=" + encodeURIComponent(toDate) + "&StatusTyp=L&RRLInd=";
  const result = jioApi('GET', path, null, fullName, userId);

  const orders = Array.isArray(result.data) ? result.data : [];
  const simOrders = orders.filter(r => r.OrdTypDesc === 'SIM & CAF' || r.OrderType === 'ZJTP');
  return { success: true, data: simOrders };
}

function saveSimOrderToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheets().filter(s => s.getSheetId() === SIM_GID)[0];
  if (!sheet) return { success: false, error: 'SIM sheet not found' };
  const today = Utilities.formatDate(new Date(), 'IST', 'dd-MM-yyyy');
  sheet.appendRow([
    today,
    data.orderId || '',
    data.customerNum || '',
    data.customerName || '',
    data.qty || '',
    data.status || 'Pending',
  ]);
  return { success: true };
}

function fetchSavedSimOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheets().filter(s => s.getSheetId() === SIM_GID)[0];
  if (!sheet) return { success: false, error: 'SIM sheet not found' };
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { success: true, data: [] };
  const data = [];
  const startRow = String(rows[0][0]).toLowerCase().includes('date') ? 1 : 0;
  for (let i = startRow; i < rows.length; i++) {
    const r = rows[i];
    let dateStr = '';
    if (r[0] instanceof Date) {
      dateStr = Utilities.formatDate(r[0], 'IST', 'dd-MM-yyyy');
    } else {
      dateStr = String(r[0] || '');
    }
    data.push({
      date: dateStr,
      orderId: String(r[1] || ''),
      customerNum: String(r[2] || ''),
      customerName: String(r[3] || ''),
      qty: String(r[4] || ''),
      status: String(r[5] || 'Pending'),
    });
  }
  return { success: true, data: data };
}

function updateSimOrderStatus(orderId, newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheets().filter(s => s.getSheetId() === SIM_GID)[0];
  if (!sheet) return { success: false, error: 'SIM sheet not found' };
  const rows = sheet.getDataRange().getValues();
  const startRow = String(rows[0][0]).toLowerCase().includes('date') ? 1 : 0;
  for (let i = startRow; i < rows.length; i++) {
    if (String(rows[i][1] || '') === String(orderId)) {
      sheet.getRange(i + 1, 6).setValue(newStatus || 'Approved');
      return { success: true };
    }
  }
  return { success: false, error: 'Order not found: ' + orderId };
}
