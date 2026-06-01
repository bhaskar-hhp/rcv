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
        result = { success: true, message: 'Config saved' };
        break;

      case 'getConfig':
        const cfg = getProps();
        result = { success: true, userName: cfg.userName, userId: cfg.userId, hasCookies: !!cfg.cookies };
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
        const today = Utilities.formatDate(new Date(), 'IST', 'dd-MMM-yyyy');
        sheet.appendRow([today, orderId || '', customerNum, amount, 'Pending']);
      }
    } catch (e) {}

    return { success: true, orderId, messages };
  } else {
    return { success: false, error: 'HTTP ' + result.status, data: result.data };
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

  const qs = Object.entries(params).map(([k, v]) => k + '=' + encodeURIComponent(v)).join('&');
  const path = '/api/dsm-orders/e-topup-orders?' + qs;

  const result = jioApi('GET', path, null, fullName, userId);
  if (result.status === 200) {
    return { success: true, count: (result.data || []).length, data: result.data };
  }
  return { success: false, error: 'HTTP ' + result.status, data: result.data };
}
