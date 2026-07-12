// ============================================================
// CHUAY Document System — Google Apps Script Backend v4.0
// Features: LIFF Login + Company Profile + All 6 Documents
// ============================================================

const SPREADSHEET_ID = '1ckKnGRcZS7RLu2qftr3WpdmDYSH4HR6oxioe3fTYH3g';
const LINE_LOGIN_CHANNEL_ID = '2010618788';

// Where a customer actually goes to subscribe/upgrade/pay right now, since
// payment isn't automated yet — every "ทักแอดมิน" message the bot sends
// MUST include this, or the customer has no real next step (this is the
// exact bug reported: the bot said "ทักแอดมิน" with no link/ID attached).
const ADMIN_CONTACT_NAME = '@แอดมินง่าย';
const ADMIN_CONTACT_LINE_ID = '@931btrgh';
const ADMIN_CONTACT_URL = 'https://page.line.biz/account/@931btrgh';
function adminContactBlock_() {
  return '👤 ' + ADMIN_CONTACT_NAME + '\n' + ADMIN_CONTACT_URL + '\n(LINE ID: ' + ADMIN_CONTACT_LINE_ID + ')';
}

// ⚠️ REPLACE THIS before relying on the QR-in-chat flow below: this must be
// a real HTTPS link to an image of your actual PromptPay/bank QR code
// (upload it anywhere public — Google Drive "anyone with the link" share,
// imgur, etc.). LINE's image message type requires a direct image URL, not
// a page URL. Left blank/placeholder, the bot will just skip the QR image
// and fall back to text-only instructions (see buildSubscribeReply_).
const PAYMENT_QR_IMAGE_URL = '';

// Bank transfer as the payment method while the PromptPay QR is still
// pending bank approval — swap PAYMENT_QR_IMAGE_URL in once that's ready;
// you can keep both shown (QR + bank details) or remove this block later,
// whichever you prefer.
// ⚠️ DOUBLE-CHECK THE ACCOUNT NUMBER AGAINST THE PHYSICAL PASSBOOK BEFORE
// DEPLOYING — this was transcribed from a photo and a single misread digit
// here means a customer's real money goes to the wrong account. Verify all
// 10 digits, not just the first few, before this goes live.
const BANK_NAME = 'ธนาคารกสิกรไทย (KBank)';
const BANK_ACCOUNT_NO = '215-1-32925-2';
const BANK_ACCOUNT_NAME = 'บจก. หลงใหล อินเตอร์เทรด';
function bankTransferBlock_() {
  return '🏦 โอนเงินผ่านบัญชีธนาคาร\n' +
    BANK_NAME + '\n' +
    'เลขบัญชี: ' + BANK_ACCOUNT_NO + '\n' +
    'ชื่อบัญชี: ' + BANK_ACCOUNT_NAME;
}

// ⚠️ SET THIS before the admin-approve feature will work: your own LINE
// user ID (the one that messages this OA when YOU chat with it, not the OA
// account's own ID/handle). Bootstrap it by messaging the bot the exact
// phrase "ไอดีของฉัน" once — it'll reply with your LINE User ID, which you
// then paste in here and redeploy. Left blank, the "อนุมัติ ..." command
// below is disabled entirely (falls through to the normal AI/menu flow) —
// this is intentional so a random customer can't approve their own package
// just by guessing the command.
const ADMIN_LINE_USER_ID = 'Ue73bf6d11364bfbb9fb4df6c727a25b5';

function isAdmin_(lineUserId) {
  return !!ADMIN_LINE_USER_ID && lineUserId === ADMIN_LINE_USER_ID;
}

function getPendingSubscriptionSheet_(ss) {
  const headers = ['Timestamp', 'LINE User ID', 'Display Name', 'Requested Package', 'Package ID', 'Status'];
  let sheet = ss.getSheetByName('PendingSubscriptions');
  if (!sheet) {
    sheet = ss.insertSheet('PendingSubscriptions');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
  }
  return sheet;
}

// Logs a subscribe request so it's sitting in one auditable place (instead
// of scattered across LINE chat history) for the admin to cross-check
// against incoming bank/PromptPay transfers — AND (if ADMIN_LINE_USER_ID is
// set) pushes a real-time notice straight to the admin's LINE with the
// exact command to type to activate it, so opening the sheet is optional
// rather than required every time.
function logPendingSubscription_(lineUserId, displayName, pkg) {
  const name = displayName || lineUserId;
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getPendingSubscriptionSheet_(ss);
    sheet.appendRow([new Date(), lineUserId, name, pkg.name + ' (' + pkg.price + ' บาท)', pkg.id, 'รอตรวจสอบ']);
  } catch (err) {
    console.log('logPendingSubscription_ failed:', err);
  }
  if (ADMIN_LINE_USER_ID) {
    pushToLine_(ADMIN_LINE_USER_ID, [{
      type: 'text',
      text: '🔔 คำขอสมัครใหม่\nชื่อ: ' + name + '\nแพ็ก: ' + pkg.name + ' (' + pkg.price + ' บาท)\n\nพิมพ์ "อนุมัติ ' + name + '" เพื่อเปิดใช้งานให้ลูกค้าคนนี้ได้เลยค่ะ'
    }]);
  }
}

function getLineDisplayName_(lineUserId) {
  try {
    const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
    if (!token || !lineUserId) return '';
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/profile/' + lineUserId, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    const data = JSON.parse(res.getContentText());
    return data.displayName || '';
  } catch (err) {
    return '';
  }
}

// Admin types "อนุมัติ <ชื่อลูกค้า>" in this same chat — matched against the
// most recent still-pending row in PendingSubscriptions for that name (does
// NOT require opening the Google Sheet). Returns a reply-message array if
// this text WAS an approve command (whether it succeeded or found nothing),
// or null if it wasn't an approve command at all — callers should fall
// through to normal handling on null.
function tryHandleAdminApprove_(text, lineUserId) {
  if (!isAdmin_(lineUserId)) return null;
  const m = (text || '').match(/^อนุมัติ\s*(.+)$/);
  if (!m) return null;
  const name = m[1].trim();
  if (!name) return [{ type: 'text', text: 'พิมพ์ชื่อลูกค้าต่อท้ายด้วยค่ะ เช่น "อนุมัติ สมชาย"' }];

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const pendingSheet = getPendingSubscriptionSheet_(ss);
  const data = pendingSheet.getDataRange().getValues();
  const lowerName = name.toLowerCase();

  // Search bottom-up so the most recent request for that name wins.
  let match = null;
  for (let i = data.length - 1; i >= 1; i--) {
    const rowName = String(data[i][2] || '').toLowerCase();
    const status = data[i][5];
    if (status === 'รอตรวจสอบ' && (rowName.indexOf(lowerName) > -1 || lowerName.indexOf(rowName) > -1)) {
      match = { rowIndex: i + 1, row: data[i] };
      break;
    }
  }
  if (!match) {
    return [{ type: 'text', text: 'ไม่พบคำขอที่รอดำเนินการของ "' + name + '" ค่ะ (เช็คสะกดชื่อ หรือดูในชีท PendingSubscriptions ได้เลยนะคะ)' }];
  }

  const customerLineUserId = match.row[1];
  const customerName = match.row[2];
  const pkgId = match.row[4];
  const pkg = PACKAGES[pkgId];
  if (!pkg) {
    return [{ type: 'text', text: 'เจอคำขอของ "' + customerName + '" แต่รหัสแพ็กเกจในชีทไม่ถูกต้อง (' + pkgId + ') กรุณาแก้ในชีท Subscriptions ด้วยตัวเองนะคะ' }];
  }

  activateSubscription_(customerLineUserId, pkgId);
  pendingSheet.getRange(match.rowIndex, 6).setValue('อนุมัติแล้ว ' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm'));

  pushToLine_(customerLineUserId, [{
    type: 'text',
    text: '🎉 เปิดใช้งานแพ็ก "' + pkg.name + '" ให้เรียบร้อยแล้วค่ะ ขอบคุณที่ใช้บริการนะคะ 💕'
  }]);

  return [{ type: 'text', text: '✅ เปิดใช้งานแพ็ก "' + pkg.name + '" ให้ ' + customerName + ' เรียบร้อยแล้วค่ะ' }];
}

// Writes/updates a customer's row in "Subscriptions" directly — used by the
// admin-approve command above. Gives them 30 days from today and resets
// their usage counter for the new billing cycle.
function activateSubscription_(lineUserId, packageId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getSubscriptionSheet_(ss);
  const now = new Date();
  const expiry = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
  const monthKey = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM');
  const found = findSubscriptionRow_(sheet, lineUserId);
  if (found) {
    sheet.getRange(found.rowIndex, 1, 1, 7).setValues([[lineUserId, packageId, now, expiry, monthKey, 0, now]]);
  } else {
    sheet.appendRow([lineUserId, packageId, now, expiry, monthKey, 0, now]);
  }
}

// Bootstrap helper: message the bot "ไอดีของฉัน" once to get your own LINE
// User ID back, so you can paste it into ADMIN_LINE_USER_ID above.
function isWhoAmIQuery_(text) {
  return (text || '').trim() === 'ไอดีของฉัน';
}


// - If the message also names/numbers a package (e.g. "สมัครแพ็คโปร"), log
//   the request immediately and reply with payment instructions + QR.
// - Otherwise show the package list again and ask them to pick one, plus
//   the direct admin contact for anyone who'd rather just message a person.
function buildSubscribeReply_(text, lineUserId) {
  const chosen = resolvePackageChoice_(text);
  if (chosen) {
    const displayName = getLineDisplayName_(lineUserId);
    logPendingSubscription_(lineUserId, displayName, chosen);
    let msg = '✅ รับคำขอสมัคร "' + chosen.name + '" (' + chosen.price + ' บาท/เดือน) แล้วนะคะ 🌸\n\n' +
      'กรุณาโอนเงิน ' + chosen.price + ' บาท ตามช่องทางด้านล่าง แล้วส่งสลิปให้แอดมินเพื่อยืนยันการเปิดใช้งานค่ะ\n\n' +
      bankTransferBlock_() +
      (PAYMENT_QR_IMAGE_URL ? '\n\n(หรือสแกน QR ที่แนบด้านล่างนี้ก็ได้ค่ะ)' : '') +
      '\n\n' + adminContactBlock_();
    if (PAYMENT_QR_IMAGE_URL) {
      return [
        { type: 'text', text: msg },
        { type: 'image', originalContentUrl: PAYMENT_QR_IMAGE_URL, previewImageUrl: PAYMENT_QR_IMAGE_URL }
      ];
    }
    return [{ type: 'text', text: msg }];
  }
  return [{
    type: 'text',
    text: buildPackageInfoMessage_() + '\n\nพิมพ์ชื่อแพ็กที่สนใจ (เช่น "สมัครแพ็คโปร") เพื่อแจ้งความประสงค์ได้เลยค่ะ 🌸'
  }];
}

// ── Packages ──────────────────────────────────────────────────
// Priced for the actual target segment: ตลาดนัด / แม่ค้า-พ่อค้าออนไลน์รายย่อย
// (mass-market micro-merchants), not SME/business buyers — keep these low
// until there's a separate SME-tier product.
// Payment collection is NOT automated yet — admin assigns/renews PAID
// packages manually by editing the "Subscriptions" sheet (LINE User ID +
// Package + Expiry Date). The "free" package below IS automated: it's
// auto-assigned the first time a new LINE user is seen (see
// getUserSubscription_), so nobody hits a dead-end "no_package" error on
// their very first try — that's the whole point of a free tier for
// acquisition with zero ad budget.
const PACKAGES = {
  free: {
    id: 'free', name: 'ทดลองใช้ฟรี', price: 0,
    quotaPerMonth: 5, hasAI: false, hasTax: false
  },
  starter59: {
    id: 'starter59', name: 'แพ็คเริ่มต้น', price: 59,
    quotaPerMonth: 30, hasAI: false, hasTax: false
  },
  pro149: {
    id: 'pro149', name: 'แพ็คโปร', price: 149,
    quotaPerMonth: null, hasAI: false, hasTax: false
  },
  premium249: {
    id: 'premium249', name: 'แพ็คพรีเมียม', price: 249,
    quotaPerMonth: null, hasAI: true, hasTax: true
  }
};
// Backward-compat aliases: rows already sitting in the live "Subscriptions"
// sheet from before this pricing update still say "trial99"/"pro299"/
// "premium990" in the Package column. Without these aliases, PACKAGES[id]
// would come back undefined for every one of those existing customers the
// moment this file is deployed, and getUserSubscription_ would treat a
// paying customer as having no package at all. Point the old ids at the
// same objects as their renamed replacements so nobody has to go hand-edit
// the sheet before this ships.
PACKAGES.trial99    = PACKAGES.starter59;
PACKAGES.pro299     = PACKAGES.pro149;
PACKAGES.premium990 = PACKAGES.premium249;

function getSubscriptionSheet_(ss) {
  const headers = ['LINE User ID', 'Package', 'Start Date', 'Expiry Date', 'Usage Month', 'Usage Count', 'Last Updated'];
  let sheet = ss.getSheetByName('Subscriptions');
  if (!sheet) {
    sheet = ss.insertSheet('Subscriptions');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
  }
  return sheet;
}

function findSubscriptionRow_(sheet, lineUserId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === lineUserId) return { rowIndex: i + 1, row: data[i] };
  }
  return null;
}

// Returns null if the user has no package or it (a PAID package) has expired.
// A brand-new LINE user (no row at all) is auto-enrolled into the "free"
// package on the spot — see createFreeTrialRow_ — instead of returning null,
// so the very first thing a new user does isn't hitting a paywall error.
function getUserSubscription_(lineUserId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getSubscriptionSheet_(ss);
  let found = findSubscriptionRow_(sheet, lineUserId);
  if (!found) {
    found = createFreeTrialRow_(sheet, lineUserId);
  }

  const packageId  = found.row[1];
  const expiryDate = found.row[3];
  const usageMonth = found.row[4];
  const usageCount = found.row[5];

  const pkg = PACKAGES[packageId];
  if (!pkg) return null; // unknown/typo'd package id in the sheet — treat as no package
  // Only PAID packages expire on a date; the free tier has no expiry date
  // and is instead capped purely by its monthly quota.
    if (packageId !== 'free' && expiryDate && new Date(expiryDate) < new Date()) {
    return { pkg: pkg, rowIndex: found.rowIndex, usageCount: 0, usageMonth: null, expired: true };
  }

  const currentMonthKey = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM');
  const usageMonthStr = String(usageMonth || '');
  const count = (usageMonthStr === currentMonthKey) ? Number(usageCount || 0) : 0;

  // Already handled above with type safety

  return { pkg: pkg, rowIndex: found.rowIndex, usageCount: count, usageMonth: currentMonthKey };
}

// Appends a new "free" package row for a LINE user we've never seen before.
// Runs once per user, right when they first authenticate or try to
// generate a document — no admin action required.
// Wrapped in a script lock: getUserSubscription_ is called both from
// getCurrentUser (on page load) and generatePDF (on submit), and those can
// land close together (two tabs, a fast double-tap). Without the lock, both
// could see "no row yet" and each append their own row for the same user —
// a silent duplicate that leaves stray rows in the sheet.
function createFreeTrialRow_(sheet, lineUserId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const existing = findSubscriptionRow_(sheet, lineUserId);
    if (existing) return existing; // another request just created it — reuse it
    const now = new Date();
    const monthKey = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM');
    const row = [lineUserId, 'free', now, '', monthKey, 0, now];
    sheet.appendRow(row);
    return { rowIndex: sheet.getLastRow(), row: row };
  } finally {
    lock.releaseLock();
  }
}

// Checks quota and, if allowed, increments usage for this month.
// Call this right before generating a document.
function checkAndConsumeQuota_(lineUserId) {
  const sub = getUserSubscription_(lineUserId);
  // Only reachable if the sheet has a genuinely unknown/typo'd package id
  // (free tier auto-enrolls every user — see getUserSubscription_).
  if (!sub) {
    return { allowed: false, reason: 'no_package' };
  }
  if (sub.expired) {
    return { allowed: false, reason: 'expired', pkg: sub.pkg };
  }
  if (sub.pkg.quotaPerMonth !== null && sub.usageCount >= sub.pkg.quotaPerMonth) {
    return {
      allowed: false,
      reason: (sub.pkg.id === 'free') ? 'free_quota_exceeded' : 'quota_exceeded',
      pkg: sub.pkg
    };
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getSubscriptionSheet_(ss);
  sheet.getRange(sub.rowIndex, 5).setValue(sub.usageMonth);      // Usage Month
  sheet.getRange(sub.rowIndex, 6).setValue(sub.usageCount + 1);  // Usage Count
  sheet.getRange(sub.rowIndex, 7).setValue(new Date());          // Last Updated
  return {
    allowed: true,
    pkg: sub.pkg,
    remaining: sub.pkg.quotaPerMonth !== null ? (sub.pkg.quotaPerMonth - sub.usageCount - 1) : null,
    // For rollbackQuotaUsage_ below — quota is consumed here, BEFORE we know
    // the PDF actually gets generated successfully. If Drive/PDF creation
    // fails afterward (storage full, transient error), the caller should
    // roll this back so the customer isn't charged a quota unit for a
    // document that was never actually produced.
    rowIndex: sub.rowIndex,
    previousUsageCount: sub.usageCount
  };
}

// Restores a customer's usage counter after a consumed quota unit turned
// out not to correspond to an actual generated document (see generatePDF's
// catch block). Safe to call even if something about the row changed
// in between — worst case it just re-writes the same value.
function rollbackQuotaUsage_(rowIndex, previousUsageCount) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getSubscriptionSheet_(ss);
    sheet.getRange(rowIndex, 6).setValue(previousUsageCount);
  } catch (err) {
    console.log('rollbackQuotaUsage_ failed:', err);
  }
}

// ── Entry point ──────────────────────────────────────────────
// NOTE: The LIFF frontend is now hosted OUTSIDE Apps Script (e.g. GitHub
// Pages) to avoid the iframe-wrapping issue that blocks liff.login().
// doGet is kept only as a fallback / health check — it is no longer the
// page the LIFF endpoint URL points to.
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: 'CHUAY backend is running' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── API router (called via fetch() from the externally-hosted frontend) ──
// Expects POST body: { action: 'functionName', payload: {...} }
// Returns: JSON of whatever the target function returns.
const API_ACTIONS = {
  getCurrentUser:     function(p) { return getCurrentUser(p.lineAuth); },
  getCompanyProfile:  function(p) { return getCompanyProfile(p.lineAuth); },
  saveCompanyProfile: function(p) { return saveCompanyProfile(p.profile, p.lineAuth); },
  getDocHistory:      function(p) { return getDocHistory(p.docType, p.lineAuth); },
  generatePDF:        function(p) { return generatePDF(p.formData, p.lineAuth); },
  getCustomers:       function(p) { return getCustomers(p.lineAuth); },
  getCashTransactions:    function(p) { return getCashTransactions(p.lineAuth); },
  saveCashTransaction:    function(p) { return saveCashTransaction(p.transaction, p.lineAuth); },
  generateCashReportPDF:  function(p) { return generateCashReportPDF(p.reportData, p.lineAuth); }
};

function doPost(e) {
  // Defensive guard: LINE's webhook "Verify" button (and periodic health
  // checks) send a POST with no body at all, just to confirm the URL
  // returns 200. Without this guard, e.postData is undefined and the
  // JSON.parse below throws.
  if (!e || !e.postData || !e.postData.contents) {
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const body = JSON.parse(e.postData.contents);

  // LINE Messaging API webhook payloads look like { events: [...] } —
  // route these to the bot handler instead of the LIFF API router.
  if (body.events) {
    return handleLineWebhook_(body);
  }

  let result;
  try {
    const action = API_ACTIONS[body.action];
    if (!action) throw new Error('Unknown action: ' + body.action);
    result = action(body.payload || {});
  } catch (err) {
    result = { success: false, error: err.message || err.toString() };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── LIFF Auth ────────────────────────────────────────────────
function getCurrentUser(lineAuth) {
  try {
    const user = verifyLineUser_(lineAuth);
    const sub = getUserSubscription_(user.lineUserId);
    return {
      loggedIn: true,
      userKey: user.userKey,
      lineUserId: user.lineUserId,
      name: user.name,
      displayName: user.name,
      pictureUrl: user.pictureUrl,
      email: user.email || '',
      package: sub ? {
        id: sub.pkg.id,
        name: sub.pkg.name,
        quotaPerMonth: sub.pkg.quotaPerMonth,
        usageCount: sub.usageCount,
        hasAI: sub.pkg.hasAI,
        hasTax: sub.pkg.hasTax,
        expired: !!sub.expired
      } : null
    };
  } catch(e) {
    console.log('getCurrentUser error:', e);
    return { loggedIn: false, error: e.message || e.toString() };
  }
}

function verifyLineUser_(lineAuth) {
  if (!lineAuth || !lineAuth.idToken) {
    throw new Error('ไม่พบ LINE ID token กรุณาเข้าสู่ระบบผ่าน LIFF');
  }

  const response = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post',
    payload: {
      id_token: lineAuth.idToken,
      client_id: LINE_LOGIN_CHANNEL_ID
    },
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const body = response.getContentText();
  if (status < 200 || status >= 300) {
    throw new Error('ยืนยันตัวตน LINE ไม่สำเร็จ: ' + body);
  }

  const payload = JSON.parse(body);
  if (payload.aud !== LINE_LOGIN_CHANNEL_ID) {
    throw new Error('LINE token ไม่ตรงกับ Channel ID');
  }
  if (!payload.sub) {
    throw new Error('LINE token ไม่มี user id');
  }

  return {
    userKey: 'line:' + payload.sub,
    lineUserId: payload.sub,
    name: payload.name || (lineAuth.displayName || 'LINE User'),
    pictureUrl: payload.picture || lineAuth.pictureUrl || '',
    email: payload.email || ''
  };
}

// ── Company Profile ──────────────────────────────────────────
function getCompanyProfile(lineAuth) {
  try {
    const user = verifyLineUser_(lineAuth);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = getProfileSheet_(ss);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === user.userKey) {
        return {
          userKey:        data[i][0],
          displayName:    data[i][1] || user.name,
          pictureUrl:     data[i][2] || user.pictureUrl,
          companyName:    data[i][3] || '',
          companyAddress: data[i][4] || '',
          companyTaxId:   data[i][5] || '',
          companyBranch:  data[i][6] || '',
          companyPhone:   data[i][7] || '',
          companyEmail:   data[i][8] || '',
          updatedAt:      data[i][9] || ''
        };
      }
    }
    return null;
  } catch(e) {
    console.log('getCompanyProfile error:', e);
    return null;
  }
}

function saveCompanyProfile(profile, lineAuth) {
  try {
    const user = verifyLineUser_(lineAuth);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = getProfileSheet_(ss);
    const now = Utilities.formatDate(new Date(),'Asia/Bangkok','dd/MM/yyyy HH:mm');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === user.userKey) {
        sheet.getRange(i+1, 1, 1, 10).setValues([[
          user.userKey,
          user.name,
          user.pictureUrl,
          profile.companyName    || '',
          profile.companyAddress || '',
          profile.companyTaxId   || '',
          profile.companyBranch  || '',
          profile.companyPhone   || '',
          profile.companyEmail   || '',
          now
        ]]);
        return { success: true, message: 'อัปเดตข้อมูลแล้ว' };
      }
    }
    // New user
    sheet.appendRow([
      user.userKey,
      user.name,
      user.pictureUrl,
      profile.companyName    || '',
      profile.companyAddress || '',
      profile.companyTaxId   || '',
      profile.companyBranch  || '',
      profile.companyPhone   || '',
      profile.companyEmail   || '',
      now
    ]);
    return { success: true, message: 'บันทึกข้อมูลแล้ว' };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function getProfileSheet_(ss) {
  const headers = ['UserKey','ชื่อ LINE','รูป LINE','ชื่อบริษัท','ที่อยู่','เลขภาษี','สำนักงาน','เบอร์โทร','อีเมลบริษัท','อัปเดตเมื่อ'];
  let sheet = ss.getSheetByName('โปรไฟล์บริษัท');
  if (!sheet) {
    sheet = ss.insertSheet('โปรไฟล์บริษัท');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#f0f0f0');
  return sheet;
}

// ── Customer memory ──────────────────────────────────────────
// Remembers each business owner's own customers/vendors (by name) so they
// don't have to retype address/tax ID/phone on every new document.
// Scoped per owner (LINE User ID) — each business only sees its own list.
function getCustomerSheet_(ss) {
  const headers = ['Owner LINE User ID', 'Customer Name', 'Address', 'Tax ID', 'Branch', 'Phone', 'Last Used'];
  let sheet = ss.getSheetByName('Customers');
  if (!sheet) {
    sheet = ss.insertSheet('Customers');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
  }
  return sheet;
}

function getCustomers(lineAuth) {
  try {
    const user = verifyLineUser_(lineAuth);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getCustomerSheet_(ss);
    const data = sheet.getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === user.lineUserId && data[i][1]) {
        list.push({
          name: data[i][1], address: data[i][2], taxId: data[i][3],
          branch: data[i][4], phone: data[i][5], lastUsed: data[i][6]
        });
      }
    }
    list.sort(function(a, b) { return new Date(b.lastUsed) - new Date(a.lastUsed); });
    return list;
  } catch (e) {
    return []; // not logged in / error — just show no suggestions rather than break the form
  }
}

// Upsert (by owner + exact name match). Called automatically from
// generatePDF every time a document is created — the customer/vendor typed
// this time becomes a saved suggestion for next time.
function saveCustomerFromDoc_(ownerLineUserId, name, address, taxId, branch, phone) {
  if (!ownerLineUserId || !name) return;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getCustomerSheet_(ss);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === ownerLineUserId && data[i][1] === name) {
      sheet.getRange(i + 1, 3, 1, 5).setValues([[address, taxId, branch, phone, new Date()]]);
      return;
    }
  }
  sheet.appendRow([ownerLineUserId, name, address, taxId, branch, phone, new Date()]);
}

// ── Accounting: cash income/expense log (premium249 preview) ──────────
// Scoped per business owner (LINE User ID). Reuses the same company
// profile (getCompanyProfile/saveCompanyProfile) as the main document app
// for the report header, so nothing needs to be entered twice.
function getCashTransactionSheet_(ss) {
  const headers = ['LINE User ID', 'Date', 'Description', 'Income', 'Expense (Goods)', 'Expense (Other)', 'Remarks', 'Created At'];
  let sheet = ss.getSheetByName('CashTransactions');
  if (!sheet) {
    sheet = ss.insertSheet('CashTransactions');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
  }
  return sheet;
}

// Shared gate for the three accounting endpoints below — this feature is
// advertised (see buildTaxFeatureMessage_) as premium249-only, but until now
// nothing actually enforced that server-side, so anyone who knew the Apps
// Script URL could call these three actions directly and use the "paid"
// feature for free regardless of package.
function requireTaxAccess_(lineUserId) {
  const sub = getUserSubscription_(lineUserId);
  if (!sub || sub.expired || !sub.pkg.hasTax) {
    throw new Error('ฟีเจอร์นี้อยู่ในแพ็คพรีเมียม 249 บาท/เดือนเท่านั้นค่ะ พิมพ์ "แพ็กเกจ" ในแชท LINE เพื่อดูรายละเอียดและอัปเกรดได้เลยนะคะ');
  }
}

function getCashTransactions(lineAuth) {
  try {
    const user = verifyLineUser_(lineAuth);
    requireTaxAccess_(user.lineUserId);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getCashTransactionSheet_(ss);
    const data = sheet.getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === user.lineUserId) {
        list.push({
          date: Utilities.formatDate(new Date(data[i][1]), 'Asia/Bangkok', 'yyyy-MM-dd'),
          itemDescription: data[i][2],
          income: Number(data[i][3] || 0),
          expenseGoods: Number(data[i][4] || 0),
          expenseOther: Number(data[i][5] || 0),
          remarks: data[i][6] || ''
        });
      }
    }
    return { success: true, transactions: list };
  } catch (e) {
    return { success: false, error: e.message || e.toString() };
  }
}

// Auto-import: called from generatePDF whenever a "ใบเสร็จรับเงิน" (receipt)
// is generated. Logs the receipt's total as income in CashTransactions,
// tagged so it's traceable back to the source document. Uses the PDF's
// own doc number as a dedupe key so re-generating the same receipt number
// doesn't double-count income.
function importReceiptAsIncome_(lineUserId, formData) {
  try {
    const total = calcTotal(formData).grand;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getCashTransactionSheet_(ss);
    const data = sheet.getDataRange().getValues();
    const dedupeTag = '[receipt:' + (formData.docNo || '') + ']';
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === lineUserId && String(data[i][6] || '').indexOf(dedupeTag) > -1) {
        return; // already imported this receipt — don't double count
      }
    }
    sheet.appendRow([
      lineUserId,
      formData.docDate || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd'),
      'รับชำระจาก ' + (formData.customerName || formData.vendorName || 'ลูกค้า') + ' (' + (formData.docNo || '') + ')',
      total,
      0,
      0,
      'นำเข้าอัตโนมัติจากใบเสร็จรับเงิน ' + dedupeTag,
      new Date()
    ]);
  } catch (e) {
    // Never let an accounting-sync hiccup break PDF generation itself.
  }
}

function saveCashTransaction(transaction, lineAuth) {
  try {
    const user = verifyLineUser_(lineAuth);
    requireTaxAccess_(user.lineUserId);
    if (!transaction || !transaction.date || !transaction.itemDescription) {
      throw new Error('กรุณากรอกวันที่และรายการให้ครบถ้วนค่ะ');
    }
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getCashTransactionSheet_(ss);
    sheet.appendRow([
      user.lineUserId,
      transaction.date,
      transaction.itemDescription,
      Number(transaction.income || 0),
      Number(transaction.expenseGoods || 0),
      Number(transaction.expenseOther || 0),
      transaction.remarks || '',
      new Date()
    ]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message || e.toString() };
  }
}

// Builds a simple cash income/expense report PDF for a date range.
// Fixes the earlier "pdf.setSharing is not a function" bug: that error
// happens when .setSharing() is called on a Blob instead of a Drive File.
// The fix is the same pattern generatePDF() already uses correctly —
// create the file with DriveApp first, THEN call .setSharing() on the
// resulting File object, never on the raw Blob.
function generateCashReportPDF(reportData, lineAuth) {
  let tmpFile = null;
  try {
    const user = verifyLineUser_(lineAuth);
    requireTaxAccess_(user.lineUserId);
    if (!reportData || !reportData.startDate || !reportData.endDate) {
      throw new Error('กรุณาเลือกช่วงวันที่สำหรับรายงานค่ะ');
    }

    const profile = getCompanyProfile(lineAuth) || {};
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getCashTransactionSheet_(ss);
    const data = sheet.getDataRange().getValues();
    const start = new Date(reportData.startDate);
    const end = new Date(reportData.endDate);
    end.setHours(23, 59, 59, 999);

    const rows = [];
    let totalIncome = 0, totalGoods = 0, totalOther = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== user.lineUserId) continue;
      const d = new Date(data[i][1]);
      if (d < start || d > end) continue;
      const income = Number(data[i][3] || 0);
      const goods  = Number(data[i][4] || 0);
      const other  = Number(data[i][5] || 0);
      totalIncome += income; totalGoods += goods; totalOther += other;
      rows.push({
        date: Utilities.formatDate(d, 'Asia/Bangkok', 'dd/MM/yyyy'),
        desc: data[i][2], income: income, goods: goods, other: other, remarks: data[i][6] || ''
      });
    }

    const html = buildCashReportHTML_(profile, reportData.startDate, reportData.endDate, rows, totalIncome, totalGoods, totalOther);

    const folder = DriveApp.getRootFolder();
    const tmpBlob = Utilities.newBlob(html, 'text/html', 'tmp_report.html');
    tmpFile = folder.createFile(tmpBlob);
    const pdfBlob = tmpFile.getAs('application/pdf');
    const fileName = makeSafeFileName_('รายงานเงินสด ' + reportData.startDate + ' ถึง ' + reportData.endDate) + '.pdf';
    pdfBlob.setName(fileName);

    // IMPORTANT: create the Drive File from the blob FIRST, then call
    // .setSharing() on that File — never on the Blob itself.
    const pdfFile = folder.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    tmpFile.setTrashed(true);
    tmpFile = null;

    return { success: true, pdfUrl: pdfFile.getUrl(), fileName: fileName };
  } catch (e) {
    if (tmpFile) { try { tmpFile.setTrashed(true); } catch (_) {} }
    return { success: false, error: e.message || e.toString() };
  }
}

function buildCashReportHTML_(profile, startDate, endDate, rows, totalIncome, totalGoods, totalOther) {
  const net = totalIncome - totalGoods - totalOther;
  const rowsHtml = rows.map(function(r) {
    return '<tr>' +
      '<td>' + esc_(r.date) + '</td>' +
      '<td>' + esc_(r.desc) + '</td>' +
      '<td style="text-align:right">' + r.income.toLocaleString('th-TH', {minimumFractionDigits:2}) + '</td>' +
      '<td style="text-align:right">' + r.goods.toLocaleString('th-TH', {minimumFractionDigits:2}) + '</td>' +
      '<td style="text-align:right">' + r.other.toLocaleString('th-TH', {minimumFractionDigits:2}) + '</td>' +
      '<td>' + esc_(r.remarks) + '</td>' +
    '</tr>';
  }).join('');

  return '<!DOCTYPE html><html lang="th"><head><meta charset="UTF-8"><style>' +
    'body{font-family:Sarabun,sans-serif;padding:24px;color:#1a1a1a;font-size:12px}' +
    'h1{font-size:18px;margin-bottom:2px}' +
    '.sub{color:#555;margin-bottom:16px;font-size:11px}' +
    'table{width:100%;border-collapse:collapse;margin-top:12px}' +
    'th,td{border:1px solid #ccc;padding:6px 8px;font-size:11px}' +
    'th{background:#f0f0f0;text-align:left}' +
    '.totals{margin-top:16px;max-width:320px;margin-left:auto}' +
    '.totals div{display:flex;justify-content:space-between;padding:4px 0;border-bottom:1px solid #eee}' +
    '.totals .net{font-weight:700;font-size:13px;border-top:2px solid #1a1a1a;margin-top:4px;padding-top:6px}' +
    '</style></head><body>' +
    '<h1>' + esc_(profile.companyName || 'รายงานเงินสดรับ-จ่าย') + '</h1>' +
    '<div class="sub">' + esc_(profile.companyAddress || '') + (profile.companyTaxId ? (' · เลขประจำตัวผู้เสียภาษี/บัตรประชาชน ' + esc_(profile.companyTaxId)) : '') + '</div>' +
    '<div class="sub">รายงานเงินสดรับ-จ่าย ช่วงวันที่ ' + esc_(startDate) + ' ถึง ' + esc_(endDate) + '</div>' +
    '<table><thead><tr><th>วันที่</th><th>รายการ</th><th style="text-align:right">รายรับ</th><th style="text-align:right">ซื้อสินค้า</th><th style="text-align:right">อื่นๆ</th><th>หมายเหตุ</th></tr></thead>' +
    '<tbody>' + (rowsHtml || '<tr><td colspan="6" style="text-align:center;color:#888">ไม่มีรายการในช่วงที่เลือก</td></tr>') + '</tbody></table>' +
    '<div class="totals">' +
      '<div><span>รวมรายรับ</span><span>' + totalIncome.toLocaleString('th-TH', {minimumFractionDigits:2}) + ' บาท</span></div>' +
      '<div><span>รวมซื้อสินค้า</span><span>' + totalGoods.toLocaleString('th-TH', {minimumFractionDigits:2}) + ' บาท</span></div>' +
      '<div><span>รวมค่าใช้จ่ายอื่นๆ</span><span>' + totalOther.toLocaleString('th-TH', {minimumFractionDigits:2}) + ' บาท</span></div>' +
      '<div class="net"><span>กำไร/ขาดทุนสุทธิ</span><span>' + net.toLocaleString('th-TH', {minimumFractionDigits:2}) + ' บาท</span></div>' +
    '</div>' +
    '</body></html>';
}

// ── PDF Generator ────────────────────────────────────────────
function generatePDF(formData, lineAuth) {
  let tmpFile = null;
  let quota = null; // declared here (not inside try) so the catch block below can see it
  try {
    const user = verifyLineUser_(lineAuth);

    // [Fix] Validate inputs BEFORE consuming quota
    if (!formData.companyName) throw new Error('กรุณากรอกชื่อบริษัท');
    if (!formData.docNo)       throw new Error('กรุณากรอกเลขที่เอกสาร');
    if (!formData.docDate)     throw new Error('กรุณาเลือกวันที่ออกเอกสาร');
    if (!formData.items || formData.items.length === 0) throw new Error('กรุณาเพิ่มรายการสินค้าอย่างน้อย 1 รายการ');

    quota = checkAndConsumeQuota_(user.lineUserId);
    if (!quota.allowed) {
      if (quota.reason === 'no_package') {
        return { success: false, upgradeRequired: true, error: 'บัญชีนี้ยังไม่มีแพ็กเกจที่ใช้งานได้ กรุณาลองเข้าสู่ระบบใหม่อีกครั้ง หรือทักแอดมิน' };
      }
      if (quota.reason === 'expired') {
        return { success: false, upgradeRequired: true, error: 'แพ็ก "' + quota.pkg.name + '" ของคุณหมดอายุแล้ว กรุณาต่ออายุเพื่อใช้งานต่อค่ะ (ทักแอดมินได้เลยนะคะ)' };
      }
      if (quota.reason === 'free_quota_exceeded') {
        return { success: false, upgradeRequired: true, error: 'ใช้งานฟรีครบ ' + quota.pkg.quotaPerMonth + ' ครั้ง/เดือนแล้ว สมัครแพ็กเริ่มต้นเพียง 59 บาท/เดือน เพื่อใช้งานต่อได้เลยค่ะ' };
      }
      return { success: false, upgradeRequired: true, error: 'ใช้งานครบโควตา ' + quota.pkg.quotaPerMonth + ' ครั้ง/เดือนของแพ็ก "' + quota.pkg.name + '" แล้ว กรุณาอัปเกรดแพ็กเกจเพื่อใช้งานต่อ' };
    }

    const html = buildDocumentHTML(formData);
    const folder = DriveApp.getRootFolder();
    const tmpBlob = Utilities.newBlob(html, 'text/html', 'tmp.html');
    tmpFile = folder.createFile(tmpBlob);
    const pdfBlob = tmpFile.getAs('application/pdf');
    const fileName = makeSafeFileName_(formData.docTypeName + ' ' + formData.docNo + ' - ' + (formData.customerName || formData.vendorName || '')) + '.pdf';
    pdfBlob.setName(fileName);
    const pdfFile = folder.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Cleanup tmp file
    tmpFile.setTrashed(true);
    tmpFile = null;

    saveRecord(formData, pdfFile.getUrl(), user);
    saveCustomerFromDoc_(
      user.lineUserId,
      formData.customerName  || formData.vendorName  || '',
      formData.customerAddress || formData.vendorAddress || '',
      formData.customerTaxId   || formData.vendorTaxId   || '',
      formData.customerBranch  || formData.vendorBranch  || '',
      formData.customerPhone   || formData.vendorPhone   || ''
    );
    
    if (formData.docType === 'receipt') {
      importReceiptAsIncome_(user.lineUserId, formData);
    }
    
    return {
      success: true,
      url: pdfFile.getUrl(),
      fileName: fileName,
      remainingQuota: quota.remaining
    };
  } catch(e) {
    // If quota was already consumed (checkAndConsumeQuota_ ran and
    // succeeded) but something after it failed — Drive error, a bug in
    // buildDocumentHTML, etc — give the quota unit back. Otherwise a
    // transient failure silently costs the customer one of their paid
    // document credits for nothing.
    if (quota && quota.allowed) {
      rollbackQuotaUsage_(quota.rowIndex, quota.previousUsageCount);
    }
    return { success: false, error: e.message || e.toString() };
  } finally {
    if (tmpFile) {
      try { tmpFile.setTrashed(true); } catch(err) {}
    }
  }
}

// ── Save record ───────────────────────────────────────────────
function saveRecord(data, pdfUrl, user) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    // Always the Thai display name (e.g. "ใบเสนอราคา") — this MUST stay
    // consistent with what's already in the live sheet. A previous edit
    // briefly changed this to the English docType id ('quotation'), which
    // would have started splitting every doc type's history across two
    // differently-named sheet tabs (the old Thai one + a new English one)
    // the moment someone generated a document — reverted.
    const sheetName = data.docTypeName || 'เอกสาร';
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(['เลขที่','วันที่','ชื่อบริษัทเรา','ชื่อลูกค้า/ผู้รับ','ยอดรวม (บาท)','ลิงก์ PDF','ผู้สร้าง','สร้างเมื่อ']);
      sheet.getRange(1,1,1,8).setFontWeight('bold').setBackground('#f0f0f0');
    }
    const totals = calcTotal(data);
    const now = Utilities.formatDate(new Date(),'Asia/Bangkok','dd/MM/yyyy HH:mm');
    sheet.appendRow([
      data.docNo,
      data.docDate,
      data.companyName,
      data.customerName || data.vendorName || '-',
      totals.grand,
      pdfUrl,
      user.userKey || '-',
      now
    ]);
  } catch(e) {
    console.log('saveRecord error:', e);
  }
}

// ── Get document history ──────────────────────────────────────
function getDocHistory(docType, lineAuth) {
  try {
    const user = verifyLineUser_(lineAuth);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(docType); // docType here is actually the Thai display name — see the frontend's call site
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][6] === user.userKey) {
        result.push({
          docNo:    data[i][0],
          docDate:  data[i][1],
          customer: data[i][3],
          total:    data[i][4],
          url:      data[i][5],
          created:  data[i][7]
        });
        if (result.length >= 5) break; // แสดง 5 รายการล่าสุด
      }
    }
    return result;
  } catch(e) {
    return [];
  }
}

// ── Calculations ──────────────────────────────────────────────
function calcTotal(data) {
  const items = data.items || [];
  let subtotal = 0;
  items.forEach(function(item) {
    const qty   = parseFloat(item.qty   || 0);
    const price = parseFloat(item.price || 0);
    const disc  = parseFloat(item.discount || 0);
    subtotal += Math.max((qty * price) - disc, 0);
  });
  const vatAmt = data.includeVat ? Math.round(subtotal * 7) / 100 : 0;
  return { subtotal: subtotal, vat: vatAmt, grand: subtotal + vatAmt };
}

function makeSafeFileName_(name) {
  return String(name || 'document')
    .replace(/[\\/:*?"<>|#%{}~&]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 160) || 'document';
}

function esc_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
    .replace(/\n/g, '<br>');
}

// ── Thai Baht text ────────────────────────────────────────────
function numberToThaiText(amount) {
  const ones = ['','หนึ่ง','สอง','สาม','สี่','ห้า','หก','เจ็ด','แปด','เก้า'];
  const digits = ['','สิบ','ร้อย','พัน','หมื่น','แสน','ล้าน'];
  function readGroup(n) {
    if (n === 0) return '';
    const s = String(n).split('').reverse();
    let r = '';
    for (let i = 0; i < s.length; i++) {
      const d = parseInt(s[i]);
      if (d === 0) continue;
      if (i === 1 && d === 1) r = 'สิบ' + r;
      else if (i === 1 && d === 2) r = 'ยี่สิบ' + r;
      // Thai reading rule: the units digit "1" is read as "เอ็ด" (not
      // "หนึ่ง") whenever there's a tens digit or higher in the same group
      // — e.g. 11="สิบเอ็ด", 21="ยี่สิบเอ็ด", 101="หนึ่งร้อยเอ็ด".
      // Only a bare "1" with nothing else in the group stays "หนึ่ง".
      else if (i === 0 && d === 1 && s.length > 1) r = 'เอ็ด' + r;
      else r = ones[d] + digits[i] + r;
    }
    return r;
  }
  if (!amount || isNaN(amount)) return 'ศูนย์บาทถ้วน';
  const rounded = Math.round(amount * 100) / 100;
  const parts = rounded.toFixed(2).split('.');
  const baht = parseInt(parts[0]);
  const satang = parseInt(parts[1]);
  let text = '';
  if (baht >= 1000000) {
    text += readGroup(Math.floor(baht / 1000000)) + 'ล้าน';
    text += readGroup(baht % 1000000);
  } else {
    text += readGroup(baht);
  }
  text += 'บาท';
  text += satang > 0 ? readGroup(satang) + 'สตางค์' : 'ถ้วน';
  return text || 'ศูนย์บาทถ้วน';
}

// ── HTML Builder ──────────────────────────────────────────────
function buildDocumentHTML(d) {
  const totals = calcTotal(d);
  const sub = totals.subtotal, vatAmt = totals.vat, total = totals.grand;
  const amountInWords = numberToThaiText(total);
  const rows = (d.items||[]).map(function(item,idx) {
    const qty=parseFloat(item.qty||0), price=parseFloat(item.price||0), disc=parseFloat(item.discount||0);
    const lineTotal=Math.max((qty*price)-disc, 0);
    return '<tr>'
      +'<td class="tc">'+(idx+1)+'</td>'
      +'<td>'+esc_(item.name||'')+(item.unit?'<br><span style="font-size:8pt;color:#888">('+esc_(item.unit)+')</span>':'')+'</td>'
      +'<td class="tc">'+qty.toLocaleString('th-TH')+'</td>'
      +'<td class="tr">'+price.toLocaleString('th-TH',{minimumFractionDigits:2})+'</td>'
      +'<td class="tr">'+(disc>0?disc.toLocaleString('th-TH',{minimumFractionDigits:2}):'-')+'</td>'
      +'<td class="tr">'+lineTotal.toLocaleString('th-TH',{minimumFractionDigits:2})+'</td>'
      +'</tr>';
  }).join('');

  if (!d.docTypeName) d.docTypeName = 'ใบเสนอราคา';
  const typeConfig={
    'ใบเสนอราคา':    {rl:'เรียน',           el:'วันหมดอายุ',          ev:d.expireDate||'-',   sig:'ผู้มีอำนาจลงนาม', note:d.note||'ราคานี้มีผลภายใน 30 วันนับจากวันที่ออกเอกสาร'},
    'ใบแจ้งหนี้':    {rl:'เรียน',           el:'วันครบกำหนดชำระ',     ev:d.dueDate||'-',      sig:'ผู้มีอำนาจลงนาม', note:d.note||'กรุณาชำระเงินภายในวันที่กำหนด'},
    'ใบสั่งซื้อ':    {rl:'ผู้ขาย / Vendor', el:'วันที่ต้องการสินค้า', ev:d.requiredDate||'-', sig:'ผู้สั่งซื้อ',       note:d.note||'กรุณาจัดส่งสินค้าตามกำหนด'},
    'ใบเสร็จรับเงิน':{rl:'ผู้ชำระเงิน',    el:'วิธีชำระเงิน',        ev:d.paymentMethod||'-',sig:'ผู้รับเงิน',        note:d.note||'ได้รับชำระเงินเรียบร้อยแล้ว'},
    'ใบส่งของ':      {rl:'ส่งถึง',          el:'วันที่จัดส่ง',        ev:d.deliveryDate||'-', sig:'ผู้รับสินค้า',      note:d.note||'กรุณาตรวจสอบสินค้าก่อนรับมอบ'},
    'ใบวางบิล':      {rl:'เรียน',           el:'วันครบกำหนดชำระ',     ev:d.dueDate||'-',      sig:'ผู้มีอำนาจลงนาม', note:d.note||'กรุณาชำระเงินตามกำหนด'},
  };
  const cfg=typeConfig[d.docTypeName]||typeConfig['ใบเสนอราคา'];
  // [Fix] Support both Purchase Orders (Vendor) and Sales Docs (Customer)
  const isPO = d.docType === 'po' || d.docTypeName === 'ใบสั่งซื้อ';
  const rName = isPO ? (d.vendorName || d.customerName || '-') : (d.customerName || d.vendorName || '-');
  const rAddr = isPO ? (d.vendorAddress || d.customerAddress || '') : (d.customerAddress || d.vendorAddress || '');
  const rTax  = isPO ? (d.vendorTaxId || d.customerTaxId || '') : (d.customerTaxId || d.vendorTaxId || '');
  const rBranch = isPO ? (d.vendorBranch || d.customerBranch || '') : (d.customerBranch || d.vendorBranch || '');
  const rPhone  = isPO ? (d.vendorPhone || d.customerPhone || '') : (d.customerPhone || d.vendorPhone || '');
  return `<!DOCTYPE html><html lang="th"><head><meta charset="UTF-8"/>
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap');
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Sarabun',sans-serif;font-size:10pt;color:#111;background:#fff;-webkit-print-color-adjust:exact}
.page{width:210mm;min-height:297mm;padding:18mm 20mm;margin:0 auto}
.hdr{display:table;width:100%;border-bottom:2px solid #111;padding-bottom:6mm;margin-bottom:6mm}
.hl,.hr{display:table-cell;vertical-align:top}
.hl{width:62%}.hr{width:38%;text-align:right}
.co-name{font-size:16pt;font-weight:700;margin-bottom:3pt}
.co-detail{font-size:8pt;line-height:1.5;color:#444}
.doc-type{font-size:14pt;font-weight:700;margin-bottom:4pt}
.doc-no{font-size:9pt;line-height:1.6}
.doc-title{font-size:20pt;font-weight:700;text-align:center;border:2px solid #111;padding:4mm 0;margin-bottom:6mm;letter-spacing:2pt}
.parties{display:table;width:100%;margin-bottom:6mm;border-collapse:collapse}
.pl,.pr{display:table-cell;vertical-align:top;width:50%;border:1px solid #ccc;padding:4mm}
.pl{border-right:none}
.p-label{font-size:8pt;font-weight:700;color:#555;margin-bottom:3pt;text-transform:uppercase}
.p-name{font-size:11pt;font-weight:700;margin-bottom:3pt}
.p-detail{font-size:8.5pt;line-height:1.5;color:#444}
table.items{width:100%;border-collapse:collapse;margin-bottom:6mm}
table.items th{background:#333;color:#fff;padding:2.5mm 3mm;font-size:8.5pt;font-weight:700;text-align:center}
table.items td{border:1px solid #ccc;padding:2mm 3mm;font-size:8.5pt;vertical-align:middle}
table.items tbody tr:nth-child(even){background:#f9f9f9}
.tc{text-align:center}.tr{text-align:right}
.total-section{width:45%;float:right;margin-bottom:4mm}
.t-row{display:table;width:100%;border-bottom:1px solid #eee;padding:1.5mm 0}
.t-label{display:table-cell;text-align:right;padding-right:4mm;font-size:9pt;color:#444;width:58%}
.t-val{display:table-cell;text-align:right;font-size:9pt;font-weight:700;width:42%}
.t-grand .t-label,.t-grand .t-val{font-size:13pt;font-weight:700;border-top:2.5px solid #111;padding-top:2mm;border-bottom:none}
.words-box{font-size:9.5pt;font-weight:700;text-align:center;border:1px solid #111;padding:3mm;background:#f9f9f9;margin-bottom:6mm}
.ftr{display:table;width:100%;border-top:1.5px solid #111;padding-top:5mm;margin-top:6mm}
.fl,.fr{display:table-cell;vertical-align:top;width:50%}
.note{font-size:8.5pt;line-height:1.6}
.sig-box{text-align:center;padding-top:2mm}
.sig-line{border-bottom:1px dashed #aaa;width:55mm;margin:12mm auto 3mm}
.sig-label{font-size:8pt;color:#555}
.clearfix:after{content:"";display:table;clear:both}
</style></head><body><div class="page">
<div class="hdr">
  <div class="hl">
    <div class="co-name">${esc_(d.companyName||'')}</div>
    <div class="co-detail">${esc_(d.companyAddress||'')}<br>
    ${d.companyTaxId?'เลขประจำตัวผู้เสียภาษี: '+esc_(d.companyTaxId):''}${d.companyBranch?' &nbsp;|&nbsp; '+esc_(d.companyBranch):''}<br>
    โทร: ${esc_(d.companyPhone||'-')} &nbsp;&nbsp; อีเมล: ${esc_(d.companyEmail||'-')}</div>
  </div>
  <div class="hr">
    <div class="doc-type">${esc_(d.docTypeName)}</div>
    <div class="doc-no">เลขที่: <strong>${esc_(d.docNo)}</strong><br>วันที่: ${esc_(d.docDate)}<br>${esc_(cfg.el)}: ${esc_(cfg.ev)}</div>
  </div>
</div>
<div class="doc-title">${esc_(d.docTypeName)}</div>
<div class="parties">
  <div class="pl">
    <div class="p-label">${esc_(cfg.rl)}</div>
    <div class="p-name">${esc_(rName)}</div>
    <div class="p-detail">${rAddr?esc_(rAddr)+'<br>':''}${rTax?'เลขประจำตัวผู้เสียภาษี: '+esc_(rTax)+'<br>':''}${rBranch?esc_(rBranch)+'<br>':''}${rPhone?'โทร: '+esc_(rPhone):''}</div>
  </div>
  <div class="pr">
    <div class="p-label">รายละเอียดเอกสาร</div>
    <div class="p-detail" style="line-height:1.8">วันที่ออกเอกสาร: ${esc_(d.docDate)}<br>${esc_(cfg.el)}: ${esc_(cfg.ev)}<br>ผู้ติดต่อ: ${esc_(d.contactPerson||'-')}</div>
  </div>
</div>
<table class="items">
  <thead><tr><th style="width:5%">ลำดับ</th><th style="width:38%">รายการ</th><th style="width:12%">จำนวน</th><th style="width:16%">ราคา/หน่วย</th><th style="width:13%">ส่วนลด</th><th style="width:16%">รวม (บาท)</th></tr></thead>
  <tbody>${rows}</tbody>
</table>
<div class="clearfix">
  <div class="total-section">
    <div class="t-row"><div class="t-label">รวมเป็นเงิน</div><div class="t-val">${sub.toLocaleString('th-TH',{minimumFractionDigits:2})}</div></div>
    ${d.includeVat?'<div class="t-row"><div class="t-label">ภาษีมูลค่าเพิ่ม 7%</div><div class="t-val">'+vatAmt.toLocaleString('th-TH',{minimumFractionDigits:2})+'</div></div>':''}
    <div class="t-row t-grand"><div class="t-label">รวมทั้งสิ้น</div><div class="t-val">${total.toLocaleString('th-TH',{minimumFractionDigits:2})}</div></div>
  </div>
</div>
<div style="clear:both;height:4mm"></div>
<div class="words-box">( ${amountInWords} )</div>
<div class="ftr">
  <div class="fl"><div class="note"><strong>หมายเหตุ:</strong><br>${esc_(cfg.note)}</div></div>
  <div class="fr"><div class="sig-box"><div class="sig-line"></div><div class="sig-label">ลงชื่อ .................................................. / วันที่ .............</div><div class="sig-label" style="margin-top:2mm">(${esc_(cfg.sig)})</div><div class="sig-label" style="margin-top:2mm">${esc_(d.companyName||'')}</div></div></div>
</div>
</div></body></html>`;
}

// ============================================================
// LINE Messaging API webhook — AI auto-reply bot
// Keyword shortcuts (same as rich menu) answer first; anything
// else is answered by Gemini (free tier), scoped to FAQ about
// this business/app only.
// ============================================================

const LIFF_ID_FOR_BOT = '2010618788-8xUdkoqR';

const KEYWORD_REPLIES_ = [
  { keywords: ['เสนอราคา', 'ใบเสนอราคา', 'quotation'], type: 'quotation', label: 'สร้างใบเสนอราคา' },
  { keywords: ['po', 'ใบสั่งซื้อ', 'สั่งซื้อ'],           type: 'po',        label: 'สร้างใบสั่งซื้อ (PO)' },
  { keywords: ['invoice', 'ใบแจ้งหนี้', 'แจ้งหนี้'],       type: 'invoice',   label: 'สร้างใบแจ้งหนี้ (Invoice)' },
  { keywords: ['ใบเสร็จ', 'receipt'],                      type: 'receipt',   label: 'สร้างใบเสร็จรับเงิน' },
  { keywords: ['ใบส่งของ', 'delivery'],                    type: 'delivery',  label: 'สร้างใบส่งของ' },
  { keywords: ['ใบวางบิล', 'billing'],                     type: 'billing',   label: 'สร้างใบวางบิล' }
];

function findKeywordReply_(text) {
  const lower = (text || '').toLowerCase();
  return KEYWORD_REPLIES_.find(function(k) {
    return k.keywords.some(function(kw) { return lower.indexOf(kw.toLowerCase()) > -1; });
  }) || null;
}

// Lists every document type with its direct deep link — used when the AI
// can't answer (e.g. Gemini hiccup) so the customer still gets something
// immediately useful instead of a dead-end error.
function buildDocumentMenuMessage_() {
  const lines = KEYWORD_REPLIES_.map(function(k) {
    return '📄 ' + k.label + '\nhttps://liff.line.me/' + LIFF_ID_FOR_BOT + '?type=' + k.type;
  });
  return 'หนูสามารถทำเอกสารได้ดังนี้ค่ะ 🌸\n\n' + lines.join('\n\n');
}

// Entry point LINE calls (routed here from doPost when body.events exists)
function handleLineWebhook_(body) {
  const events = body.events || [];
  const cache = CacheService.getScriptCache();
  events.forEach(function(ev) {
    try {
      if (ev.type !== 'message' || ev.message.type !== 'text') return;

      // LINE resends the webhook if our endpoint doesn't respond within its
      // timeout — a real risk here since a single message can trigger a
      // Gemini API call plus several Sheets reads/writes. Without this
      // guard, a slow-but-otherwise-successful request gets replayed and
      // silently double-logs a pending subscription, double-pushes the
      // admin notification, or double-consumes a quota unit.
      const dedupeKey = 'line_evt_' + (ev.webhookEventId || ev.message.id);
      if (cache.get(dedupeKey)) return;
      cache.put(dedupeKey, '1', 21600); // 6h — the max CacheService TTL, comfortably past any retry window

      const text = ev.message.text;
      const replyToken = ev.replyToken;
      const lineUserId = ev.source && ev.source.userId;

      const adminApproveReply = tryHandleAdminApprove_(text, lineUserId);
      const match = findKeywordReply_(text);
      let replyMessages;

      // [Fix] Ensure priority routing for subscribe/approve flows
      if (adminApproveReply) {
        replyMessages = adminApproveReply;
      } else if (isWhoAmIQuery_(text)) {
        replyMessages = [{ type: 'text', text: 'LINE User ID ของคุณคือ:\n' + lineUserId }];
      } else if (isSubscribeIntentQuery_(text)) {
        // [Fix] Check subscribe intent BEFORE general document menu to catch "สมัคร" correctly
        replyMessages = buildSubscribeReply_(text, lineUserId);
      } else if (match) {
        replyMessages = [{ type: 'text', text: '👉 ' + match.label + '\nกดลิงก์นี้เพื่อเริ่มได้เลยค่ะ 💕\nhttps://liff.line.me/' + LIFF_ID_FOR_BOT + '?type=' + match.type }];
      } else if (isMenuQuery_(text)) {
        replyMessages = [{ type: 'text', text: buildDocumentMenuMessage_() }];
      } else if (isPackageInfoQuery_(text)) {
        replyMessages = [{ type: 'text', text: buildPackageInfoMessage_() }];
      } else if (isTaxFeatureQuery_(text)) {
        replyMessages = [{ type: 'text', text: buildTaxFeatureMessage_(lineUserId) }];
      } else {
        replyMessages = [{ type: 'text', text: handleAiOrUpsell_(text, lineUserId) }];
      }

      replyMessagesToLine_(replyToken, replyMessages);
    } catch (err) {
      // Swallow errors per-event so one bad event doesn't break the whole webhook batch.
    }
  });
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function isMenuQuery_(text) {
  const lower = (text || '').toLowerCase();
  return ['เมนู', 'menu', 'ทำอะไรได้บ้าง', 'ช่วยอะไรได้บ้าง'].some(function(kw) {
    return lower.indexOf(kw) > -1;
  });
}

function isPackageInfoQuery_(text) {
  const lower = (text || '').toLowerCase();
  return ['แพ็กเกจ', 'แพคเกจ', 'แพ็คเกจ', 'สมัครสมาชิก', 'ราคาแพ็ก', 'package'].some(function(kw) {
    return lower.indexOf(kw) > -1;
  });
}

// This is the fix for the reported bug: a customer typing "สมัครที่ไหน"
// didn't match any keyword above (it has no full "แพ็กเกจ"/"สมัครสมาชิก"
// substring) and fell straight through to the general AI fallback, which
// just repeated the document-menu message instead of telling them where to
// actually sign up. These broader subscribe/pay phrases catch that case.
function isSubscribeIntentQuery_(text) {
  const lower = (text || '').trim().toLowerCase();
  // [Fix] Add "สมัคร" as a standalone keyword and cover "สมัครแพ็คโปร" etc.
  // We use regex or a more flexible check to ensure "สมัคร" alone works.
  const keywords = [
    'สมัคร', 'สมัครที่ไหน', 'สมัครยังไง', 'สมัครแบบไหน', 'จะสมัคร', 'อยากสมัคร', 'สนใจสมัคร',
    'อัปเกรด', 'อัพเกรด', 'อยากอัปเกรด',
    'จ่ายเงิน', 'จ่ายยังไง', 'จ่ายตังยังไง', 'ชำระเงิน', 'โอนเงิน', 'โอนตัง', 'payment'
  ];
  return keywords.some(function(kw) { return lower.indexOf(kw) > -1; });
}

// Loose match for a package the customer names/numbers when replying to the
// package list (e.g. "เอาแพ็ค 2", "อยากได้แพ็คโปร", "สมัครพรีเมียม"). Digit
// choices only count right after "แพ็ค/แพ็ก/แพค" specifically, so a stray
// number elsewhere in the message (a date, a quantity) can't misfire.
function resolvePackageChoice_(text) {
  const lower = (text || '').toLowerCase();
  // [Fix] Better matching for package names to avoid missing "สมัครแพ็คโปร"
  if (lower.indexOf('พรีเมียม') > -1 || lower.indexOf('premium') > -1 || lower.indexOf('249') > -1) return PACKAGES.premium249;
  if (lower.indexOf('โปร') > -1 || lower.indexOf('pro') > -1 || lower.indexOf('149') > -1) return PACKAGES.pro149;
  if (lower.indexOf('เริ่มต้น') > -1 || lower.indexOf('starter') > -1 || lower.indexOf('59') > -1) return PACKAGES.starter59;
  
  const m = lower.match(/แพ็?[คก]\s*([123])/);
  if (m) {
    if (m[1] === '3') return PACKAGES.premium249;
    if (m[1] === '2') return PACKAGES.pro149;
    if (m[1] === '1') return PACKAGES.starter59;
  }
  return null;
}

function isTaxFeatureQuery_(text) {
  const lower = (text || '').toLowerCase();
  return ['ภาษี', 'ยื่นภาษี', 'ทำบัญชี', 'บัญชี'].some(function(kw) {
    return lower.indexOf(kw) > -1;
  });
}

function buildPackageInfoMessage_() {
  return '📦 แพ็กเกจของ ง่าย ผู้ช่วยทำเอกสาร\n\n' +
    '🆓 ทดลองใช้ฟรี\n' +
    '• สร้างเอกสารได้ 5 ครั้ง/เดือน ไม่มีค่าใช้จ่าย\n\n' +
    '1️⃣ แพ็คเริ่มต้น 59 บาท/เดือน\n' +
    '• สร้างเอกสารได้ 30 ครั้ง/เดือน\n\n' +
    '2️⃣ แพ็คโปร 149 บาท/เดือน\n' +
    '• สร้างเอกสารได้ไม่จำกัด\n\n' +
    '3️⃣ แพ็คพรีเมียม 249 บาท/เดือน\n' +
    '• สร้างเอกสารได้ไม่จำกัด\n' +
    '• ระบบทำบัญชี/ยื่นภาษีรายเดือน (เร็วๆ นี้)\n\n' +
    '(น้องแชทตอบคำถามได้ฟรีทุกแพ็กเลยค่ะ 💕)\n\n' +
    'สนใจสมัคร/อัปเกรดแพ็กเกจ ทักแอดมินได้เลยนะคะ 🙏💕\n\n' +
    adminContactBlock_();
}

function buildTaxFeatureMessage_(lineUserId) {
  const sub = lineUserId ? getUserSubscription_(lineUserId) : null;
  if (sub && sub.pkg.hasTax) {
    return '🧾 ระบบทำบัญชี/ยื่นภาษีรายเดือน กำลังพัฒนาอยู่ค่ะ จะเปิดให้ใช้งานเร็วๆ นี้นะคะ (ท่านมีสิทธิ์ใช้งานฟีเจอร์นี้อยู่แล้วในแพ็ก ' + sub.pkg.name + ')';
  }
  return '🧾 ฟีเจอร์ระบบบัญชี/ยื่นภาษี อยู่ในแพ็คพรีเมียม 249 บาท/เดือนเท่านั้นนะคะ พิมพ์ "แพ็กเกจ" เพื่อดูรายละเอียดและอัปเกรดได้เลยค่ะ';
}

// Only package premium249 (hasAI) gets real AI answers; everyone else gets an upsell.
// AI FAQ assistant answers everyone immediately — first point of contact
// for any message that isn't a keyword shortcut (e.g. a customer just
// saying "สวัสดี" gets a real, immediate reply, not a paywall message).
// Not gated by package: this is basic customer service, not a premium perk.
function handleAiOrUpsell_(text, lineUserId) {
  return askGemini_(text);
}

// Restricted-scope FAQ system prompt — only answers questions about this
// business / how to use the CHUAY document app. Anything else gets a
// polite redirect instead of a free-ranging AI conversation.
const BOT_SYSTEM_PROMPT_ =
  'คุณคือผู้ช่วยตอบคำถามของ "ง่าย ผู้ช่วยทำเอกสาร" เป็นผู้หญิงตอบแบบน่ารักเป็นกันเอง ซึ่งเป็นแอปสร้างเอกสารธุรกิจผ่าน LINE ' +
  '(ใบเสนอราคา, ใบสั่งซื้อ PO, ใบแจ้งหนี้ Invoice, ใบเสร็จรับเงิน, ใบส่งของ, ใบวางบิล) ' +
  'คุณคือด่านแรกที่ลูกค้าทักมาคุยด้วย — ถ้าลูกค้าทักทายเฉยๆ เช่น "สวัสดี", "หวัดดี", "Hello" ให้ทักทายกลับอย่างอบอุ่นเป็นกันเอง แนะนำตัวสั้นๆ ว่าช่วยอะไรได้บ้าง แล้วถามว่าวันนี้อยากให้ช่วยเรื่องอะไร ' +
  'ตอบเฉพาะคำถามเกี่ยวกับวิธีใช้งานแอปนี้ ฟีเจอร์ต่างๆ และบริการของบริษัทเท่านั้น ' +
  'ตอบสั้น กระชับ เป็นกันเอง เป็นภาษาไทย ไม่เกิน 3-4 ประโยค ' +
  'ถ้าคำถามอยู่นอกขอบเขตนี้โดยสิ้นเชิง (เช่น เรื่องทั่วไปที่ไม่เกี่ยวกับแอปหรือบริการเลย) ' +
  'ให้ตอบอย่างสุภาพว่าไม่สามารถช่วยเรื่องนี้ได้ และแนะนำให้พิมพ์ถามเกี่ยวกับการใช้งานแอปแทน ' +
  'อย่าสร้างข้อมูลราคา/เงื่อนไขที่ไม่แน่ใจขึ้นมาเอง หากไม่ทราบให้แนะนำให้ติดต่อทีมงานโดยตรง';

function askGemini_(userText) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return 'ขออภัยค่ะ ระบบ AI ยังไม่ได้ตั้งค่า รบกวนติดต่อผู้ดูแลระบบนะคะ 🙏';

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent';
  const payload = {
    system_instruction: { parts: [{ text: BOT_SYSTEM_PROMPT_ }] },
    contents: [{ role: 'user', parts: [{ text: userText }] }]
  };

  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const data = JSON.parse(res.getContentText());
    if (data.candidates && data.candidates[0] && data.candidates[0].content) {
      return data.candidates[0].content.parts[0].text.trim();
    }
    return buildDocumentMenuMessage_();
  } catch (err) {
    return buildDocumentMenuMessage_();
  }
}

function replyToLine_(replyToken, text) {
  replyMessagesToLine_(replyToken, [{ type: 'text', text: text }]);
}

// Like replyToLine_ but takes an array of LINE message objects, so a single
// reply can include e.g. a text message plus a QR-code image message (used
// by the subscribe flow below — LINE's reply API supports up to 5 messages
// per reply token).
function replyMessagesToLine_(replyToken, messages) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) { console.log('LINE_CHANNEL_ACCESS_TOKEN is missing!'); return; }
  try {
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ replyToken: replyToken, messages: messages }),
      muteHttpExceptions: true
    });
    console.log('LINE reply status:', res.getResponseCode(), res.getContentText());
  } catch (err) {
    console.log('replyMessagesToLine_ error:', err);
  }
}

// Sends a message outside of a reply-token context (e.g. notifying the
// admin the instant a new subscribe request comes in, or telling a
// customer their package just got activated) — used by the admin-approve
// flow so neither side has to wait on a reply-token from their own message.
function pushToLine_(toUserId, messages) {
  try {
    const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
    if (!token || !toUserId) return;
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ to: toUserId, messages: messages }),
      muteHttpExceptions: true
    });
    console.log('LINE push status:', res.getResponseCode(), res.getContentText());
  } catch (err) {
    console.log('pushToLine_ error:', err);
  }
}
