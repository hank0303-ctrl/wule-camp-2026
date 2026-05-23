const SPREADSHEET_ID = '1_lTR7lCP_QS2fg3-b2R26Z8Zpmw8LFTd2169JF5Ranw';
const REGISTRATION_SHEET = '報名資料';
const INSURANCE_SHEET = '保險資料';
const TIMEZONE = 'Asia/Taipei';

const INSURANCE_HEADERS = [
  '送出時間',
  '學員姓名',
  '家長電話後4碼',
  '報名期數',
  '總費用',
  '需付訂金',
  '訂金狀態',
  '已付金額',
  '剩餘尾款',
  '尾款期限',
  '小朋友身分證字號',
  '監護人姓名',
  '聯絡地址',
  '備註',
  '資料確認'
];

function doGet(e) {
  try {
    const action = String(e.parameter.action || 'ping');

    if (action === 'lookup') {
      const result = lookupRegistration_(e.parameter.name || '', e.parameter.phoneLast4 || '');
      return json_(e, result);
    }

    return json_(e, {
      ok: true,
      service: 'wule-insurance-api',
      actions: ['lookup', 'submitInsurance']
    });
  } catch (err) {
    return json_(e, { ok: false, message: err.message });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse((e.postData && e.postData.contents) || '{}');

    if (data.action !== 'submitInsurance') {
      throw new Error('不支援的動作');
    }

    const lookup = lookupRegistration_(data.studentName || '', data.phoneLast4 || '');
    if (!lookup.ok || !lookup.registration) {
      throw new Error('查無報名資料，請重新確認姓名與電話後 4 碼');
    }

    const reg = lookup.registration;
    const sheet = getInsuranceSheet_();
    sheet.appendRow([
      formatDateTime_(new Date()),
      reg.name,
      String(data.phoneLast4 || ''),
      reg.sessions,
      reg.totalFeeText,
      reg.depositText,
      reg.depositStatusText,
      '$' + reg.paidAmount.toLocaleString(),
      '$' + reg.remainingAmount.toLocaleString(),
      reg.paymentDeadline,
      data.childId || '',
      data.guardian || '',
      data.address || '',
      data.note || '',
      data.confirmed ? '已確認' : ''
    ]);

    return json_(null, { ok: true });
  } catch (err) {
    return json_(null, { ok: false, message: err.message });
  }
}

function authorize() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ss.getSheetByName(REGISTRATION_SHEET).getLastRow();
  getInsuranceSheet_();
  return 'ok';
}

function lookupRegistration_(name, phoneLast4) {
  const targetName = normalizeText_(name);
  const targetPhone = String(phoneLast4 || '').replace(/\D/g, '').slice(-4);
  if (!targetName || targetPhone.length !== 4) {
    return { ok: false, message: '請輸入小朋友姓名與家長電話後 4 碼' };
  }

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(REGISTRATION_SHEET);
  if (!sheet) throw new Error('找不到報名資料分頁');

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return { ok: false, message: '目前沒有報名資料' };

  const headers = values[0].map(String);
  const rows = values.slice(1);
  const matches = rows
    .map((row, index) => rowToObject_(headers, row, index + 2))
    .filter(row => normalizeText_(row['學員姓名']) === targetName)
    .filter(row => String(row['聯絡電話'] || '').replace(/\D/g, '').slice(-4) === targetPhone);

  if (matches.length === 0) {
    return { ok: false, message: '查無資料，請確認姓名是否與報名時相同、電話後 4 碼是否正確' };
  }

  if (matches.length > 1) {
    return { ok: false, message: '查到多筆相同資料，請聯繫武樂協助確認' };
  }

  return { ok: true, registration: buildRegistrationView_(matches[0]) };
}

function buildRegistrationView_(row) {
  const baseTotal = parseMoney_(row['總費用']);
  const deposit = parseMoney_(row['需付訂金']);
  const adjustmentLabel = row['費用調整說明（給家長看）'] || '';
  const adjustmentAmount = parseMoney_(row['費用調整金額（+加收/-優惠）']);
  const adjustedTotalOverride = hasValue_(row['調整後總費用（可留空）'])
    ? parseMoney_(row['調整後總費用（可留空）'])
    : 0;
  const total = adjustedTotalOverride > 0 ? adjustedTotalOverride : Math.max(baseTotal + adjustmentAmount, 0);
  const status = String(row['訂金狀態'] || '待繳');
  const isPaid = status === '已繳';
  const paidOverride = hasValue_(row['已收款金額（可留空）'])
    ? parseMoney_(row['已收款金額（可留空）'])
    : null;
  const paid = paidOverride !== null ? paidOverride : (isPaid ? deposit : 0);
  const remaining = Math.max(total - paid, 0);

  return {
    rowNum: row.rowNum,
    name: row['學員姓名'] || '',
    gender: row['性別'] || '',
    age: row['年齡'] || '',
    birthday: row['出生年月日'] || '',
    sessions: row['報名期數'] || '',
    baseTotalFee: baseTotal,
    baseTotalFeeText: formatMoney_(baseTotal),
    totalFee: total,
    totalFeeText: formatMoney_(total),
    deposit: deposit,
    depositText: formatMoney_(deposit),
    adjustmentLabel: adjustmentLabel,
    adjustmentAmount: adjustmentAmount,
    adjustmentAmountText: formatSignedMoney_(adjustmentAmount),
    adjustedTotalOverride: adjustedTotalOverride,
    paymentMethod: row['付款方式'] || '',
    depositStatusText: status,
    paidAmount: paid,
    paidAmountText: formatMoney_(paid),
    remainingAmount: remaining,
    remainingAmountText: formatMoney_(remaining),
    paymentDeadline: '2026/6/23 以前',
    parentName: row['家長姓名'] || '',
    phoneMasked: maskPhone_(row['聯絡電話'] || ''),
    aftercare: row['延托服務'] || '未填',
    equipment: row['裝備'] || '未填'
  };
}

function getInsuranceSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(INSURANCE_SHEET);
  if (!sheet) sheet = ss.insertSheet(INSURANCE_SHEET);

  const firstRow = sheet.getRange(1, 1, 1, INSURANCE_HEADERS.length).getValues()[0];
  const needsHeader = firstRow.every(cell => cell === '');
  if (needsHeader) {
    sheet.getRange(1, 1, 1, INSURANCE_HEADERS.length).setValues([INSURANCE_HEADERS]);
    sheet.getRange(1, 1, 1, INSURANCE_HEADERS.length)
      .setBackground('#1A1610')
      .setFontColor('#E8C97A')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function rowToObject_(headers, row, rowNum) {
  const obj = { rowNum };
  headers.forEach((key, i) => obj[key] = row[i] || '');
  return obj;
}

function normalizeText_(value) {
  return String(value || '').replace(/\s+/g, '').trim();
}

function hasValue_(value) {
  return String(value || '').trim() !== '';
}

function parseMoney_(value) {
  const num = Number(String(value || '').replace(/[^0-9.-]/g, ''));
  return Number.isFinite(num) ? num : 0;
}

function formatMoney_(value) {
  return '$' + Number(value || 0).toLocaleString();
}

function formatSignedMoney_(value) {
  const amount = Number(value || 0);
  if (amount > 0) return '+$' + amount.toLocaleString();
  if (amount < 0) return '-$' + Math.abs(amount).toLocaleString();
  return '$0';
}

function maskPhone_(value) {
  const phone = String(value || '').replace(/\D/g, '');
  if (phone.length < 4) return '';
  return phone.slice(0, 4) + '***' + phone.slice(-3);
}

function formatDateTime_(date) {
  return Utilities.formatDate(date, TIMEZONE, 'yyyy/MM/dd HH:mm:ss');
}

function json_(e, payload) {
  const output = JSON.stringify(payload);
  const callback = e && e.parameter && e.parameter.callback;
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + output + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}
