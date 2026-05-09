// ============================================================
// SEED CARD GAS スクリプト v2
// ============================================================

var CONFIG = {
  STRIPE_URL_INITIAL: 'https://buy.stripe.com/fZu6oH6Gf1pw8iN7R8es008',
  STRIPE_URL_RENEWAL: 'https://buy.stripe.com/9B63cvc0z7NU0QlfjAes009',
  PORTAL_URL:         'https://myseedcard.japan-sperm.com/portal.html',
  CARD_BASE_URL:      'https://myseedcard.japan-sperm.com/',
  FROM_NAME:          '一般社団法人 日本精子協会\n〒104-0061 東京都中央区銀座1丁目12番4号 N&E BLD.6F\njapan.sperm.association@gmail.com'
};

// 新規セットアップ時に一度だけ実行 → 管理者パスワード設定
function setupAdminKey() {
  var password = 'ここにパスワードを入力してください';
  PropertiesService.getScriptProperties().setProperty('ADMIN_KEY', password);
  Logger.log('管理者パスワードを設定しました: ' + password);
}

// ---- シート取得 ----

function getInitialSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('初回申請');
}
function getRenewalSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('更新申請');
}

// 新規シートのヘッダーを初期化（新規セットアップ時に自動実行）
function ensureSheetHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var initSheet = ss.getSheetByName('初回申請');
  if (!initSheet) initSheet = ss.insertSheet('初回申請');
  if (initSheet.getLastRow() === 0) {
    initSheet.getRange(1, 1, 1, 22).setValues([[
      'タイムスタンプ','メールアドレス','氏名（漢字）','氏名（ローマ字）','フリガナ','生年月日',
      '電話番号','郵便番号','住所',
      '精液量（mL）','精子濃度（万/mL）','総精子数（万）','正常形態率（%）',
      '精液検査証（写真）','審査結果',
      '会員番号','認定日','認定期限','カードURL','支払い状況','発送状況','更新通知'
    ]]);
    Logger.log('初回申請シートのヘッダーを作成しました');
  }

  var renSheet = ss.getSheetByName('更新申請');
  if (!renSheet) renSheet = ss.insertSheet('更新申請');
  if (renSheet.getLastRow() === 0) {
    renSheet.getRange(1, 1, 1, 7).setValues([[
      'タイムスタンプ','メールアドレス','氏名（漢字）','会員番号',
      '精液検査証（写真）','審査結果','元行番号'
    ]]);
    Logger.log('更新申請シートのヘッダーを作成しました');
  }
}

// ---- 列検出 ----

function getCols(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var c = {};
  for (var i = 0; i < headers.length; i++) {
    var s = String(headers[i]).trim();
    if (!c.NAME_KANJI  && (s.indexOf('漢字') >= 0 || s === '氏名' || s === 'お名前' || s === '名前' || s.indexOf('お名前') >= 0)) { c.NAME_KANJI = i + 1; }
    if (!c.NAME_ROMAJI && (s.indexOf('ローマ字') >= 0 || s.indexOf('英語') >= 0))            { c.NAME_ROMAJI = i + 1; }
    if (!c.NAME_KANA   && (s.indexOf('フリガナ') >= 0 || s.indexOf('ふりがな') >= 0 || s.indexOf('カタカナ') >= 0)) { c.NAME_KANA = i + 1; }
    if (!c.DOB         && s.indexOf('生年月日') >= 0)                                         { c.DOB = i + 1; }
    if (!c.TEL         && (s.indexOf('電話') >= 0 || s.toLowerCase() === 'tel'))              { c.TEL = i + 1; }
    // メール列：複数ある場合は後の列を優先（Googleフォーム自動収集列B列より、F列などを優先するため）
    if (s.length <= 50 && (s.indexOf('メール') >= 0 || s.indexOf('mail') >= 0 || s.indexOf('Mail') >= 0)) { c.EMAIL = i + 1; }
    if (!c.POSTAL      && s.indexOf('郵便') >= 0)                                             { c.POSTAL = i + 1; }
    if (!c.ADDRESS     && s.indexOf('住所') >= 0)                                             { c.ADDRESS = i + 1; }
    if (!c.VOLUME      && s.indexOf('精液量') >= 0)                                           { c.VOLUME = i + 1; }
    if (!c.CONCENTRATION && s.indexOf('精子濃度') >= 0)                                       { c.CONCENTRATION = i + 1; }
    if (!c.TOTAL       && s.indexOf('総精子数') >= 0)                                         { c.TOTAL = i + 1; }
    if (!c.MORPHOLOGY  && s.indexOf('正常形態率') >= 0)                                       { c.MORPHOLOGY = i + 1; }
    if (!c.PHOTO       && s.indexOf('写真') >= 0)                                             { c.PHOTO = i + 1; }
    if (!c.RESULT      && (s.indexOf('合否') >= 0 || s.indexOf('審査結果') >= 0))            { c.RESULT = i + 1; }
    if (!c.MEMBER_NO   && s.indexOf('会員番号') >= 0)                                         { c.MEMBER_NO = i + 1; }
    if (!c.CERT_DATE   && s.indexOf('認定日') >= 0 && s.indexOf('期限') < 0)                 { c.CERT_DATE = i + 1; }
    if (!c.EXPIRY_DATE && s.indexOf('認定期限') >= 0)                                         { c.EXPIRY_DATE = i + 1; }
    if (!c.CARD_URL    && s.indexOf('カードURL') >= 0)                                        { c.CARD_URL = i + 1; }
    if (!c.PAYMENT     && s.indexOf('支払い') >= 0)                                           { c.PAYMENT = i + 1; }
    if (!c.SHIPPING    && s.indexOf('発送') >= 0 && s.indexOf('住所') < 0)                   { c.SHIPPING = i + 1; }
    if (!c.RENEWAL_NOTICE && s.indexOf('更新通知') >= 0)                                      { c.RENEWAL_NOTICE = i + 1; }
    if (!c.ORIGINAL_ROW   && s.indexOf('元行番号') >= 0)                                      { c.ORIGINAL_ROW = i + 1; }
    if (!c.TIMESTAMP   && (s.indexOf('タイムスタンプ') >= 0 || s.indexOf('Timestamp') >= 0)) { c.TIMESTAMP = i + 1; }
  }
  // タイムスタンプが見つからなければ列1とみなす
  if (!c.TIMESTAMP) c.TIMESTAMP = 1;
  return c;
}

// ---- 行データ取得 ----

function getRowData(sheet, row, c) {
  var nameKanji     = c.NAME_KANJI    ? sheet.getRange(row, c.NAME_KANJI).getValue()    : '';
  var nameRomaji    = c.NAME_ROMAJI   ? sheet.getRange(row, c.NAME_ROMAJI).getValue()   : '';
  var email         = c.EMAIL         ? sheet.getRange(row, c.EMAIL).getValue()         : '';
  var postal        = c.POSTAL        ? sheet.getRange(row, c.POSTAL).getValue()        : '';
  var address       = c.ADDRESS       ? sheet.getRange(row, c.ADDRESS).getValue()       : '';
  var volume        = c.VOLUME        ? parseFloat(sheet.getRange(row, c.VOLUME).getValue()) : 0;
  var concentration = c.CONCENTRATION ? parseFloat(sheet.getRange(row, c.CONCENTRATION).getValue()) : 0;
  var total         = c.TOTAL         ? parseFloat(sheet.getRange(row, c.TOTAL).getValue()) : 0;
  var morphology    = c.MORPHOLOGY    ? parseFloat(sheet.getRange(row, c.MORPHOLOGY).getValue()) : 0;
  var photo         = c.PHOTO         ? sheet.getRange(row, c.PHOTO).getValue()         : '';
  var memberNo      = c.MEMBER_NO     ? sheet.getRange(row, c.MEMBER_NO).getValue()     : '';
  return {
    nameKanji:     String(nameKanji     || '').trim(),
    nameRomaji:    String(nameRomaji    || '').trim(),
    email:         String(email         || '').trim(),
    postal:        String(postal        || '').trim(),
    address:       String(address       || '').trim(),
    volume:        isNaN(volume)        ? 0 : volume,
    concentration: isNaN(concentration) ? 0 : concentration,
    total:         isNaN(total)         ? 0 : total,
    morphology:    isNaN(morphology)    ? 0 : morphology,
    photo:         String(photo         || '').trim(),
    memberNo:      String(memberNo      || '').trim()
  };
}

// ---- フォーム送信トリガー（Googleフォーム経由の場合のみ使用） ----

function onFormSubmit(e) {
  var sheet = e.range.getSheet();
  var row   = e.range.getRow();
  if (sheet.getName() === '初回申請') {
    handleInitialSubmit(sheet, row);
  } else if (sheet.getName() === '更新申請') {
    handleRenewalSubmit(sheet, row);
  }
}

function handleInitialSubmit(sheet, row) {
  var c    = getCols(sheet);
  var data = getRowData(sheet, row, c);
  if (c.RESULT) { sheet.getRange(row, c.RESULT).setValue('審査中'); }
  sendStaffNotification(data);
  sendAcknowledgmentEmail(data);
}

function handleRenewalSubmit(sheet, row) {
  var c        = getCols(sheet);
  var data     = getRowData(sheet, row, c);
  var originalRow = findOriginalRow(data.memberNo);
  if (originalRow > 0) {
    var initSheet = getInitialSheet(); var ic = getCols(initSheet);
    if (ic.NAME_KANJI && c.NAME_KANJI) {
      var correctName = String(initSheet.getRange(originalRow, ic.NAME_KANJI).getValue()).trim();
      if (correctName) { sheet.getRange(row, c.NAME_KANJI).setValue(correctName); data.nameKanji = correctName; }
    }
    if (ic.MEMBER_NO && c.MEMBER_NO) {
      var correctNo = String(initSheet.getRange(originalRow, ic.MEMBER_NO).getValue()).trim();
      if (correctNo) { sheet.getRange(row, c.MEMBER_NO).setValue(correctNo); data.memberNo = correctNo; }
    }
  }
  if (c.ORIGINAL_ROW) { sheet.getRange(row, c.ORIGINAL_ROW).setValue(originalRow > 0 ? originalRow : '未登録'); }
  if (c.RESULT) { sheet.getRange(row, c.RESULT).setValue('審査中'); }
  sendRenewalStaffNotification(data, originalRow);
}

function findOriginalRow(memberNoInput) {
  var sheet = getInitialSheet();
  var c     = getCols(sheet);
  if (!c.MEMBER_NO) { return -1; }
  var num = parseInt(String(memberNoInput).replace(/[^0-9]/g, ''), 10);
  if (isNaN(num)) { return -1; }
  for (var i = 2; i <= sheet.getLastRow(); i++) {
    if (parseInt(sheet.getRange(i, c.MEMBER_NO).getValue(), 10) === num) { return i; }
  }
  return -1;
}

// ---- カスタムフォーム（ポータル）からの申請受付 ----

function handleSubmitApply(p) {
  ensureSheetHeaders();
  var sheet = getInitialSheet();
  var c = getCols(sheet);

  var nameKanji   = String(p.nameKanji   || '').trim();
  var nameKana    = String(p.nameKana    || '').trim();
  var nameRomaji  = String(p.nameRomaji  || '').trim();
  var dob         = String(p.dob         || '').trim();
  var email       = String(p.email       || '').trim().toLowerCase();
  var tel         = String(p.tel         || '').trim();
  var zip         = String(p.zip         || '').trim();
  var address     = String(p.address     || '').trim();
  var volume      = parseFloat(p.volume      || '0') || 0;
  var concentration = parseFloat(p.concentration || '0') || 0;
  var total       = parseFloat(p.total       || '0') || 0;
  var morphology  = parseFloat(p.morphology  || '0') || 0;
  var photoData   = String(p.photoData   || '').trim();
  var photoName   = String(p.photoName   || 'photo.jpg').trim();

  if (!nameKanji || !email) { return { ok: false, error: '必須項目が不足しています' }; }
  if (!isValidEmail(email)) { return { ok: false, error: 'メールアドレスが正しくありません' }; }

  // 重複申請チェック
  for (var i = 2; i <= sheet.getLastRow(); i++) {
    var ex = c.EMAIL ? String(sheet.getRange(i, c.EMAIL).getValue()).trim().toLowerCase() : '';
    if (ex === email) { return { ok: false, error: 'このメールアドレスはすでに登録されています' }; }
  }

  // 写真をDriveに保存
  var photoUrl = '';
  if (photoData) {
    try {
      var folder = getPhotoFolder();
      var blob = Utilities.newBlob(Utilities.base64Decode(photoData), 'image/jpeg', photoName);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoUrl = file.getUrl();
    } catch(err) { Logger.log('Photo upload error: ' + err); }
  }

  // スプレッドシートに書き込み
  var newRow = sheet.getLastRow() + 1;
  var now = new Date();
  sheet.getRange(newRow, c.TIMESTAMP).setValue(now);
  if (c.EMAIL)         sheet.getRange(newRow, c.EMAIL).setValue(email);
  if (c.NAME_KANJI)    sheet.getRange(newRow, c.NAME_KANJI).setValue(nameKanji);
  if (c.NAME_KANA)     sheet.getRange(newRow, c.NAME_KANA).setValue(nameKana);
  if (c.NAME_ROMAJI)   sheet.getRange(newRow, c.NAME_ROMAJI).setValue(nameRomaji);
  if (c.DOB)           sheet.getRange(newRow, c.DOB).setValue(dob);
  if (c.TEL)           sheet.getRange(newRow, c.TEL).setValue(tel);
  if (c.POSTAL)        sheet.getRange(newRow, c.POSTAL).setValue(zip);
  if (c.ADDRESS)       sheet.getRange(newRow, c.ADDRESS).setValue(address);
  if (c.VOLUME)        sheet.getRange(newRow, c.VOLUME).setValue(volume);
  if (c.CONCENTRATION) sheet.getRange(newRow, c.CONCENTRATION).setValue(concentration);
  if (c.TOTAL)         sheet.getRange(newRow, c.TOTAL).setValue(total);
  if (c.MORPHOLOGY)    sheet.getRange(newRow, c.MORPHOLOGY).setValue(morphology);
  if (c.PHOTO)         sheet.getRange(newRow, c.PHOTO).setValue(photoUrl);
  if (c.RESULT)        sheet.getRange(newRow, c.RESULT).setValue('審査中');

  // 通知
  sendStaffNotification({ nameKanji:nameKanji, nameRomaji:nameRomaji, email:email, postal:zip, address:address, photo:photoUrl, volume:volume, concentration:concentration, total:total, morphology:morphology });
  sendAcknowledgmentEmail({ nameKanji:nameKanji, email:email });

  return { ok: true };
}

function handleSubmitRenew(p) {
  var email     = String(p.email     || '').trim().toLowerCase();
  var photoData = String(p.photoData || '').trim();
  var photoName = String(p.photoName || 'photo.jpg').trim();

  if (!email) { return { ok: false, error: 'メールアドレスが必要です' }; }

  // 初回申請シートで会員を特定
  var initSheet = getInitialSheet();
  var ic = getCols(initSheet);
  var memberNo = '', nameKanji = '', initRow = -1;
  for (var i = 2; i <= initSheet.getLastRow(); i++) {
    var rowEmail = ic.EMAIL ? String(initSheet.getRange(i, ic.EMAIL).getValue()).trim().toLowerCase() : '';
    if (rowEmail === email) {
      initRow  = i;
      memberNo = ic.MEMBER_NO  ? String(initSheet.getRange(i, ic.MEMBER_NO).getValue()).trim()  : '';
      nameKanji= ic.NAME_KANJI ? String(initSheet.getRange(i, ic.NAME_KANJI).getValue()).trim() : '';
      break;
    }
  }
  if (initRow < 0) { return { ok: false, error: '会員情報が見つかりません' }; }

  // 写真をDriveに保存
  var photoUrl = '';
  if (photoData) {
    try {
      var folder = getPhotoFolder();
      var blob = Utilities.newBlob(Utilities.base64Decode(photoData), 'image/jpeg', photoName);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoUrl = file.getUrl();
    } catch(err) { Logger.log('Photo upload error: ' + err); }
  }

  // 更新申請シートに書き込み
  ensureSheetHeaders();
  var renSheet = getRenewalSheet();
  var rc = getCols(renSheet);
  var newRow = renSheet.getLastRow() + 1;
  renSheet.getRange(newRow, rc.TIMESTAMP).setValue(new Date());
  if (rc.EMAIL)        renSheet.getRange(newRow, rc.EMAIL).setValue(email);
  if (rc.NAME_KANJI)   renSheet.getRange(newRow, rc.NAME_KANJI).setValue(nameKanji);
  if (rc.MEMBER_NO)    renSheet.getRange(newRow, rc.MEMBER_NO).setValue(memberNo);
  if (rc.PHOTO)        renSheet.getRange(newRow, rc.PHOTO).setValue(photoUrl);
  if (rc.RESULT)       renSheet.getRange(newRow, rc.RESULT).setValue('審査中');
  if (rc.ORIGINAL_ROW) renSheet.getRange(newRow, rc.ORIGINAL_ROW).setValue(initRow);

  sendRenewalStaffNotification({ nameKanji:nameKanji, email:email, memberNo:memberNo, photo:photoUrl, volume:0, concentration:0, total:0, morphology:0 }, initRow);

  return { ok: true };
}

// メールアドレスで会員情報を取得（マイページ用）
function handleGetMemberByEmail(email) {
  if (!email) { return { ok: false, error: 'メールアドレスが必要です' }; }

  var sheet = getInitialSheet();
  var c = getCols(sheet);

  for (var i = 2; i <= sheet.getLastRow(); i++) {
    var rowEmail = c.EMAIL ? String(sheet.getRange(i, c.EMAIL).getValue()).trim().toLowerCase() : '';
    if (rowEmail !== email.toLowerCase()) continue;

    var result     = c.RESULT     ? String(sheet.getRange(i, c.RESULT).getValue()).trim()     : '';
    var payment    = c.PAYMENT    ? String(sheet.getRange(i, c.PAYMENT).getValue()).trim()    : '';
    var shipping   = c.SHIPPING   ? String(sheet.getRange(i, c.SHIPPING).getValue()).trim()   : '';
    var expiryDate = c.EXPIRY_DATE? String(sheet.getRange(i, c.EXPIRY_DATE).getValue()).trim(): '';

    // 更新申請シートにより新しい情報があれば上書き
    var renewStatus = '';
    var rSheet = getRenewalSheet(); var rc = getCols(rSheet);
    for (var j = 2; j <= rSheet.getLastRow(); j++) {
      var rEmail = rc.EMAIL ? String(rSheet.getRange(j, rc.EMAIL).getValue()).trim().toLowerCase() : '';
      if (rEmail === email.toLowerCase()) {
        var rResult = rc.RESULT ? String(rSheet.getRange(j, rc.RESULT).getValue()).trim() : '';
        if (!rResult || rResult === '審査中') { renewStatus = '更新審査中'; }
        break;
      }
    }

    var status = renewStatus || deriveStatus(result, payment, shipping, expiryDate);

    return {
      ok: true,
      member: {
        nameKanji:  c.NAME_KANJI  ? String(sheet.getRange(i, c.NAME_KANJI).getValue()).trim()  : '',
        nameKana:   c.NAME_KANA   ? String(sheet.getRange(i, c.NAME_KANA).getValue()).trim()   : '',
        memberNo:   c.MEMBER_NO   ? String(sheet.getRange(i, c.MEMBER_NO).getValue()).trim()   : '',
        expireDate: expiryDate,
        status:     status,
        tel:        c.TEL         ? String(sheet.getRange(i, c.TEL).getValue()).trim()         : '',
        address:    c.ADDRESS     ? String(sheet.getRange(i, c.ADDRESS).getValue()).trim()     : ''
      }
    };
  }
  return { ok: false, error: '会員情報が見つかりません' };
}

// result/payment/shipping → ポータル表示用ステータスに変換
function deriveStatus(result, payment, shipping, expiryDate) {
  if (shipping === '発送済み' || shipping === '発送済') {
    if (expiryDate) {
      try {
        var exp = new Date(expiryDate.replace(/年|月/g, '-').replace('日', ''));
        if (!isNaN(exp.getTime()) && exp < new Date()) { return '期限切れ'; }
      } catch(e) {}
    }
    return '有効';
  }
  if (shipping === '発送待ち') { return '発送待ち'; }
  if (result === '合格' && (payment === '支払い待ち' || payment === '')) { return '支払い待ち'; }
  if (result === '合格' && payment === '支払い済み') { return '発送待ち'; }
  return '審査中';
}

// 写真保存フォルダ取得（なければ自動作成）
function getPhotoFolder() {
  var folderId = PropertiesService.getScriptProperties().getProperty('PHOTO_FOLDER_ID');
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch(e) {}
  }
  var folders = DriveApp.getFoldersByName('SEED CARD 検査証');
  if (folders.hasNext()) { return folders.next(); }
  var folder = DriveApp.createFolder('SEED CARD 検査証');
  PropertiesService.getScriptProperties().setProperty('PHOTO_FOLDER_ID', folder.getId());
  return folder;
}

// ---- 編集トリガー ----

function onEditInstallable(e) {
  var sheet = e.range.getSheet();
  var row   = e.range.getRow();
  if (row < 2) { return; }
  if (sheet.getName() === '初回申請') {
    handleInitialEdit(e, sheet, row);
  } else if (sheet.getName() === '更新申請') {
    handleRenewalEdit(e, sheet, row);
  }
}

function handleInitialEdit(e, sheet, row) {
  var c     = getCols(sheet);
  var col   = e.range.getColumn();
  var value = String(e.range.getValue()).trim();
  var data  = getRowData(sheet, row, c);
  if (c.RESULT && col === c.RESULT) {
    if (value === '合格') {
      if (c.PAYMENT) { sheet.getRange(row, c.PAYMENT).setValue('支払い待ち'); }
      sendPaymentEmail(data, row);
    } else if (value === '不合格') {
      sendRejectionEmail(data);
    }
  }
  if (c.PAYMENT && col === c.PAYMENT && value === '支払い済み') {
    issueCard(sheet, row, c, data);
  }
  if (c.SHIPPING && col === c.SHIPPING && value === '発送済み') {
    var cardUrl = c.CARD_URL ? String(sheet.getRange(row, c.CARD_URL).getValue()).trim() : '';
    sendShippingEmail(data, cardUrl);
  }
}

function handleRenewalEdit(e, sheet, row) {
  var c     = getCols(sheet);
  var col   = e.range.getColumn();
  var value = String(e.range.getValue()).trim();
  var data  = getRowData(sheet, row, c);
  if (c.RESULT && col === c.RESULT) {
    if (value === '合格') {
      var originalRow = c.ORIGINAL_ROW ? parseInt(sheet.getRange(row, c.ORIGINAL_ROW).getValue(), 10) : -1;
      if (originalRow > 0) {
        sendRenewalPaymentEmail(data, originalRow);
      } else {
        GmailApp.sendEmail(Session.getEffectiveUser().getEmail(), '【SEEDCARD】更新エラー：元行番号が見つかりません', '会員番号：' + data.memberNo + '\n手動で対応してください。');
      }
    } else if (value === '不合格') {
      sendRenewalRejectionEmail(data);
    }
  }
}

// ---- カード発行 ----

function issueCard(sheet, row, c, data) {
  var memberNo   = getNextMemberNumber(sheet, c);
  var uuid       = Utilities.getUuid();
  var cardUrl    = CONFIG.CARD_BASE_URL + '?t=' + uuid;
  var today      = new Date();
  var expiry     = new Date(today.getTime());
  expiry.setFullYear(expiry.getFullYear() + 1);
  expiry.setDate(expiry.getDate() - 1);
  var certDate   = Utilities.formatDate(today,  'Asia/Tokyo', 'yyyy年MM月dd日');
  var expiryDate = Utilities.formatDate(expiry, 'Asia/Tokyo', 'yyyy年MM月dd日');
  if (c.MEMBER_NO)      { sheet.getRange(row, c.MEMBER_NO).setValue(memberNo); }
  if (c.CERT_DATE)      { sheet.getRange(row, c.CERT_DATE).setValue(certDate); }
  if (c.EXPIRY_DATE)    { sheet.getRange(row, c.EXPIRY_DATE).setValue(expiryDate); }
  if (c.CARD_URL)       { sheet.getRange(row, c.CARD_URL).setValue(cardUrl); }
  if (c.SHIPPING)       { sheet.getRange(row, c.SHIPPING).setValue('発送待ち'); }
  if (c.RENEWAL_NOTICE) { sheet.getRange(row, c.RENEWAL_NOTICE).setValue(''); }
  sendIssuanceEmail(data, memberNo, cardUrl, certDate, expiryDate);
}

function getNextMemberNumber(sheet, c) {
  var year     = parseInt(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy'));
  var yearBase = year * 100000;
  var startNo  = yearBase + 777;
  if (!c.MEMBER_NO) { return startNo; }
  var max = startNo - 1;
  for (var i = 2; i <= sheet.getLastRow(); i++) {
    var v = sheet.getRange(i, c.MEMBER_NO).getValue();
    if (typeof v === 'number' && Math.floor(v / 100000) === year && v > max) { max = v; }
  }
  return max + 1;
}

// ---- 更新フロー ----

function checkRenewal() {
  var sheet   = getInitialSheet();
  var c       = getCols(sheet);
  if (!c.EXPIRY_DATE || !c.RENEWAL_NOTICE || !c.CARD_URL) { return; }
  var today   = new Date();
  for (var i = 2; i <= sheet.getLastRow(); i++) {
    var cardUrl = c.CARD_URL ? String(sheet.getRange(i, c.CARD_URL).getValue()).trim() : '';
    if (!cardUrl) { continue; }
    var expiryVal = sheet.getRange(i, c.EXPIRY_DATE).getValue();
    if (!expiryVal) { continue; }
    var expiry   = new Date(expiryVal);
    var daysLeft = Math.floor((expiry - today) / (1000 * 60 * 60 * 24));
    var notice   = String(sheet.getRange(i, c.RENEWAL_NOTICE).getValue()).trim();
    var data     = getRowData(sheet, i, c);
    if (!isValidEmail(data.email)) { continue; }
    if (daysLeft <= 60 && daysLeft > 30 && notice === '') {
      sendRenewalReminderEmail(data);
      sheet.getRange(i, c.RENEWAL_NOTICE).setValue('2ヶ月前通知済み');
    }
  }
}

function processRenewal(originalRow) {
  var props   = PropertiesService.getScriptProperties();
  var lockKey = 'renewal_done_' + originalRow;
  var lastRun = parseInt(props.getProperty(lockKey) || '0');
  if (Date.now() - lastRun < 23 * 3600000) { return; }
  props.setProperty(lockKey, String(Date.now()));

  var sheet = getInitialSheet();
  var c     = getCols(sheet);
  if (!c.EXPIRY_DATE) { return; }

  var base = new Date();
  var newExp = new Date(base.getTime());
  newExp.setFullYear(newExp.getFullYear() + 1);
  newExp.setDate(newExp.getDate() - 1);
  var newExpiryDate = Utilities.formatDate(newExp, 'Asia/Tokyo', 'yyyy年MM月dd日');
  sheet.getRange(originalRow, c.EXPIRY_DATE).setValue(newExpiryDate);
  if (c.RENEWAL_NOTICE) { sheet.getRange(originalRow, c.RENEWAL_NOTICE).setValue(''); }

  var renewalSheet = getRenewalSheet();
  var rc = getCols(renewalSheet);
  for (var i = 2; i <= renewalSheet.getLastRow(); i++) {
    var origRowVal = rc.ORIGINAL_ROW ? parseInt(renewalSheet.getRange(i, rc.ORIGINAL_ROW).getValue(), 10) : -1;
    if (origRowVal === originalRow) {
      if (c.VOLUME        && rc.VOLUME)        { sheet.getRange(originalRow, c.VOLUME).setValue(renewalSheet.getRange(i, rc.VOLUME).getValue()); }
      if (c.CONCENTRATION && rc.CONCENTRATION) { sheet.getRange(originalRow, c.CONCENTRATION).setValue(renewalSheet.getRange(i, rc.CONCENTRATION).getValue()); }
      if (c.TOTAL         && rc.TOTAL)         { sheet.getRange(originalRow, c.TOTAL).setValue(renewalSheet.getRange(i, rc.TOTAL).getValue()); }
      if (c.MORPHOLOGY    && rc.MORPHOLOGY)    { sheet.getRange(originalRow, c.MORPHOLOGY).setValue(renewalSheet.getRange(i, rc.MORPHOLOGY).getValue()); }
      break;
    }
  }

  var data    = getRowData(sheet, originalRow, c);
  var cardUrl = c.CARD_URL ? String(sheet.getRange(originalRow, c.CARD_URL).getValue()).trim() : '';
  sendRenewalConfirmationEmail(data, newExpiryDate, cardUrl);
}

// ---- メール関数 ----

function isValidEmail(email) {
  return typeof email === 'string' && email.length > 0 && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function sendAcknowledgmentEmail(data) {
  if (!isValidEmail(data.email)) { return; }
  GmailApp.sendEmail(data.email, '【SEED CARD】申請を受け付けました',
    (data.nameKanji || 'お客様') + ' 様\n\n'
    + 'SEED CARD へのご申請ありがとうございます。\n'
    + '内容を確認中です。審査結果は数日以内にメールでお知らせいたします。\n\n'
    + CONFIG.FROM_NAME);
}

function sendStaffNotification(data) {
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var name = data.nameKanji || '（名前取得失敗）';
  var photoLine = data.photo ? '検査証画像：\n' + data.photo + '\n\n' : '';
  GmailApp.sendEmail(Session.getEffectiveUser().getEmail(),
    '【SEED CARD申請】' + name + '様',
    '新しい申請が届きました。管理パネルで審査してください。\n\n'
    + '申請者：' + name + '\nメール：' + data.email + '\n住所：' + data.postal + ' ' + data.address + '\n\n'
    + '精液量：' + data.volume + ' mL\n精子濃度：' + data.concentration + ' 万/mL\n総精子数：' + data.total + ' 万\n正常形態率：' + data.morphology + ' %\n\n'
    + photoLine
    + '▼ 管理パネル\nhttps://myseedcard.japan-sperm.com/admin.html\n\n'
    + '▼ スプレッドシート\nhttps://docs.google.com/spreadsheets/d/' + spreadsheetId);
}

function sendRenewalStaffNotification(data, originalRow) {
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var name   = data.nameKanji || '（名前取得失敗）';
  var status = originalRow > 0 ? '会員確認済み（行' + originalRow + '）' : '★会員番号が見つかりません。確認してください';
  var photoLine = data.photo ? '検査証画像：\n' + data.photo + '\n\n' : '';
  GmailApp.sendEmail(Session.getEffectiveUser().getEmail(),
    '【SEED CARD更新申請】' + name + '様',
    '更新申請が届きました。管理パネルで審査してください。\n\n'
    + '申請者：' + name + '\n会員番号：' + data.memberNo + '\nメール：' + data.email + '\n会員状態：' + status + '\n\n'
    + photoLine
    + '▼ 管理パネル\nhttps://myseedcard.japan-sperm.com/admin.html\n\n'
    + '▼ スプレッドシート\nhttps://docs.google.com/spreadsheets/d/' + spreadsheetId);
}

function sendPaymentEmail(data, row) {
  if (!isValidEmail(data.email)) { return; }
  GmailApp.sendEmail(data.email, '【SEED CARD】審査通過のご連絡',
    (data.nameKanji || 'お客様') + ' 様\n\n'
    + '精液検査の結果を確認いたしました。WHO 2021基準を全項目クリアされています。\n\n'
    + '以下のリンクより、認定料のお支払いをお願いいたします。\n\n'
    + '>> お支払いはこちら（22,000円）\n'
    + CONFIG.STRIPE_URL_INITIAL + '?client_reference_id=' + row + '\n\n'
    + 'お支払い確認後、SEED CARDを発行・発送いたします。\n\n' + CONFIG.FROM_NAME);
}

function sendRejectionEmail(data) {
  if (!isValidEmail(data.email)) { return; }
  GmailApp.sendEmail(data.email, '【SEED CARD】審査結果のご連絡',
    (data.nameKanji || 'お客様') + ' 様\n\n'
    + '精液検査の結果を確認いたしましたが、現時点では認定基準を満たしていないため、今回のご申請は見送りとさせていただきます。\n'
    + '検査結果を改善された後に、再度ご申請ください。\n\n' + CONFIG.FROM_NAME);
}

function sendRenewalPaymentEmail(data, originalRow) {
  if (!isValidEmail(data.email)) { return; }
  GmailApp.sendEmail(data.email, '【SEED CARD】更新審査通過のご連絡',
    (data.nameKanji || 'お客様') + ' 様\n\n'
    + '更新申請の検査結果を確認いたしました。WHO 2021基準を全項目クリアされています。\n\n'
    + '>> 更新手続きはこちら（10,000円）\n'
    + CONFIG.STRIPE_URL_RENEWAL + '?client_reference_id=renew_' + originalRow + '\n\n'
    + 'お支払い確認後、認定期限を1年間延長いたします。\n\n' + CONFIG.FROM_NAME);
}

function sendRenewalRejectionEmail(data) {
  if (!isValidEmail(data.email)) { return; }
  GmailApp.sendEmail(data.email, '【SEED CARD】更新審査結果のご連絡',
    (data.nameKanji || 'お客様') + ' 様\n\n'
    + '更新申請の検査結果を確認いたしましたが、現時点では認定基準を満たしていないため、今回の更新は見送りとさせていただきます。\n\n' + CONFIG.FROM_NAME);
}

function sendShippingEmail(data, cardUrl) {
  if (!isValidEmail(data.email)) { return; }
  GmailApp.sendEmail(data.email, '【SEED CARD】カードを発送しました',
    (data.nameKanji || 'お客様') + ' 様\n\nSEED CARDを発送いたしました。数日以内にお届けいたします。\n\n'
    + 'カードが届く前でも、以下のマイページから認定情報をご確認いただけます。\n\n'
    + CONFIG.PORTAL_URL + '\n\n' + CONFIG.FROM_NAME);
}

function sendRenewalReminderEmail(data) {
  if (!isValidEmail(data.email)) { return; }
  GmailApp.sendEmail(data.email, '【SEED CARD】認定更新のご案内',
    (data.nameKanji || 'お客様') + ' 様\n\n'
    + '認定期限まで約2ヶ月となりました。\n\n'
    + '更新を希望される場合は、お早めに精液検査をお受けください。\n'
    + '検査後、以下のマイページから更新申請をお手続きください。\n\n'
    + CONFIG.PORTAL_URL + '\n\n'
    + '審査通過後に更新料（10,000円）のお支払いリンクをお送りします。\n\n' + CONFIG.FROM_NAME);
}

function sendIssuanceEmail(data, memberNo, cardUrl, certDate, expiryDate) {
  if (!isValidEmail(data.email)) { return; }
  GmailApp.sendEmail(data.email, '【SEED CARD】認定カードを発行しました',
    (data.nameKanji || 'お客様') + ' 様\n\n'
    + 'お支払いいただきありがとうございます。SEED CARD を発行いたしました。\n\n'
    + '会員番号：No.' + String(memberNo) + '\n認定日：' + certDate + '\n認定期限：' + expiryDate + '\n\n'
    + '▼ あなたの認定ページ\n' + cardUrl + '\n\n'
    + '▼ マイページ（ログイン）\n' + CONFIG.PORTAL_URL + '\n\n'
    + '物理カードのNFCをスマートフォンにかざすと認定情報ページが開きます。\n\n' + CONFIG.FROM_NAME);
}

function sendRenewalConfirmationEmail(data, newExpiryDate, cardUrl) {
  if (!isValidEmail(data.email)) { return; }
  GmailApp.sendEmail(data.email, '【SEED CARD】認定を更新しました',
    (data.nameKanji || 'お客様') + ' 様\n\n'
    + '更新料のお支払いを確認いたしました。認定期限を更新いたしました。\n\n'
    + '新しい認定期限：' + newExpiryDate + '\n\n'
    + 'マイページ：\n' + CONFIG.PORTAL_URL + '\n\n'
    + 'なお、新しいカードは送付いたしません。現在お手持ちのカードを引き続きご使用ください。\n\n' + CONFIG.FROM_NAME);
}

// ---- DriveサムネイルURL変換 ----

function driveUrlToThumb(url) {
  if (!url) return '';
  var s = String(url).trim();
  var m = s.match(/[?&]id=([^&]+)/) || s.match(/\/d\/([^\/\?]+)/);
  return m ? 'https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w600' : s;
}

// ---- マジックリンク認証 ----

function checkRateLimit(key) {
  var props = PropertiesService.getScriptProperties();
  var last  = parseInt(props.getProperty('rl_' + key) || '0');
  if (Date.now() - last < 1 * 60 * 1000) { return false; }
  props.setProperty('rl_' + key, String(Date.now()));
  return true;
}

function sendMagicLink(email) {
  var out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  if (!email) { out.setContent(JSON.stringify({error:'no_email'})); return out; }
  if (!checkRateLimit('ml_' + email.toLowerCase())) {
    out.setContent(JSON.stringify({error:'rate_limit'})); return out;
  }
  var sheet = getInitialSheet(); var c = getCols(sheet);
  var found = false;
  for (var i = 2; i <= sheet.getLastRow(); i++) {
    var rowEmail = c.EMAIL ? String(sheet.getRange(i, c.EMAIL).getValue()).trim().toLowerCase() : '';
    if (rowEmail && rowEmail === email.toLowerCase()) { found = true; break; }
  }
  if (!found) { out.setContent(JSON.stringify({error:'not_found'})); return out; }

  var token  = Utilities.getUuid();
  var expiry = Date.now() + 24 * 3600000;
  PropertiesService.getScriptProperties().setProperty('magic_' + token, email.toLowerCase() + '|' + expiry);

  // ポータルのURLに token パラメータで送る
  var magicUrl = CONFIG.PORTAL_URL + '?token=' + token;
  GmailApp.sendEmail(email, '【SEED CARD】会員マイページへのログインリンク',
    'SEED CARD 会員マイページへのログインリンクをお送りします。\n\n'
    + '▼ ログインする（有効期限：24時間）\n' + magicUrl + '\n\n'
    + 'このリンクはあなた専用です。第三者と共有しないでください。\n\n'
    + '━━━━━━━━━━━━━━━━━━━━\n一般社団法人 日本精子協会\njapan.sperm.association@gmail.com\n━━━━━━━━━━━━━━━━━━━━',
    {name: '一般社団法人 日本精子協会'});
  out.setContent(JSON.stringify({ok: true}));
  return out;
}

// ---- JSON配信（カードページ用） ----

function formatDateJP(d) {
  if (!d) return '';
  var date = new Date(d);
  if (isNaN(date.getTime())) return String(d);
  return date.getFullYear() + '年' + (date.getMonth() + 1) + '月' + date.getDate() + '日';
}

function normalizeToWan(val) {
  var s = String(val || '').replace(/[,，\s　]/g, '');
  if (!s || s === '0') return 0;
  if (/[億万]/.test(s)) {
    var om = s.match(/([0-9.]+)\s*億/); var mm = s.match(/([0-9.]+)\s*万/);
    return Math.round((om ? parseFloat(om[1]) * 10000 : 0) + (mm ? parseFloat(mm[1]) : 0));
  }
  var n = parseFloat(s);
  if (isNaN(n) || n <= 0) return 0;
  if (n >= 1000000) return Math.round(n / 10000);
  if (n >= 1000)    return Math.round(n);
  if (n >= 10)      return Math.round(n * 100);
  return Math.round(n * 10000);
}

function serveJson(t) {
  var out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  if (!t) { out.setContent(JSON.stringify({error:'no_token'})); return out; }
  var sheet = getInitialSheet(); var c = getCols(sheet);
  for (var i = 2; i <= sheet.getLastRow(); i++) {
    var cardUrl = c.CARD_URL ? String(sheet.getRange(i, c.CARD_URL).getValue()).trim() : '';
    if (cardUrl.indexOf(t) >= 0) {
      var data = getRowData(sheet, i, c);
      var memberNo = c.MEMBER_NO   ? sheet.getRange(i, c.MEMBER_NO).getValue()   : 0;
      var certDate = c.CERT_DATE   ? sheet.getRange(i, c.CERT_DATE).getValue()   : '';
      var expiry   = c.EXPIRY_DATE ? sheet.getRange(i, c.EXPIRY_DATE).getValue() : '';
      out.setContent(JSON.stringify({
        name:data.nameKanji||'', nameRomaji:data.nameRomaji||'',
        memberNo:memberNo||0, certDate:formatDateJP(certDate), expiryDate:formatDateJP(expiry),
        vol:data.volume, conc:data.concentration, total:data.total, morph:data.morphology
      }));
      return out;
    }
  }
  out.setContent(JSON.stringify({error:'not_found'})); return out;
}

// ---- doGet ----

function doGet(e) {
  var t = e && e.parameter && e.parameter.t ? String(e.parameter.t).trim() : '';

  // ── マジックリンク送信 ──
  if (e && e.parameter && e.parameter.action === 'sendmagiclink') {
    return sendMagicLink(String(e.parameter.email || '').trim());
  }

  // ── 会員番号でマジックリンク送信 ──
  if (e && e.parameter && e.parameter.action === 'sendmagicbymemberno') {
    var out = ContentService.createTextOutput();
    out.setMimeType(ContentService.MimeType.JSON);
    var inputNo = String(e.parameter.memberNo || '').replace(/[^0-9]/g, '');
    if (!inputNo) { out.setContent(JSON.stringify({error:'missing_params'})); return out; }
    var s5 = getInitialSheet(); var c5 = getCols(s5);
    for (var i = 2; i <= s5.getLastRow(); i++) {
      var sNo = String(c5.MEMBER_NO ? s5.getRange(i, c5.MEMBER_NO).getValue() : '').replace(/[^0-9]/g,'');
      if (sNo && sNo === inputNo) {
        var em5 = c5.EMAIL ? String(s5.getRange(i, c5.EMAIL).getValue()).trim() : '';
        if (!em5) { out.setContent(JSON.stringify({error:'no_email'})); return out; }
        return sendMagicLink(em5);
      }
    }
    out.setContent(JSON.stringify({error:'not_found'})); return out;
  }

  // ── 会員番号＋氏名照合 → 登録メールにマジックリンク送信（メアド忘れ用） ──
  if (e && e.parameter && e.parameter.action === 'lookupbymember') {
    var out = ContentService.createTextOutput();
    out.setMimeType(ContentService.MimeType.JSON);
    var lmNo   = String(e.parameter.memberNo || '').replace(/[^0-9]/g, '');
    var lmName = String(e.parameter.name     || '').trim();
    if (!lmNo || !lmName) { out.setContent(JSON.stringify({error:'missing_params'})); return out; }
    var lmSheet = getInitialSheet(); var lmC = getCols(lmSheet);
    for (var i = 2; i <= lmSheet.getLastRow(); i++) {
      var lmSheetNo   = String(lmC.MEMBER_NO  ? lmSheet.getRange(i, lmC.MEMBER_NO).getValue()  : '').replace(/[^0-9]/g,'');
      var lmSheetName = String(lmC.NAME_KANJI ? lmSheet.getRange(i, lmC.NAME_KANJI).getValue() : '').trim();
      if (lmSheetNo && lmSheetNo === lmNo && lmSheetName === lmName) {
        var lmEmail = lmC.EMAIL ? String(lmSheet.getRange(i, lmC.EMAIL).getValue()).trim() : '';
        if (!lmEmail) { out.setContent(JSON.stringify({error:'no_email'})); return out; }
        return sendMagicLink(lmEmail);
      }
    }
    out.setContent(JSON.stringify({error:'not_found'})); return out;
  }

  // ── メールアドレスで会員番号をメール送信（会員番号忘れ用） ──
  if (e && e.parameter && e.parameter.action === 'sendmemberno') {
    var out = ContentService.createTextOutput();
    out.setMimeType(ContentService.MimeType.JSON);
    var snEmail = String(e.parameter.email || '').trim().toLowerCase();
    if (!snEmail) { out.setContent(JSON.stringify({error:'missing_params'})); return out; }
    var snSheet = getInitialSheet(); var snC = getCols(snSheet);
    for (var i = 2; i <= snSheet.getLastRow(); i++) {
      var snRowEmail = snC.EMAIL ? String(snSheet.getRange(i, snC.EMAIL).getValue()).trim().toLowerCase() : '';
      if (snRowEmail && snRowEmail === snEmail) {
        var snNo   = snC.MEMBER_NO  ? String(snSheet.getRange(i, snC.MEMBER_NO).getValue()).trim()  : '';
        var snName = snC.NAME_KANJI ? String(snSheet.getRange(i, snC.NAME_KANJI).getValue()).trim() : '';
        if (!snNo) { out.setContent(JSON.stringify({error:'no_member_no'})); return out; }
        GmailApp.sendEmail(snEmail, '【SEED CARD】会員番号のご案内',
          (snName || 'お客様') + ' 様\n\n会員番号をお知らせします。\n\n'
          + '会員番号：No.' + snNo + '\n\n'
          + 'マイページへのログインは以下からどうぞ。\n' + CONFIG.PORTAL_URL + '\n\n'
          + CONFIG.FROM_NAME,
          {name: '一般社団法人 日本精子協会'});
        out.setContent(JSON.stringify({ok: true})); return out;
      }
    }
    out.setContent(JSON.stringify({error:'not_found'})); return out;
  }

  // ── トークン検証（ポータルログイン用） ──
  if (e && e.parameter && e.parameter.action === 'verifytoken') {
    var out = ContentService.createTextOutput();
    out.setMimeType(ContentService.MimeType.JSON);
    var tok = String(e.parameter.token || '').trim();
    if (!tok) { out.setContent(JSON.stringify({error:'no_token'})); return out; }
    var props = PropertiesService.getScriptProperties();
    var val   = props.getProperty('magic_' + tok);
    if (!val) { out.setContent(JSON.stringify({error:'invalid'})); return out; }
    var parts = val.split('|');
    var tokEmail = parts[0]; var tokExpiry = parseInt(parts[1] || '0');
    if (Date.now() > tokExpiry) { props.deleteProperty('magic_' + tok); out.setContent(JSON.stringify({error:'expired'})); return out; }
    out.setContent(JSON.stringify({ok: true, email: tokEmail}));
    return out;
  }

  // ── メールアドレスで会員情報取得（マイページ用） ──
  if (e && e.parameter && e.parameter.action === 'getmemberbyemail') {
    var out = ContentService.createTextOutput();
    out.setMimeType(ContentService.MimeType.JSON);
    var gmEmail = String(e.parameter.email || '').trim().toLowerCase();
    out.setContent(JSON.stringify(handleGetMemberByEmail(gmEmail)));
    return out;
  }

  // ── 管理者API ──
  if (e && e.parameter && e.parameter.action && e.parameter.action.indexOf('admin_') === 0) {
    var out = ContentService.createTextOutput();
    out.setMimeType(ContentService.MimeType.JSON);
    var adminKey   = String(e.parameter.key || '').trim();
    var storedKey  = String(PropertiesService.getScriptProperties().getProperty('ADMIN_KEY') || '').trim();
    if (!storedKey || adminKey !== storedKey) { out.setContent(JSON.stringify({error:'unauthorized'})); return out; }

    if (e.parameter.action === 'admin_get') {
      var sheet = getInitialSheet(); var c = getCols(sheet);
      var initial = [];
      for (var i = 2; i <= sheet.getLastRow(); i++) {
        var rName = c.NAME_KANJI ? String(sheet.getRange(i, c.NAME_KANJI).getValue()).trim() : '';
        var rEmail= c.EMAIL      ? String(sheet.getRange(i, c.EMAIL).getValue()).trim()      : '';
        if (!rName && !rEmail) continue;
        var rTs = sheet.getRange(i, c.TIMESTAMP).getValue();
        initial.push({
          row:i, ts: rTs ? Utilities.formatDate(new Date(rTs), 'Asia/Tokyo', 'MM/dd HH:mm') : '',
          name:rName,
          nameRomaji:  c.NAME_ROMAJI   ? String(sheet.getRange(i, c.NAME_ROMAJI).getValue()).trim()  : '',
          email:rEmail,
          postal:      c.POSTAL        ? String(sheet.getRange(i, c.POSTAL).getValue()).trim()       : '',
          address:     c.ADDRESS       ? String(sheet.getRange(i, c.ADDRESS).getValue()).trim()      : '',
          volume:      c.VOLUME        ? Number(sheet.getRange(i, c.VOLUME).getValue())       || 0 : 0,
          concentration: c.CONCENTRATION ? Number(sheet.getRange(i, c.CONCENTRATION).getValue()) || 0 : 0,
          total:       c.TOTAL         ? Number(sheet.getRange(i, c.TOTAL).getValue())        || 0 : 0,
          morphology:  c.MORPHOLOGY    ? Number(sheet.getRange(i, c.MORPHOLOGY).getValue())   || 0 : 0,
          result:      c.RESULT        ? String(sheet.getRange(i, c.RESULT).getValue()).trim()       : '',
          payment:     c.PAYMENT       ? String(sheet.getRange(i, c.PAYMENT).getValue()).trim()      : '',
          shipping:    c.SHIPPING      ? String(sheet.getRange(i, c.SHIPPING).getValue()).trim()     : '',
          memberNo:    c.MEMBER_NO     ? String(sheet.getRange(i, c.MEMBER_NO).getValue()).trim()    : '',
          certDate:    c.CERT_DATE     ? String(sheet.getRange(i, c.CERT_DATE).getValue()).trim()    : '',
          expiryDate:  c.EXPIRY_DATE   ? String(sheet.getRange(i, c.EXPIRY_DATE).getValue()).trim()  : '',
          cardUrl:     c.CARD_URL      ? String(sheet.getRange(i, c.CARD_URL).getValue()).trim()     : '',
          photo:       driveUrlToThumb(c.PHOTO ? String(sheet.getRange(i, c.PHOTO).getValue()).trim() : '')
        });
      }
      var rs = getRenewalSheet(); var rc = getCols(rs);
      var renewal = [];
      for (var j = 2; j <= rs.getLastRow(); j++) {
        var rrName = rc.NAME_KANJI ? String(rs.getRange(j, rc.NAME_KANJI).getValue()).trim() : '';
        var rrNo   = rc.MEMBER_NO  ? String(rs.getRange(j, rc.MEMBER_NO).getValue()).trim()  : '';
        if (!rrName && !rrNo) continue;
        var rrTs = rs.getRange(j, rc.TIMESTAMP).getValue();
        renewal.push({
          row:j, ts: rrTs ? Utilities.formatDate(new Date(rrTs), 'Asia/Tokyo', 'MM/dd HH:mm') : '',
          memberNo:rrNo, name:rrName,
          volume:        rc.VOLUME         ? Number(rs.getRange(j, rc.VOLUME).getValue())        || 0 : 0,
          concentration: rc.CONCENTRATION  ? Number(rs.getRange(j, rc.CONCENTRATION).getValue()) || 0 : 0,
          total:         rc.TOTAL          ? Number(rs.getRange(j, rc.TOTAL).getValue())         || 0 : 0,
          morphology:    rc.MORPHOLOGY     ? Number(rs.getRange(j, rc.MORPHOLOGY).getValue())    || 0 : 0,
          result:        rc.RESULT         ? String(rs.getRange(j, rc.RESULT).getValue()).trim()       : '',
          originalRow:   rc.ORIGINAL_ROW   ? parseInt(rs.getRange(j, rc.ORIGINAL_ROW).getValue(), 10) : -1,
          photo:         driveUrlToThumb(rc.PHOTO ? String(rs.getRange(j, rc.PHOTO).getValue()).trim() : '')
        });
      }
      out.setContent(JSON.stringify({initial:initial, renewal:renewal}));
      return out;
    }

    if (e.parameter.action === 'admin_setresult') {
      var arRow = parseInt(e.parameter.row || '0', 10);
      var arResult = String(e.parameter.result || '').trim();
      if (!arRow || ['合格','不合格'].indexOf(arResult) < 0) { out.setContent(JSON.stringify({error:'invalid_params'})); return out; }
      var arSheet = getInitialSheet(); var arC = getCols(arSheet);
      var arData = getRowData(arSheet, arRow, arC);
      if (arC.RESULT)  { arSheet.getRange(arRow, arC.RESULT).setValue(arResult); }
      if (arResult === '合格') {
        if (arC.PAYMENT) { arSheet.getRange(arRow, arC.PAYMENT).setValue('支払い待ち'); }
        sendPaymentEmail(arData, arRow);
      } else { sendRejectionEmail(arData); }
      out.setContent(JSON.stringify({success:true})); return out;
    }

    if (e.parameter.action === 'admin_setshipping') {
      var asRow = parseInt(e.parameter.row || '0', 10);
      if (!asRow) { out.setContent(JSON.stringify({error:'invalid_params'})); return out; }
      var asSheet = getInitialSheet(); var asC = getCols(asSheet);
      var asData = getRowData(asSheet, asRow, asC);
      if (asC.SHIPPING) { asSheet.getRange(asRow, asC.SHIPPING).setValue('発送済み'); }
      var asCardUrl = asC.CARD_URL ? String(asSheet.getRange(asRow, asC.CARD_URL).getValue()).trim() : '';
      sendShippingEmail(asData, asCardUrl);
      out.setContent(JSON.stringify({success:true})); return out;
    }

    if (e.parameter.action === 'admin_updatemember') {
      var umRow = parseInt(e.parameter.row || '0', 10);
      if (!umRow) { out.setContent(JSON.stringify({error:'invalid_params'})); return out; }
      var umSheet = getInitialSheet(); var umC = getCols(umSheet);
      if (e.parameter.name       && umC.NAME_KANJI)   { umSheet.getRange(umRow, umC.NAME_KANJI).setValue(String(e.parameter.name).trim()); }
      if (e.parameter.nameRomaji !== undefined && umC.NAME_ROMAJI) { umSheet.getRange(umRow, umC.NAME_ROMAJI).setValue(String(e.parameter.nameRomaji).trim()); }
      if (e.parameter.email      && umC.EMAIL)        { umSheet.getRange(umRow, umC.EMAIL).setValue(String(e.parameter.email).trim().toLowerCase()); }
      if (e.parameter.postal     !== undefined && umC.POSTAL)     { umSheet.getRange(umRow, umC.POSTAL).setValue(String(e.parameter.postal).trim()); }
      if (e.parameter.address    !== undefined && umC.ADDRESS)    { umSheet.getRange(umRow, umC.ADDRESS).setValue(String(e.parameter.address).trim()); }
      if (e.parameter.expiryDate !== undefined && umC.EXPIRY_DATE){ umSheet.getRange(umRow, umC.EXPIRY_DATE).setValue(String(e.parameter.expiryDate).trim()); }
      if (e.parameter.volume        !== undefined && umC.VOLUME)       { var v1=parseFloat(e.parameter.volume);       if(!isNaN(v1)) umSheet.getRange(umRow, umC.VOLUME).setValue(v1); }
      if (e.parameter.concentration !== undefined && umC.CONCENTRATION){ var v2=parseFloat(e.parameter.concentration); if(!isNaN(v2)) umSheet.getRange(umRow, umC.CONCENTRATION).setValue(v2); }
      if (e.parameter.total         !== undefined && umC.TOTAL)        { var v3=parseFloat(e.parameter.total);         if(!isNaN(v3)) umSheet.getRange(umRow, umC.TOTAL).setValue(v3); }
      if (e.parameter.morphology    !== undefined && umC.MORPHOLOGY)   { var v4=parseFloat(e.parameter.morphology);    if(!isNaN(v4)) umSheet.getRange(umRow, umC.MORPHOLOGY).setValue(v4); }
      out.setContent(JSON.stringify({success:true})); return out;
    }

    if (e.parameter.action === 'admin_deletemember') {
      var dmRow = parseInt(e.parameter.row || '0', 10);
      if (!dmRow || dmRow < 2) { out.setContent(JSON.stringify({error:'invalid_params'})); return out; }
      var dmSheet = getInitialSheet();
      if (dmRow > dmSheet.getLastRow()) { out.setContent(JSON.stringify({error:'row_not_found'})); return out; }
      dmSheet.deleteRow(dmRow);
      out.setContent(JSON.stringify({success:true})); return out;
    }

    if (e.parameter.action === 'admin_setrenewalresult') {
      var rrRow = parseInt(e.parameter.row || '0', 10);
      var rrResult = String(e.parameter.result || '').trim();
      if (!rrRow || ['合格','不合格'].indexOf(rrResult) < 0) { out.setContent(JSON.stringify({error:'invalid_params'})); return out; }
      var rrSheet = getRenewalSheet(); var rrC = getCols(rrSheet);
      var rrData = getRowData(rrSheet, rrRow, rrC);
      if (rrC.RESULT) { rrSheet.getRange(rrRow, rrC.RESULT).setValue(rrResult); }
      if (rrResult === '合格') {
        var rrOrigRow = rrC.ORIGINAL_ROW ? parseInt(rrSheet.getRange(rrRow, rrC.ORIGINAL_ROW).getValue(), 10) : -1;
        if (rrOrigRow > 0) { sendRenewalPaymentEmail(rrData, rrOrigRow); }
      } else { sendRenewalRejectionEmail(rrData); }
      out.setContent(JSON.stringify({success:true})); return out;
    }

    out.setContent(JSON.stringify({error:'unknown_action'})); return out;
  }

  // ── カードページJSON（api=1）──
  if (e && e.parameter && e.parameter.api === '1') {
    return serveJson(t);
  }

  // ── カードページHTMLサーブ ──
  if (!t) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:40px">URLが正しくありません。</p>');
  }
  var sheet = getInitialSheet(); var c = getCols(sheet);
  for (var i = 2; i <= sheet.getLastRow(); i++) {
    var cardUrl = c.CARD_URL ? String(sheet.getRange(i, c.CARD_URL).getValue()).trim() : '';
    if (cardUrl.indexOf(t) >= 0) {
      var data = getRowData(sheet, i, c);
      var memberNo  = c.MEMBER_NO   ? sheet.getRange(i, c.MEMBER_NO).getValue()   : 0;
      var certDate  = c.CERT_DATE   ? sheet.getRange(i, c.CERT_DATE).getValue()   : '';
      var expiry    = c.EXPIRY_DATE ? sheet.getRange(i, c.EXPIRY_DATE).getValue() : '';
      var memberData = {
        name:data.nameKanji||'', nameRomaji:data.nameRomaji||'', memberNo:memberNo||0,
        certDate:formatDateJP(certDate), expiryDate:formatDateJP(expiry),
        vol:data.volume, conc:data.concentration, total:data.total, morph:data.morphology
      };
      var html = HtmlService.createHtmlOutputFromFile('card');
      var content = html.getContent();
      var noStr = 'No.' + String(memberData.memberNo || 0);
      content = content.split('{{NAME_KANJI}}').join(memberData.name)
        .split('{{NAME_ROMAJI}}').join((memberData.nameRomaji||'').toUpperCase()||'MEMBER')
        .split('{{MEMBER_NO}}').join(noStr)
        .split('{{CERT_YEAR}}').join(String(memberData.certDate||'').substring(0,4))
        .split('{{CERT_DATE}}').join(memberData.certDate||'')
        .split('{{EXPIRY_DATE}}').join(memberData.expiryDate||'')
        .split('{{VOL}}').join(memberData.vol   > 0 ? String(Math.round(memberData.vol*10)/10) : '—')
        .split('{{CONC}}').join(memberData.conc  > 0 ? String(Math.round(memberData.conc))     : '—')
        .split('{{TOTAL}}').join(memberData.total > 0 ? String(Math.round(memberData.total/10000*10)/10) : '—')
        .split('{{MORPH}}').join(memberData.morph > 0 ? String(Math.round(memberData.morph))   : '—');
      content = content.replace('</body>', '<script>window.__SD='+JSON.stringify(memberData)+';<\/script></body>');
      return HtmlService.createHtmlOutput(content).setTitle('SEED CARD｜認定').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }
  return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:40px">認定情報が見つかりません。</p>');
}

// ---- doPost（Stripe webhook + ポータルフォーム送信） ----

function doPost(e) {
  // ポータルからのフォーム送信（URL-encoded、actionパラメータあり）
  if (e.parameter && e.parameter.action) {
    var out = ContentService.createTextOutput();
    out.setMimeType(ContentService.MimeType.JSON);
    var action = String(e.parameter.action).trim();
    if (action === 'submitapply') {
      out.setContent(JSON.stringify(handleSubmitApply(e.parameter)));
    } else if (action === 'submitrenew') {
      out.setContent(JSON.stringify(handleSubmitRenew(e.parameter)));
    } else {
      out.setContent(JSON.stringify({error:'unknown_action'}));
    }
    return out;
  }

  // Stripe Webhook（JSONボディ）
  try {
    var raw     = e.postData ? e.postData.contents : '';
    var payload = JSON.parse(raw);
    var evType  = payload.type || '';
    if (evType === 'checkout.session.completed' || evType === 'payment_intent.succeeded') {
      var obj   = payload.data && payload.data.object ? payload.data.object : {};
      var refId = String(obj.client_reference_id || '').trim();
      if (refId.indexOf('renew_') === 0) {
        var row = parseInt(refId.replace('renew_', ''), 10);
        if (!isNaN(row)) { processRenewal(row); }
      } else if (refId) {
        markPaymentCompleteByRow(parseInt(refId, 10));
      } else {
        var email = (obj.customer_details && obj.customer_details.email) || obj.customer_email || obj.receipt_email || '';
        if (email) { markPaymentCompleteByEmail(email); }
      }
    }
  } catch(err) { Logger.log('webhook error: ' + err); }
  return ContentService.createTextOutput('ok');
}

function markPaymentCompleteByRow(row) {
  var sheet = getInitialSheet(); var c = getCols(sheet);
  if (!c.PAYMENT) { return; }
  if (String(sheet.getRange(row, c.PAYMENT).getValue()).trim() === '支払い待ち') {
    sheet.getRange(row, c.PAYMENT).setValue('支払い済み');
    issueCard(sheet, row, c, getRowData(sheet, row, c));
  }
}

function markPaymentCompleteByEmail(email) {
  var sheet = getInitialSheet(); var c = getCols(sheet);
  if (!c.EMAIL || !c.PAYMENT) { return; }
  for (var i = 2; i <= sheet.getLastRow(); i++) {
    if (String(sheet.getRange(i, c.EMAIL).getValue()).trim().toLowerCase() === email.toLowerCase()) {
      if (String(sheet.getRange(i, c.PAYMENT).getValue()).trim() === '支払い待ち') {
        sheet.getRange(i, c.PAYMENT).setValue('支払い済み');
        issueCard(sheet, i, c, getRowData(sheet, i, c));
      }
      break;
    }
  }
}

// ---- 30日未対応の申請を自動削除 ----

function checkExpiredApplications() {
  var sheet = getInitialSheet();
  var c = getCols(sheet);
  if (!c.TIMESTAMP || !c.RESULT) { return; }

  var now = new Date();
  var limit = 30 * 24 * 3600000; // 30日（ミリ秒）

  // 行がずれないよう後ろから走査
  for (var i = sheet.getLastRow(); i >= 2; i--) {
    var result  = c.RESULT  ? String(sheet.getRange(i, c.RESULT).getValue()).trim()  : '';
    var payment = c.PAYMENT ? String(sheet.getRange(i, c.PAYMENT).getValue()).trim() : '';

    // 対象：審査中 または 合格後に支払い待ちのまま
    var inProgress = (result === '審査中') || (result === '合格' && payment === '支払い待ち');
    if (!inProgress) { continue; }

    var ts = sheet.getRange(i, c.TIMESTAMP).getValue();
    if (!ts) { continue; }
    var tsDate = new Date(ts);
    if (isNaN(tsDate.getTime())) { continue; }
    if (now - tsDate <= limit) { continue; }

    // 削除前に本人へメール通知
    var email = c.EMAIL     ? String(sheet.getRange(i, c.EMAIL).getValue()).trim()     : '';
    var name  = c.NAME_KANJI? String(sheet.getRange(i, c.NAME_KANJI).getValue()).trim(): '';
    if (isValidEmail(email)) {
      GmailApp.sendEmail(email, '【SEED CARD】申請が自動キャンセルされました',
        (name || 'お客様') + ' 様\n\n'
        + '申請から30日以上が経過したため、申請データを自動的に削除いたしました。\n\n'
        + '再度お申し込みを希望される場合は、以下のポータルより新規申請をお願いいたします。\n\n'
        + CONFIG.PORTAL_URL + '\n\n' + CONFIG.FROM_NAME,
        {name: '一般社団法人 日本精子協会'});
    }
    sheet.deleteRow(i);
    Logger.log('自動削除: 行' + i + ' / ' + name + ' / ' + email);
  }
}

// 毎日午前2時に実行するトリガーを設定（一度だけ実行する）
function setupAutoDeleteTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'checkExpiredApplications') { ScriptApp.deleteTrigger(t); }
  });
  ScriptApp.newTrigger('checkExpiredApplications').timeBased().everyDays(1).atHour(2).create();
  Logger.log('自動削除トリガーを設定しました');
}

// ---- 週次バックアップ ----

function weeklyBackup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var folders = DriveApp.getFoldersByName('SEED CARD バックアップ');
  var folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder('SEED CARD バックアップ');
  var today   = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var copyName = 'SEED CARD バックアップ_' + today;
  var existing = folder.getFilesByName(copyName);
  if (existing.hasNext()) { return; }
  DriveApp.getFileById(ss.getId()).makeCopy(copyName, folder);
  var files = []; var iter = folder.getFiles();
  while (iter.hasNext()) { var f = iter.next(); files.push({file:f, date:f.getDateCreated()}); }
  files.sort(function(a,b){return b.date-a.date;});
  for (var i = 12; i < files.length; i++) { files[i].file.setTrashed(true); }
}

function setupWeeklyBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'weeklyBackup') { ScriptApp.deleteTrigger(t); }
  });
  ScriptApp.newTrigger('weeklyBackup').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3).create();
}
