/**
 * Bonus Track server. Deploy as owner after configuring the linked private sheet.
 * Public functions validate all input. Functions ending "_" are editor-only helpers.
 */
const POLICY_VERSION = '2026-11-27';
const HEADERS = [
  '접수일', '보내는 사람', '작성자 이메일', '받는 사람', '받을 이메일',
  '편지', '발송 예정일', '발송 상태', '실제 발송일',
  '접수번호', '편지 유형', '수신 동의 링크 해시', '동의 기록'
];

function setupProject() {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('편지함에서 연결된 프로젝트를 열어주세요.');
  properties.setProperty('SHEET_ID', spreadsheet.getId());
  if (!properties.getProperty('TOKEN_SECRET')) {
    properties.setProperty('TOKEN_SECRET', Utilities.getUuid() + Utilities.getUuid());
  }
  if (!properties.getProperty('ACCEPTING')) properties.setProperty('ACCEPTING', 'false');
  const sheet = spreadsheet.getSheetByName('편지접수');
  if (!sheet) throw new Error('편지접수 시트를 찾을 수 없습니다.');
  const existing = sheet.getRange(1, 1, 1, 9).getDisplayValues()[0];
  if (existing.some((value, index) => value.replace(/\s/g, '') !== HEADERS[index].replace(/\s/g, ''))) {
    throw new Error('첫 9개 항목을 확인해주세요. 기존 내용을 덮어쓰지 않았습니다.');
  }
  sheet.getRange(1, 10, 1, 4).setValues([HEADERS.slice(9)]);
  sheet.setFrozenRows(1);
  // Requests only email-send permission, never inbox-reading permission.
  MailApp.getRemainingDailyQuota();
  return '저장 연결 준비 완료. ACCEPTING=false 상태입니다.';
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('보너스트랙 | 말랑한 마음')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function assertOpen_() {
  if (PropertiesService.getScriptProperties().getProperty('ACCEPTING') !== 'true') {
    throw new Error('접수 준비 중입니다. 잠시 후 다시 확인해주세요.');
  }
}

function sheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('편지함 연결을 확인해주세요.');
  return SpreadsheetApp.openById(id).getSheetByName('편지접수');
}

function text_(value, limit, label) {
  if (typeof value !== 'string' || !value.trim() || value.length > limit) {
    throw new Error(label + '을 확인해주세요.');
  }
  return value.trim();
}

function email_(value) {
  const email = text_(value, 254, '이메일').toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || /[\r\n]/.test(email)) {
    throw new Error('이메일 주소를 확인해주세요.');
  }
  return email;
}

function id_(value) {
  if (typeof value !== 'string' || !/^[a-f0-9-]{36}$/.test(value)) {
    throw new Error('화면을 새로 열고 다시 시도해주세요.');
  }
  return value;
}

function hash_(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value)
    .map(byte => ('0' + ((byte + 256) % 256).toString(16)).slice(-2)).join('');
}

function invitation_(id) {
  const secret = PropertiesService.getScriptProperties().getProperty('TOKEN_SECRET');
  return Utilities.computeHmacSha256Signature(id, secret)
    .map(byte => ('0' + ((byte + 256) % 256).toString(16)).slice(-2)).join('');
}

function locked_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function findRow_(sheet, column, value) {
  if (sheet.getLastRow() < 2) return 0;
  const cell = sheet.getRange(2, column, sheet.getLastRow() - 1, 1)
    .createTextFinder(value).matchEntireCell(true).useRegularExpression(false).findNext();
  return cell ? cell.getRow() : 0;
}

function invitationRow_(sheet, token) {
  if (typeof token !== 'string' || !/^[a-f0-9]{64}$/.test(token)) {
    throw new Error('동의 요청 링크를 확인해주세요.');
  }
  const row = findRow_(sheet, 12, hash_(token));
  if (!row) throw new Error('유효한 동의 요청을 찾을 수 없습니다.');
  return row;
}

function timestamp_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function safeCell_(value) {
  // Literal strings prevent spreadsheet formulas from untrusted free text.
  return "'" + value;
}

function requireConsent_(payload) {
  if (!payload || payload.policyVersion !== POLICY_VERSION ||
      payload.privacy !== true || payload.delivery !== true) {
    throw new Error('최신 개인정보 안내와 발송 방식에 동의해주세요.');
  }
}

function verificationKey_(payload) {
  const purpose = payload.purpose === 'recipient' ? 'recipient' : 'sender';
  const resource = purpose === 'recipient' ? text_(payload.token, 64, '링크') : id_(payload.id);
  return 'verify:' + hash_(purpose + '|' + resource + '|' + email_(payload.email));
}

function requestEmailCode(payload) {
  assertOpen_();
  requireConsent_(payload);
  return locked_(() => {
    if (payload.purpose === 'recipient') {
      const sheet = sheet_();
      const row = invitationRow_(sheet, payload.token);
      if (sheet.getRange(row, 8).getDisplayValue() !== '수신 동의 대기') {
        throw new Error('이미 처리된 동의 요청입니다.');
      }
    }
    const cache = CacheService.getScriptCache();
    const key = verificationKey_(payload);
    if (cache.get(key + ':cooldown')) throw new Error('1분 후 다시 요청해주세요.');
    const countKey = 'daily:' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd');
    const count = Number(cache.get(countKey) || 0);
    if (count >= 80 || MailApp.getRemainingDailyQuota() < 1) {
      throw new Error('오늘의 이메일 확인 요청이 많습니다. 문의 이메일로 연락해주세요.');
    }
    const code = String(parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 10), 16) % 1000000)
      .padStart(6, '0');
    cache.put(key, JSON.stringify({ digest: hash_(code), attempts: 0 }), 600);
    cache.put(key + ':cooldown', '1', 60);
    cache.put(countKey, String(count + 1), 86400);
    MailApp.sendEmail({
      to: email_(payload.email),
      subject: '[말랑한 마음] 보너스트랙 이메일 확인',
      body: '요청하신 이메일 확인 번호는 ' + code +
        '입니다. 10분 안에 입력해주세요.\n\n본인이 요청하지 않았다면 무시하셔도 됩니다.\n문의: mellowmind.official@gmail.com',
      name: '말랑한 마음',
      replyTo: 'mellowmind.official@gmail.com'
    });
    return { expiresIn: 600 };
  });
}

function verifyCode_(payload) {
  const key = verificationKey_(payload);
  const cache = CacheService.getScriptCache();
  const raw = cache.get(key);
  if (!raw) throw new Error('이메일 확인 번호를 다시 요청해주세요.');
  const challenge = JSON.parse(raw);
  if (challenge.attempts >= 5) throw new Error('확인 횟수를 초과했습니다. 번호를 다시 요청해주세요.');
  challenge.attempts += 1;
  cache.put(key, JSON.stringify(challenge), 600);
  if (hash_(String(payload.code || '')) !== challenge.digest) {
    throw new Error('이메일 확인 번호가 맞지 않습니다.');
  }
  return key;
}

function submitLetter(payload) {
  assertOpen_();
  requireConsent_(payload);
  const id = id_(payload.id);
  const email = email_(payload.email);
  const sender = text_(payload.sender, 50, '보내는 사람');
  const letter = text_(payload.letter, 5000, '편지');
  const other = payload.audience === 'other';
  if (!other && payload.audience !== 'self') throw new Error('받는 대상을 선택해주세요.');
  if (other && payload.sharing !== true) throw new Error('받는 분에게 전달되는 정보에 동의해주세요.');
  return locked_(() => {
    const sheet = sheet_();
    const existing = findRow_(sheet, 10, id);
    const token = other ? invitation_(id) : '';
    const receipt = { id: id, invitationToken: token, status: other ? '수신 동의 대기' : '발송 대기' };
    if (existing) {
      const row = sheet.getRange(existing, 1, 1, 13).getValues()[0];
      if (row[2] !== email || row[1] !== sender || row[5] !== letter || row[10] !== payload.audience) {
        throw new Error('이미 접수된 편지입니다. 수정은 문의 이메일로 요청해주세요.');
      }
      return receipt;
    }
    const key = verifyCode_(Object.assign({}, payload, { purpose: 'sender' }));
    const now = timestamp_();
    const metadata = {
      version: POLICY_VERSION,
      sender: { privacy: true, delivery: true, sharing: other, at: now, emailVerifiedAt: now },
      recipient: other ? null : { privacy: true, delivery: true, at: now, emailVerifiedAt: now }
    };
    const values = [
      now, sender, email, other ? '' : sender, other ? '' : email,
      letter, '2027-12', receipt.status, '', id, payload.audience,
      other ? hash_(token) : '', JSON.stringify(metadata)
    ].map(safeCell_);
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, values.length).setValues([values]);
    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove(key);
    return receipt;
  });
}

function getInvitation(token) {
  assertOpen_();
  const sheet = sheet_();
  const row = invitationRow_(sheet, token);
  if (sheet.getRange(row, 8).getDisplayValue() !== '수신 동의 대기') {
    throw new Error('이미 처리된 동의 요청입니다.');
  }
  return { sender: sheet.getRange(row, 2).getDisplayValue(), policyVersion: POLICY_VERSION };
}

function acceptInvitation(payload) {
  assertOpen_();
  requireConsent_(payload);
  const name = text_(payload.name, 50, '받는 사람');
  const email = email_(payload.email);
  return locked_(() => {
    const sheet = sheet_();
    const row = invitationRow_(sheet, payload.token);
    const status = sheet.getRange(row, 8).getDisplayValue();
    if (status !== '수신 동의 대기') {
      if (status === '발송 대기' && sheet.getRange(row, 5).getDisplayValue() === email) {
        return { status: '발송 대기' };
      }
      throw new Error('이미 처리된 동의 요청입니다.');
    }
    const key = verifyCode_(Object.assign({}, payload, { purpose: 'recipient' }));
    const metadata = JSON.parse(sheet.getRange(row, 13).getDisplayValue());
    const now = timestamp_();
    metadata.recipient = { privacy: true, delivery: true, at: now, emailVerifiedAt: now, version: POLICY_VERSION };
    const values = sheet.getRange(row, 1, 1, 13).getValues()[0];
    values[3] = name;
    values[4] = email;
    values[7] = '발송 대기';
    values[12] = JSON.stringify(metadata);
    sheet.getRange(row, 1, 1, 13).setValues([values.map(value => safeCell_(String(value)))]);
    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove(key);
    return { status: '발송 대기' };
  });
}
