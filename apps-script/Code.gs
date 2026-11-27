/**
 * Bonus Track server. Deploy as owner after configuring the linked private sheet.
 * Public functions validate all input. Functions ending "_" are editor-only helpers.
 */
const POLICY_VERSION = '2026-11-27-direct-v1';
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
  if (existing.some((value, index) => value.replace(/\s/g, '').replace('받을사람', '받는사람') !== HEADERS[index].replace(/\s/g, ''))) {
    throw new Error('첫 9개 항목을 확인해주세요. 기존 내용을 덮어쓰지 않았습니다.');
  }
  sheet.getRange(1, 10, 1, 4).setValues([HEADERS.slice(9)]);
  sheet.setFrozenRows(1);
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
  const recipient = other ? text_(payload.recipient, 50, '받는 사람') : sender;
  const recipientEmail = other ? email_(payload.recipientEmail) : email;
  return locked_(() => {
    const sheet = sheet_();
    const existing = findRow_(sheet, 10, id);
    const receipt = { id: id, status: '발송 대기' };
    if (existing) {
      const row = sheet.getRange(existing, 1, 1, 13).getValues()[0];
      if (row[2] !== email || row[1] !== sender || row[5] !== letter || row[10] !== payload.audience || row[3] !== recipient || row[4] !== recipientEmail) {
        throw new Error('이미 접수된 편지입니다. 수정은 문의 이메일로 요청해주세요.');
      }
      receipt.status = row[7];
      return receipt;
    }
    const now = timestamp_();
    const metadata = {
      version: POLICY_VERSION,
      sender: { privacy: true, delivery: true, sharing: other, at: now, emailVerified: false },
      recipient: other ? null : { privacy: true, delivery: true, at: now, emailVerified: false },
      recipientProvidedBy: 'sender',
      recipientConsentCollected: !other
    };
    const values = [
      now, sender, email, recipient, recipientEmail,
      letter, '2027-12', receipt.status, '', id, payload.audience,
      '', JSON.stringify(metadata)
    ].map(safeCell_);
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, values.length).setValues([values]);
    SpreadsheetApp.flush();
    return receipt;
  });
}

function getServiceState() {
  return { accepting: PropertiesService.getScriptProperties().getProperty('ACCEPTING') === 'true' };
}

/** Run from the editor to open or pause receipts after deployment checks. */
function openReceipts() {
  if (!SpreadsheetApp.getActiveSpreadsheet()) throw new Error('편지함에 연결된 편집기에서 실행해주세요.');
  PropertiesService.getScriptProperties().setProperty('ACCEPTING', 'true');
}
function pauseReceipts() {
  if (!SpreadsheetApp.getActiveSpreadsheet()) throw new Error('편지함에 연결된 편집기에서 실행해주세요.');
  PropertiesService.getScriptProperties().setProperty('ACCEPTING', 'false');
}
