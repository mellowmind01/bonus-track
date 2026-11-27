/** Editor-only checks. Test records are marked as excluded from delivery. */
function testReceiptConnection() {
  if (!SpreadsheetApp.getActiveSpreadsheet()) throw new Error('편집기에서만 실행할 수 있습니다.');
  const sheet = sheet_();
  const testRows = [];
  const base = {
    id: Utilities.getUuid(), audience: 'other', sender: '[테스트] 간편 접수',
    email: 'test-sender@example.com', recipient: '[테스트] 받는 분',
    recipientEmail: 'test-recipient@example.com',
    letter: '=1+1\n[테스트] 이메일 인증 없는 접수',
    policyVersion: POLICY_VERSION, privacy: true, delivery: true, sharing: true
  };
  function expectRefusal(payload) {
    let refused = false;
    try { submitLetter(payload); } catch (error) { refused = true; }
    if (!refused) throw new Error('잘못된 입력 또는 미동의가 거부되지 않았습니다.');
  }
  try {
    expectRefusal(Object.assign({}, base, { privacy: false }));
    expectRefusal(Object.assign({}, base, { sharing: false }));
    expectRefusal(Object.assign({}, base, { recipientEmail: 'invalid' }));
    const receipt = submitLetter(base);
    const row = findRow_(sheet, 10, base.id);
    testRows.push(row);
    const stored = sheet.getRange(row, 1, 1, 13).getValues()[0];
    if (receipt.status !== '발송 대기' || stored[3] !== base.recipient ||
        stored[4] !== base.recipientEmail) throw new Error('직접 수신자 정보 저장 실패');
    const duplicate = submitLetter(base);
    if (duplicate.id !== receipt.id ||
        sheet.getRange(2, 10, sheet.getLastRow() - 1, 1).createTextFinder(base.id)
          .matchEntireCell(true).findAll().length !== 1) throw new Error('중복 방지 실패');
    if (sheet.getRange(row, 6).getFormula()) throw new Error('편지가 수식으로 실행됐습니다.');
    const metadata = JSON.parse(stored[12]);
    if (!metadata.sender.at || metadata.sender.emailVerified !== false ||
        metadata.recipientConsentCollected !== false || metadata.recipient !== null) {
      throw new Error('실제 확인하지 않은 동의나 이메일 인증이 기록됐습니다.');
    }
    const self = Object.assign({}, base, {
      id: Utilities.getUuid(), audience: 'self',
      recipient: '', recipientEmail: '', sharing: false, letter: '[테스트] 나에게 간편 접수'
    });
    submitLetter(self);
    const selfRow = findRow_(sheet, 10, self.id);
    testRows.push(selfRow);
    if (sheet.getRange(selfRow, 5).getDisplayValue() !== self.email) {
      throw new Error('나에게 편지 저장 실패');
    }
    console.log('확인 번호 없는 자기 편지·다른 사람 편지 접수, 필수 동의, 이메일 형식, 중복 방지, 동의 기록 확인 완료');
  } finally {
    for (const row of testRows) {
      if (row) sheet.getRange(row, 8).setValue('테스트 완료 · 발송 제외');
    }
  }
}
