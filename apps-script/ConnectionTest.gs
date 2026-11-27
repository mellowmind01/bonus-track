/** Editor-only future-self receipt checks. */
function testReceiptConnection() {
  if (!SpreadsheetApp.getActiveSpreadsheet()) throw new Error('편집기에서만 실행할 수 있습니다.');
  const sheet = sheet_();
  const payload = {
    id: Utilities.getUuid(), audience: 'self', sender: '[테스트] 미래의 나',
    email: 'self-test@example.com', letter: '=1+1\n[테스트] 미래의 나에게',
    policyVersion: POLICY_VERSION, privacy: true, delivery: true
  };
  function expectRefusal(value) {
    let refused = false;
    try { submitLetter(value); } catch (error) { refused = true; }
    if (!refused) throw new Error('제한해야 할 접수가 허용됐습니다.');
  }
  let row = 0;
  try {
    expectRefusal(Object.assign({}, payload, { privacy: false }));
    expectRefusal(Object.assign({}, payload, { email: 'invalid' }));
    expectRefusal(Object.assign({}, payload, {
      audience: 'other', recipient: '다른 사람', recipientEmail: 'other@example.com', sharing: true
    }));
    // Extra recipient fields can never redirect a future-self letter.
    const receipt = submitLetter(Object.assign({}, payload, { recipientEmail: 'other@example.com' }));
    row = findRow_(sheet, 10, payload.id);
    const values = sheet.getRange(row, 1, 1, 13).getValues()[0];
    if (receipt.status !== '발송 대기' || values[3] !== payload.sender ||
        values[4] !== payload.email || values[10] !== 'self') {
      throw new Error('자기 편지의 받는 사람 정보 저장 실패');
    }
    submitLetter(payload);
    if (sheet.getRange(2, 10, sheet.getLastRow() - 1, 1).createTextFinder(payload.id)
      .matchEntireCell(true).findAll().length !== 1) throw new Error('중복 방지 실패');
    if (sheet.getRange(row, 6).getFormula()) throw new Error('수식 차단 실패');
    const metadata = JSON.parse(values[12]);
    if (!metadata.sender.at || metadata.sender.emailVerified !== false || !metadata.recipientConsentCollected) {
      throw new Error('동의 기록 실패');
    }
    console.log('자기 편지 직접 접수, 타인 편지 거부, 필수 동의, 이메일 형식, 중복 방지, 수식 차단 확인 완료');
  } finally {
    if (row) sheet.getRange(row, 8).setValue('테스트 완료 · 발송 제외');
  }
}
