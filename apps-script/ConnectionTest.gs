/**
 * Editor-only integration check. Keeps labelled test rows with a non-sendable status.
 */
function testReceiptConnection() {
  if (!SpreadsheetApp.getActiveSpreadsheet()) throw new Error('편집기에서만 실행할 수 있습니다.');
  const payload = {
    id: Utilities.getUuid(), audience: 'other', sender: '[테스트] 소중한 사람',
    email: 'test-sender@example.com', letter: '=1+1\n[테스트] 수식으로 실행되면 안 됩니다.',
    policyVersion: POLICY_VERSION, privacy: true, delivery: true, sharing: true, code: '123456',
    purpose: 'sender'
  };
  const key = verificationKey_(payload);
  CacheService.getScriptCache().put(key, JSON.stringify({ digest: hash_('123456'), attempts: 0 }), 600);
  const sheet = sheet_();
  let row = 0;
  try {
    const receipt = submitLetter(payload);
    row = findRow_(sheet, 10, payload.id);
    if (!row || receipt.status !== '수신 동의 대기') throw new Error('다른 사람 편지 저장 실패');
    const duplicate = submitLetter(payload);
    if (duplicate.id !== receipt.id || sheet.getLastRow() !== row) throw new Error('중복 접수 처리 실패');
    if (sheet.getRange(row, 6).getFormula()) throw new Error('편지 수식 실행 차단 실패');
    const invitation = getInvitation(receipt.invitationToken);
    if (invitation.sender !== payload.sender || 'letter' in invitation) throw new Error('동의 링크 정보 제한 실패');
    const recipient = {
      token: receipt.invitationToken, name: '[테스트] 받는 분', email: 'test-recipient@example.com',
      purpose: 'recipient', policyVersion: POLICY_VERSION, privacy: true, delivery: true, code: '654321'
    };
    let refused = false;
    try {
      acceptInvitation(Object.assign({}, recipient, { privacy: false }));
    } catch (error) { refused = true; }
    if (!refused) throw new Error('미동의 차단 실패');
    CacheService.getScriptCache().put(verificationKey_(recipient),
      JSON.stringify({ digest: hash_('654321'), attempts: 0 }), 600);
    acceptInvitation(recipient);
    acceptInvitation(recipient);
    if (sheet.getRange(row, 8).getDisplayValue() !== '발송 대기') throw new Error('수신 동의 저장 실패');
    const metadata = JSON.parse(sheet.getRange(row, 13).getDisplayValue());
    if (!metadata.recipient.emailVerifiedAt || !metadata.sender.at) throw new Error('동의 기록 실패');
    console.log('다른 사람 편지 저장 / 중복 방지 / 수식 차단 / 동의 누락 차단 / 수신 동의 기록 확인 완료');
  } finally {
    if (row) {
      sheet.getRange(row, 8).setValue('테스트 완료 · 발송 제외');
      sheet.getRange(row, 12).clearContent();
    }
    CacheService.getScriptCache().remove(key);
  }
}
