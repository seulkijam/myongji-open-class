/**
 * 명지중 수업 나눔 참관 신청 → Google 시트 저장
 *
 * 설치: 신청을 받을 Google 시트에서 [확장 프로그램 > Apps Script]를 열고
 *       이 파일 내용을 그대로 붙여 넣은 뒤 웹 앱으로 배포합니다. (작업메모.md 참고)
 */

const SHEET_NAME = '참관 신청';

// 사이트(index.html)의 수업 id와 같아야 합니다.
const LESSONS = [
  ['info', '정보 5교시'],
  ['tech', '기술 5교시'],
  ['eng', '영어 5교시'],
  ['math', '수학 6교시'],
  ['rel', '종교 6교시'],
  ['share', '사례 나눔 7교시']
];

const OFFICES = ['동부', '서부', '남부', '북부', '중부', '강남서초', '강동송파', '강서양천', '동작관악', '성동광진', '성북강북'];

const HEADERS = ['신청 일시', '지원청', '학교', '이름', '휴대폰 번호', '이메일', '참관 수업']
  .concat(LESSONS.map(l => l[1] + ' (인원)'))
  .concat(['개인정보 동의']);

const PHONE_COL = 5; // '휴대폰 번호' 열 (A=1)

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);

    // 자동 입력(스팸)은 저장하지 않고 성공처럼 응답
    if (d.website) return json({ ok: true });

    const text = (v, max) => String(v == null ? '' : v).trim().slice(0, max);
    const office = text(d.office, 10);
    const school = text(d.school, 40);
    const name = text(d.name, 20);
    const phone = String(d.phone || '').replace(/\D/g, '');
    const email = text(d.email, 80);
    const picked = Array.isArray(d.lessons) ? d.lessons : [];

    if (OFFICES.indexOf(office) < 0) return json({ ok: false, message: '지원청을 선택해 주세요.' });
    if (!school || !name) return json({ ok: false, message: '학교와 이름을 입력해 주세요.' });
    if (!/^01[016789]\d{7,8}$/.test(phone)) return json({ ok: false, message: '휴대폰 번호를 확인해 주세요.' });
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return json({ ok: false, message: '이메일 주소를 확인해 주세요.' });
    const chosen = LESSONS.filter(l => picked.indexOf(l[0]) >= 0);
    if (!chosen.length) return json({ ok: false, message: '참관 희망 수업을 골라 주세요.' });
    if (d.agree !== true) return json({ ok: false, message: '개인정보 수집·이용에 동의해 주세요.' });

    const row = [
      new Date(), office, school, name, phone, email,
      chosen.map(l => l[1]).join(', ')
    ]
      .concat(LESSONS.map(l => (picked.indexOf(l[0]) >= 0 ? 1 : '')))
      .concat(['동의']);

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = getSheet();
      const r = sheet.getLastRow() + 1;
      // 휴대폰 번호 앞자리 0이 사라지지 않도록 텍스트 서식으로 저장
      sheet.getRange(r, PHONE_COL).setNumberFormat('@');
      sheet.getRange(r, 1, 1, row.length).setValues([row]);
    } finally {
      lock.releaseLock();
    }
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, message: '접수 중 오류가 생겼습니다. 잠시 후 다시 시도해 주세요.' });
  }
}

// 주소를 브라우저로 열었을 때 동작 확인용
function doGet() {
  return json({ ok: true, message: '참관 신청 접수 중' });
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#f1e2cf');
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange(2, PHONE_COL, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
  }
  return sheet;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
