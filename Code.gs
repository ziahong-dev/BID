/**
 * BID 설문 응답 수집 백엔드 (Google Apps Script)
 * ─ 설문 사이트가 보낸 응답(JSON)을 받아 구글 시트에 한 줄씩 저장합니다.
 * ─ 시트는 구글 드라이브 안의 파일이므로, 응답이 곧 드라이브에 쌓입니다.
 *
 * 사용법
 *  1) 드라이브에서 빈 구글 시트를 하나 만들고, URL의 /d/ 와 /edit 사이 문자열(시트 ID)을 복사
 *  2) script.google.com → 새 프로젝트 → 이 코드 전체 붙여넣기
 *  3) 아래 SHEET_ID 값에 복사한 시트 ID를 붙여넣기
 *  4) 배포 → 새 배포 → 유형: 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → 배포
 *  5) 발급된 /exec URL을 설문 HTML의 ENDPOINT 에 붙여넣기
 */

const SHEET_ID   = 'YOUR_SHEET_ID_HERE';   // ← 1번에서 복사한 시트 ID로 교체
const SHEET_NAME = 'responses';
const SECRET     = '';                     // (선택) 스팸 방지용 토큰. 쓰려면 HTML에도 같은 값 추가

// 설문 HTML 의 상태 키와 동일한 순서 — 디자인을 바꿔도 이 키들만 유지하면 됩니다.
const FIELDS = ['_id','_ts','q1','q2','q3','q4','q4_other','q5','q5_other','q6','q6_other',
  'q7','q8','q8_other','q9','q9_other','q10_rank1','q10_rank2','q10_rank3','q10_rank4','q10_rank5',
  'q11','q12','q13','q14'];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (SECRET && data._token !== SECRET) {
      return json({ ok: false, error: 'unauthorized' });
    }
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) sh.appendRow(FIELDS);   // 헤더 자동 생성

    const row = FIELDS.map(function (f) {
      if (f === '_ts') return data._ts || new Date().toISOString();
      if (f.indexOf('q10_rank') === 0) {
        const idx = parseInt(f.slice(8), 10) - 1;      // q10_rank1 → index 0
        return (data.q10 && data.q10[idx]) ? data.q10[idx] : '';
      }
      const v = data[f];
      return Array.isArray(v) ? v.join('; ') : (v == null ? '' : v);
    });
    sh.appendRow(row);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// 배포 확인용 (브라우저로 /exec 열면 표시)
function doGet() {
  return ContentService.createTextOutput('BID survey endpoint is running.');
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
