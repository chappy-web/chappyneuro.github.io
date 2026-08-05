/**
 * ============================================================================
 *  LatentBrain 홈페이지 문의 폼 백엔드  (Google Apps Script)
 *  ---------------------------------------------------------------------------
 *  기능 : 1) 문의 내용을 구글 시트에 자동 저장
 *         2) moonhj@lat-brain.com 으로 문의 알림 메일 발송
 *         3) 문의한 방문자에게 "접수 완료" 자동 회신 메일 발송
 *  비용 : 무료 (Google 계정만 있으면 됨, 별도 서비스 가입 불필요)
 * ============================================================================
 *
 *  ■ 설치 방법 (딱 한 번, 약 10분)
 *
 *  1) 구글 시트 새로 만들기
 *     - https://sheets.new  접속 → 시트 이름을 "LatentBrain 문의접수" 등으로 변경
 *
 *  2) Apps Script 편집기 열기
 *     - 시트 상단 메뉴 [확장 프로그램] → [Apps Script] 클릭
 *
 *  3) 코드 붙여넣기
 *     - 편집기에 있던 기본 코드(function myFunction...)를 전부 지우고
 *       이 파일의 내용을 전부 복사해서 붙여넣기 → 저장(Ctrl+S)
 *
 *  4) 배포하기
 *     - 우측 상단 [배포] → [새 배포]
 *     - 톱니바퀴 아이콘 → [웹 앱] 선택
 *     - 설명       : 문의폼 v1  (아무 문구나 가능)
 *     - 실행 사용자 : 나 (본인 계정)
 *     - 액세스 권한 : ★ 모든 사용자 ★   ← 반드시 이걸로 선택
 *     - [배포] 클릭 → 권한 승인 화면이 뜨면
 *         [고급] → [(프로젝트명)(으)로 이동] → [허용]
 *       ※ "이 앱은 확인되지 않았습니다" 경고는 본인이 만든 스크립트라 정상입니다.
 *
 *  5) 웹 앱 URL 복사
 *     - 배포 완료 화면의 "웹 앱 URL" (https://script.google.com/macros/s/.../exec)
 *     - index-lat.html 에서  CONTACT_ENDPOINT = "여기에_APPS_SCRIPT_웹앱_URL_붙여넣기"
 *       부분의 따옴표 안에 붙여넣기
 *
 *  ■ 코드를 수정한 뒤에는
 *     [배포] → [배포 관리] → 연필(수정) → 버전 [새 버전] → [배포]
 *     ※ [새 배포]로 다시 만들면 URL이 바뀌니 주의하세요.
 *
 *  ■ 발송 한도 (Google 무료 할당량)
 *     - 일반 Gmail 계정      : 하루 100통
 *     - Google Workspace 계정 : 하루 1,500통
 *     문의 1건당 2통(알림+자동회신)이 나가므로 실질 한도는 그 절반입니다.
 * ============================================================================
 */

// ───────── 설정값 (이 부분만 필요에 따라 수정) ─────────
const ADMIN_EMAIL  = 'moonhj@lat-brain.com';   // 문의 알림을 받을 주소
const SENDER_NAME  = 'LatentBrain';            // 자동회신 메일에 표시될 발신자명
const REPLY_SUBJECT = '[LatentBrain] 문의가 정상적으로 접수되었습니다';
// ──────────────────────────────────────────────────────


/** 폼 전송(POST)을 처리하는 메인 함수 */
function doPost(e) {
  try {
    const p = (e && e.parameter) ? e.parameter : {};

    // 봇 차단용 honeypot 필드가 채워져 있으면 조용히 무시
    if (p.botcheck) {
      return jsonOut({ result: 'success' });
    }

    const name    = (p.name    || '').toString().trim();
    const email   = (p.email   || '').toString().trim();
    const company = (p.company || '').toString().trim();
    const type    = (p.inquiry_type || '').toString().trim();
    const message = (p.message || '').toString().trim();

    if (!name || !email || !message) {
      return jsonOut({ result: 'error', message: '필수 항목이 비어 있습니다.' });
    }

    const now = new Date();

    // 1) 구글 시트에 저장
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['접수일시', '이름', '소속', '이메일', '문의유형', '문의내용']);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    }
    sheet.appendRow([now, name, company, email, type, message]);

    // 2) 관리자에게 문의 알림
    const adminBody =
      '홈페이지를 통해 새로운 문의가 접수되었습니다.\n\n' +
      '─────────────────────────────\n' +
      '· 접수일시 : ' + Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm') + '\n' +
      '· 이름     : ' + name + '\n' +
      '· 소속     : ' + (company || '-') + '\n' +
      '· 이메일   : ' + email + '\n' +
      '· 문의유형 : ' + (type || '-') + '\n' +
      '─────────────────────────────\n\n' +
      message + '\n\n' +
      '※ 이 메일에 그대로 [답장]하면 문의자에게 회신됩니다.';

    MailApp.sendEmail({
      to:       ADMIN_EMAIL,
      subject:  (p.subject || '[LatentBrain 홈페이지] 새 문의가 도착했습니다') + ' - ' + name,
      body:     adminBody,
      replyTo:  email,          // 답장하면 바로 문의자에게 감
      name:     SENDER_NAME
    });

    // 3) 문의자에게 접수 확인 자동회신
    const replyBody =
      name + '님, 안녕하세요.\n\n' +
      'LatentBrain 홈페이지를 통해 보내주신 문의가 정상적으로 접수되었습니다.\n' +
      '담당자가 내용을 확인한 뒤 영업일 기준 1~2일 이내에 회신드리겠습니다.\n\n' +
      '─── 보내주신 내용 ───\n' +
      '· 문의유형 : ' + (type || '-') + '\n' +
      '· 접수일시 : ' + Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm') + '\n\n' +
      message + '\n' +
      '──────────────────\n\n' +
      '문의해 주셔서 감사합니다.\n\n' +
      'LatentBrain 드림\n' +
      ADMIN_EMAIL + '\n' +
      '충청남도 천안시 동남구 백석대학로 1-9\n\n' +
      '※ 본 메일은 발신 전용이 아니며, 그대로 답장하셔도 확인 가능합니다.';

    MailApp.sendEmail({
      to:      email,
      subject: REPLY_SUBJECT,
      body:    replyBody,
      replyTo: ADMIN_EMAIL,
      name:    SENDER_NAME
    });

    return jsonOut({ result: 'success' });

  } catch (err) {
    return jsonOut({ result: 'error', message: String(err) });
  }
}


/** 브라우저에서 URL을 그냥 열었을 때 배포 확인용 */
function doGet() {
  return jsonOut({ result: 'ready', message: 'LatentBrain contact endpoint is running.' });
}


/** JSON 응답 헬퍼 */
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * 배포 전 테스트용 함수.
 * 편집기 상단에서 이 함수를 선택하고 [실행]하면
 * 관리자 메일과 자동회신이 정상 발송되는지 확인할 수 있습니다.
 */
function testSend() {
  const fake = {
    parameter: {
      name: '테스트',
      email: ADMIN_EMAIL,          // 본인에게 자동회신도 오도록 설정
      company: '테스트기관',
      inquiry_type: '기타 / Other',
      message: '테스트 문의입니다.'
    }
  };
  Logger.log(doPost(fake).getContent());
}
