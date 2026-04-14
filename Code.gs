/**
 * ============================================================
 * THE AI INSIGHT - Newsletter Backend (Google Apps Script)
 * ============================================================
 * Google Sheets를 DB로 사용하는 뉴스레터 구독 관리 시스템
 *
 * 시트 구조:
 *   [subscribers] email | name | subscribed_at | status | source | unsubscribed_at
 *   [newsletters] id | title | date | summary | content_url | sent_count
 *   [config]      key | value
 *   [logs]        timestamp | action | detail
 */

// ─── 설정 ───────────────────────────────────────────────────
// [보안] 비밀번호와 API 키는 PropertiesService에서 로드
// 최초 1회 setupSecureProperties() 실행 필요 (아래 함수 참고)
const CONFIG = {
  SPREADSHEET_ID: '1XpalUZ3ap_97U4VPwCPD5603o107WAS9-B6ebxSPtFc',
  PUBLIC_SHEET_ID: '1bPc6XBR6qo7eGFUybuSNp8-fmeIFWT55PJdnE4LQ_QM',
  get ADMIN_PASSWORD() {
    return PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || 'CHANGE_ME';
  },
  INTERNAL_EMAILS: [
    // 고정 명단 없음 — 모든 구독자는 구독 요청 메일을 통해 자발적으로 구독
  ],
  EXTERNAL_MAILER: {
    provider: 'gmail',
    woorimail: {
      get authkey() { return PropertiesService.getScriptProperties().getProperty('WOORIMAIL_AUTHKEY') || ''; },
      domain: 'seoyoneh.com',
      senderEmail: 'aa200526@seoyoneh.com',
      senderName: 'THE AI INSIGHT'
    }
  }
};

/**
 * [최초 1회 실행] 비밀번호와 API 키를 PropertiesService에 안전하게 저장
 * Apps Script 편집기에서 이 함수를 한 번 실행한 뒤, 아래 값들을 원하는 비밀번호로 변경하세요.
 */
/**
 * [최초 1회 실행] 비밀번호와 API 키를 PropertiesService에 안전하게 저장
 * ★ 보안 주의: 실행 후 아래 값들을 소스에서 삭제하거나 빈 문자열로 교체하세요.
 *   이미 실행 완료된 상태이므로, 값은 PropertiesService에 안전하게 저장되어 있습니다.
 */
function setupSecureProperties() {
  var props = PropertiesService.getScriptProperties();
  // ★ 이미 설정 완료됨 — 평문 노출 방지를 위해 값을 제거합니다.
  // 비밀번호 변경이 필요하면 아래 빈 문자열에 새 값을 넣고 실행 후 다시 비워주세요.
  var newPassword = '';       // ← 변경 시에만 입력
  var newAuthKey = '';        // ← 변경 시에만 입력

  if (newPassword) props.setProperty('ADMIN_PASSWORD', newPassword);
  if (newAuthKey) props.setProperty('WOORIMAIL_AUTHKEY', newAuthKey);

  if (!newPassword && !newAuthKey) {
    addLog('SECURITY', 'setupSecureProperties 호출됨 — 변경할 값 없음 (이미 설정 완료)');
    return;
  }
  addLog('SECURITY', 'PropertiesService 보안 키 업데이트 완료');
}

// ─── 스프레드시트 헬퍼 ────────────────────────────────────────
function getSheet(name) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initSheet(sheet, name);
  }
  return sheet;
}

function initSheet(sheet, name) {
  const headers = {
    subscribers: ['email', 'name', 'subscribed_at', 'status', 'source', 'unsubscribed_at'],
    newsletters: ['id', 'title', 'date', 'summary', 'content_url', 'sent_count'],
    config: ['key', 'value'],
    logs: ['timestamp', 'action', 'detail']
  };
  if (headers[name]) {
    sheet.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
    sheet.getRange(1, 1, 1, headers[name].length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  // 샘플 뉴스레터 데이터 삽입
  if (name === 'newsletters') {
    const samples = [
      ['NL-2026-001', 'AI Agent 시대의 업무 자동화 전략', '2026-04-01', 'Anthropic Claude, OpenAI Operator 등 AI Agent가 기업 업무를 어떻게 바꾸고 있는지 분석합니다.', '', 6],
      ['NL-2026-002', 'RAG vs Fine-tuning: 우리 회사에 맞는 AI 도입 방법은?', '2026-03-15', '사내 데이터 활용을 위한 RAG와 Fine-tuning 비교 가이드. 비용, 성능, 유지보수 관점에서 정리했습니다.', '', 6],
      ['NL-2026-003', 'MCP(Model Context Protocol)로 AI 도구 통합하기', '2026-03-01', 'Anthropic의 MCP 프로토콜을 활용해 사내 시스템과 AI를 연결하는 실전 가이드입니다.', '', 6]
    ];
    sheet.getRange(2, 1, samples.length, samples[0].length).setValues(samples);
  }
}

function addLog(action, detail) {
  try {
    const sheet = getSheet('logs');
    // [보안] 로그에 이메일 마스킹 적용
    var maskedDetail = String(detail || '').replace(
      /([a-zA-Z0-9])[^\s@]*(@[^\s,)]+)/g,
      '$1***$2'
    );
    sheet.appendRow([new Date().toISOString(), action, maskedDetail]);
  } catch (e) {
    console.log('Log error: ' + e.message);
  }
}

// ─── 보안 헬퍼 함수 ──────────────────────────────────────────

/** HTML 태그 제거 (Stored XSS 방지) */
function stripHtml(str) {
  if (!str) return '';
  return String(str).replace(/<[^>]*>/g, '').replace(/[<>]/g, '');
}

/** 이메일 마스킹 (로그용) */
function maskEmail(email) {
  if (!email || email.indexOf('@') === -1) return email || '익명';
  var parts = email.split('@');
  return parts[0].charAt(0) + '***@' + parts[1];
}

/** 제네릭 에러 메시지 (내부 구조 노출 방지) */
function safeErrorMessage(error) {
  // 내부 스택 트레이스, 시트 이름 등을 숨김
  var msg = error.message || '알 수 없는 오류';
  if (msg.indexOf('Spreadsheet') > -1 || msg.indexOf('sheet') > -1) {
    return '서버 내부 오류가 발생했습니다.';
  }
  return msg;
}

// ─── 웹 앱 엔드포인트 (JSONP 지원) ──────────────────────────
function doGet(e) {
  const params = e.parameter || {};
  const action = params.action || '';
  const callback = params.callback || '';

  let result;
  try {
    switch (action) {
      case 'subscribe':
        result = subscribe(params.email, params.name, params.source || 'web');
        break;
      case 'unsubscribe':
        result = unsubscribe(params.email);
        break;
      case 'getArchive':
        result = getArchive();
        break;
      case 'getStats':
        result = getStats(params.password);
        break;
      case 'getSubscribers':
        result = getSubscribers(params.password);
        break;
      case 'deleteSubscriber':
        result = deleteSubscriber(params.password, params.email);
        break;
      case 'sendNewsletter':
        result = sendNewsletter(params.password, params.subject, params.content);
        break;
      case 'testSend':
        result = sendTestNewsletter(params.password, params.subject, params.content, params.email);
        break;
      case 'addSubscriber':
        result = addSubscriberAdmin(params.password, params.email, params.name);
        break;
      case 'changeStatus':
        result = changeSubscriberStatus(params.password, params.email, params.status);
        break;
      case 'saveSchedule':
        result = saveScheduleConfig(params.password, params.days, params.hour);
        break;
      case 'sendPromo':
        result = sendPromoEmails(params.password, params.emails);
        break;
      case 'auth':
        result = authenticate(params.password);
        break;
      case 'submitFeedback':
        result = submitFeedback(params.email, params.rating, params.category, params.message);
        break;
      case 'getFeedback':
        result = getFeedback(params.password);
        break;
      case 'requestPasswordReset':
        result = requestPasswordReset();
        break;
      case 'resetPassword':
        result = resetPassword(params.token, params.newPassword);
        break;
      case 'track':
        // 읽음 추적: 이메일 내 투명 픽셀이 로드될 때 호출
        recordOpen(params.email, params.nlId);
        // 1x1 투명 GIF를 data URI로 포함한 최소 HTML 반환
        return HtmlService.createHtmlOutput(
          '<img src="data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7" width="1" height="1">'
        ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      default:
        result = { success: false, message: '알 수 없는 요청입니다.' };
    }
  } catch (error) {
    result = { success: false, message: safeErrorMessage(error) };
    addLog('ERROR', action + ': ' + error.message);
  }

  // JSONP 콜백이 있으면 JSONP로 응답 (file:// CORS 우회)
  // [보안] callback 이름을 영문/숫자/언더스코어만 허용 (인젝션 방지)
  if (callback) {
    if (!/^[a-zA-Z_$][a-zA-Z0-9_$]*$/.test(callback)) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Invalid callback name' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  // JSON body 지원 — 관리자 페이지(admin.html)가 POST로 JSON을 보냄
  try {
    if (e && e.postData && e.postData.contents) {
      var parsed = JSON.parse(e.postData.contents);
      e.parameter = Object.assign(e.parameter || {}, parsed);
    }
  } catch (err) {
    // JSON 파싱 실패 시 기존 parameter 사용
  }

  // 관리자 페이지에서 사용하는 snake_case action 이름을 camelCase로 매핑
  var a = (e.parameter && e.parameter.action) || '';
  var aliasMap = {
    'send_newsletter': 'sendNewsletter',
    'test_send': 'testSend',
    'delete_subscriber': 'deleteSubscriber',
    'change_status': 'changeStatus',
    'add_subscriber': 'addSubscriber',
    'save_schedule': 'saveSchedule',
    'send_promo': 'sendPromo'
  };
  if (aliasMap[a]) e.parameter.action = aliasMap[a];

  return doGet(e);
}

// ─── 구독 관리 ──────────────────────────────────────────────
function subscribe(email, name, source) {
  if (!email || !validateEmail(email)) {
    return { success: false, message: '올바른 이메일 주소를 입력해주세요.' };
  }

  // [보안] 구독 스팸 방어 — 같은 IP/세션에서 1분 내 중복 요청 차단
  var cache = CacheService.getScriptCache();
  var subKey = 'sub_' + email.toLowerCase().replace(/[^a-z0-9]/g, '_');
  if (cache.get(subKey)) {
    return { success: false, message: '잠시 후 다시 시도해주세요.' };
  }
  cache.put(subKey, '1', 60); // 60초간 중복 방지

  // 이름 살균
  name = stripHtml(name || '').substring(0, 100);

  const sheet = getSheet('subscribers');
  const data = sheet.getDataRange().getValues();

  // 이미 구독 중인지 확인
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === email.toLowerCase()) {
      if (data[i][3] === 'active') {
        return { success: false, message: '이미 구독 중인 이메일입니다.' };
      } else {
        // 재구독
        sheet.getRange(i + 1, 4).setValue('active');
        sheet.getRange(i + 1, 6).setValue('');
        addLog('RESUBSCRIBE', email);
        return { success: true, message: '다시 구독이 시작되었습니다! 환영합니다.' };
      }
    }
  }

  // 신규 구독
  sheet.appendRow([
    email.toLowerCase(),
    name || '',
    new Date().toISOString(),
    'active',
    source,
    ''
  ]);

  addLog('SUBSCRIBE', email + ' (' + source + ')');

  // 환영 이메일 발송
  try {
    sendWelcomeEmail(email, name);
  } catch (e) {
    console.log('Welcome email failed: ' + e.message);
  }

  return { success: true, message: '구독이 완료되었습니다! THE AI INSIGHT 뉴스레터를 통해 최신 AI 트렌드를 만나보세요.' };
}

function unsubscribe(email) {
  if (!email || !validateEmail(email)) {
    return { success: false, message: '올바른 이메일 주소를 입력해주세요.' };
  }

  const sheet = getSheet('subscribers');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === email.toLowerCase()) {
      if (data[i][3] === 'active') {
        sheet.getRange(i + 1, 4).setValue('inactive');
        sheet.getRange(i + 1, 6).setValue(new Date().toISOString());
        addLog('UNSUBSCRIBE', email);
        return { success: true, message: '구독이 취소되었습니다. 다시 돌아오실 때 언제든 환영합니다!' };
      } else {
        return { success: false, message: '이미 구독이 취소된 이메일입니다.' };
      }
    }
  }

  return { success: false, message: '등록되지 않은 이메일 주소입니다.' };
}

// ─── 아카이브 ──────────────────────────────────────────────
function getArchive() {
  const sheet = getSheet('newsletters');
  const data = sheet.getDataRange().getValues();

  const newsletters = [];
  for (let i = 1; i < data.length; i++) {
    newsletters.push({
      id: data[i][0],
      title: data[i][1],
      date: data[i][2],
      summary: data[i][3],
      contentUrl: data[i][4],
      sentCount: data[i][5]
    });
  }

  // 최신순 정렬
  newsletters.sort((a, b) => new Date(b.date) - new Date(a.date));

  return { success: true, data: newsletters };
}

// ─── 통계 ───────────────────────────────────────────────────
function getStats(password) {
  if (!authenticate(password).success) {
    return { success: false, message: '인증에 실패했습니다.' };
  }

  const sheet = getSheet('subscribers');
  const data = sheet.getDataRange().getValues();

  let totalSubscribers = 0;
  let activeSubscribers = 0;
  let inactiveSubscribers = 0;
  const monthlyGrowth = {};
  const sourceCounts = {};

  for (let i = 1; i < data.length; i++) {
    totalSubscribers++;

    if (data[i][3] === 'active') {
      activeSubscribers++;
    } else {
      inactiveSubscribers++;
    }

    // 월별 성장
    const month = data[i][2] ? data[i][2].toString().substring(0, 7) : 'unknown';
    monthlyGrowth[month] = (monthlyGrowth[month] || 0) + 1;

    // 유입 경로
    const source = data[i][4] || 'direct';
    sourceCounts[source] = (sourceCounts[source] || 0) + 1;
  }

  // 뉴스레터 발송 통계
  const nlSheet = getSheet('newsletters');
  const nlData = nlSheet.getDataRange().getValues();
  const totalNewsletters = nlData.length - 1;

  return {
    success: true,
    data: {
      totalSubscribers,
      activeSubscribers,
      inactiveSubscribers,
      totalNewsletters,
      monthlyGrowth,
      sourceCounts,
      internalCount: CONFIG.INTERNAL_EMAILS.length
    }
  };
}

// ─── 구독자 관리 (관리자) ───────────────────────────────────
function getSubscribers(password) {
  if (!authenticate(password).success) {
    return { success: false, message: '인증에 실패했습니다.' };
  }

  const sheet = getSheet('subscribers');
  const data = sheet.getDataRange().getValues();

  const subscribers = [];

  // 1) 내부 구독자 (CONFIG.INTERNAL_EMAILS — AX추진팀 고정 명단) 먼저 표시
  //    sheet에 중복 존재 시 내부 표시로 덮어씀
  const internalSet = {};
  CONFIG.INTERNAL_EMAILS.forEach(function(email) {
    const key = String(email).toLowerCase();
    internalSet[key] = true;
    subscribers.push({
      email: email,
      name: 'AX추진팀',
      subscribedAt: '',
      status: 'active',
      source: 'internal',
      unsubscribedAt: '',
      isInternal: true  // UI에서 삭제 버튼 비활성화 등 구분용
    });
  });

  // 2) 외부 구독자 (subscribers 시트)
  for (let i = 1; i < data.length; i++) {
    const email = data[i][0];
    if (!email) continue;
    // 내부 명단에 이미 있는 이메일은 스킵 (중복 방지)
    if (internalSet[String(email).toLowerCase()]) continue;
    subscribers.push({
      email: email,
      name: data[i][1],
      subscribedAt: data[i][2],
      status: data[i][3],
      source: data[i][4],
      unsubscribedAt: data[i][5],
      isInternal: false
    });
  }

  return { success: true, data: subscribers };
}

function deleteSubscriber(password, email) {
  if (!authenticate(password).success) {
    return { success: false, message: '인증에 실패했습니다.' };
  }

  // 내부 구독자(AX추진팀 고정 명단)는 삭제 불가 — Code.gs CONFIG.INTERNAL_EMAILS 에서 직접 수정해야 함
  const isInternal = CONFIG.INTERNAL_EMAILS.some(function(e) {
    return String(e).toLowerCase() === String(email).toLowerCase();
  });
  if (isInternal) {
    return { success: false, message: '내부 구독자는 삭제할 수 없습니다. Code.gs의 INTERNAL_EMAILS에서 수정하세요.' };
  }

  const sheet = getSheet('subscribers');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === email.toLowerCase()) {
      sheet.deleteRow(i + 1);
      addLog('DELETE', email + ' (by admin)');
      return { success: true, message: email + ' 구독자가 삭제되었습니다.' };
    }
  }

  return { success: false, message: '해당 이메일을 찾을 수 없습니다.' };
}

/**
 * 관리자가 구독자 직접 추가
 */
function addSubscriberAdmin(password, email, name) {
  if (!authenticate(password).success) {
    return { success: false, message: '인증에 실패했습니다.' };
  }
  if (!email || !validateEmail(email)) {
    return { success: false, message: '올바른 이메일을 입력하세요.' };
  }
  var sheet = getSheet('subscribers');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === email.toLowerCase()) {
      return { success: false, message: '이미 존재하는 이메일입니다.' };
    }
  }
  sheet.appendRow([
    email.toLowerCase(),
    stripHtml(name || '').substring(0, 100),
    new Date().toISOString(),
    'active',
    'admin',
    ''
  ]);
  addLog('ADMIN_ADD', email);
  return { success: true, message: email + ' 추가 완료' };
}

/**
 * 구독자 상태 변경 (active / inactive)
 */
function changeSubscriberStatus(password, email, status) {
  if (!authenticate(password).success) {
    return { success: false, message: '인증에 실패했습니다.' };
  }
  var allowed = ['active', 'inactive'];
  if (allowed.indexOf(status) === -1) {
    return { success: false, message: '잘못된 상태값입니다.' };
  }
  var sheet = getSheet('subscribers');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(email).toLowerCase()) {
      sheet.getRange(i + 1, 4).setValue(status);
      if (status === 'inactive') {
        sheet.getRange(i + 1, 6).setValue(new Date().toISOString());
      } else {
        sheet.getRange(i + 1, 6).setValue('');
      }
      addLog('CHANGE_STATUS', email + ' → ' + status);
      return { success: true, message: '상태 변경 완료' };
    }
  }
  return { success: false, message: '해당 이메일을 찾을 수 없습니다.' };
}

/**
 * 발송 주기 설정 저장 (config 시트)
 */
function saveScheduleConfig(password, days, hour) {
  if (!authenticate(password).success) {
    return { success: false, message: '인증에 실패했습니다.' };
  }
  var sheet = getSheet('config');
  var data = sheet.getDataRange().getValues();
  var updates = { 'schedule.days': days || '', 'schedule.hour': String(hour || '7') };
  Object.keys(updates).forEach(function(k) {
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === k) {
        sheet.getRange(i + 1, 2).setValue(updates[k]);
        found = true;
        break;
      }
    }
    if (!found) sheet.appendRow([k, updates[k]]);
  });
  addLog('SAVE_SCHEDULE', 'days=' + days + ' hour=' + hour);
  return { success: true, message: '발송 주기 저장 완료' };
}

/**
 * 테스트 수신 이메일 1명에게 뉴스레터 발송
 */
function sendTestNewsletter(password, subject, htmlContent, email) {
  if (!authenticate(password).success) {
    return { success: false, message: '인증에 실패했습니다.' };
  }
  if (!subject || !email) {
    return { success: false, message: '제목과 테스트 수신 이메일이 필요합니다.' };
  }
  var wrapped = wrapWithKeepAll(htmlContent || ('<p>' + subject + '</p>'));
  GmailApp.sendEmail(email, '[THE AI INSIGHT-테스트] ' + subject, '', {
    htmlBody: wrapped,
    name: 'THE AI INSIGHT - AX추진팀'
  });
  addLog('TEST_SEND', email);
  return { success: true, message: email + ' 테스트 발송 완료' };
}

/**
 * 구독 요청 메일(promo_email) 다중 발송
 * emails: 쉼표로 구분된 이메일 목록
 */
function sendPromoEmails(password, emails) {
  if (!authenticate(password).success) {
    return { success: false, message: '인증에 실패했습니다.' };
  }
  if (!emails) {
    return { success: false, message: '수신 이메일이 없습니다.' };
  }

  var list = String(emails).split(/[,\s]+/).map(function(e) { return e.trim().toLowerCase(); }).filter(function(e) { return e && validateEmail(e); });
  if (!list.length) {
    return { success: false, message: '유효한 이메일이 없습니다.' };
  }

  var webappUrl = ScriptApp.getService().getUrl();
  var subscribeUrl = webappUrl + '?action=subscribe';

  // 구독 요청 메일 HTML — promo_email.html 템플릿과 동일한 디자인을 코드로 구성
  var html = buildPromoEmailHtml(subscribeUrl);
  var subject = '[THE AI INSIGHT] 서연이화 AX추진팀 AI 뉴스레터 구독 안내';

  var sent = 0, failed = [];
  list.forEach(function(to) {
    try {
      GmailApp.sendEmail(to, subject, '', {
        htmlBody: html,
        name: 'THE AI INSIGHT - AX추진팀'
      });
      sent++;
    } catch (e) {
      failed.push(to + ': ' + e.message);
    }
  });

  addLog('SEND_PROMO', sent + '명 발송, ' + failed.length + '건 실패');
  return {
    success: true,
    message: sent + '명에게 구독 요청 메일 발송 완료' + (failed.length ? ' (실패: ' + failed.length + ')' : ''),
    sent: sent,
    failed: failed
  };
}

/**
 * 구독 안내 메일 본문 생성 (promo_email.html 템플릿과 동일 스타일)
 */
function buildPromoEmailHtml(subscribeUrl) {
  var archiveUrl = subscribeUrl.replace('?action=subscribe', '');
  return ''
    + '<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"></head>'
    + '<body style="margin:0;padding:0;background:#f0f4f3;font-family:\'맑은 고딕\',\'Malgun Gothic\',\'Apple SD Gothic Neo\',sans-serif;word-break:keep-all;overflow-wrap:break-word;">'
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f0f4f3;padding:40px 16px;"><tr><td align="center">'
    + '<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:20px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.06);">'
    // 컬러 바
    + '<tr><td style="font-size:0;line-height:0;height:5px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="table-layout:fixed;"><tr>'
    + '<td style="height:5px;background:#f59e0b;"></td><td style="height:5px;background:#dc2626;"></td><td style="height:5px;background:#be185d;"></td>'
    + '<td style="height:5px;background:#059669;"></td><td style="height:5px;background:#0e7490;"></td><td style="height:5px;background:#4f46e5;"></td>'
    + '<td style="height:5px;background:#c2410c;"></td>'
    + '</tr></table></td></tr>'
    // 헤더
    + '<tr><td style="padding:48px 44px 32px;text-align:center;">'
    + '<p style="margin:0 0 12px;font-size:11px;color:#059669;font-weight:700;letter-spacing:0.5px;">서연이화 AX추진팀</p>'
    + '<h1 style="margin:0 0 12px;font-size:34px;font-weight:900;color:#0f172a;letter-spacing:-1.5px;">THE AI INSIGHT</h1>'
    + '<p style="margin:0;font-size:15px;color:#64748b;line-height:1.8;">매일 아침, 최신 AI 뉴스·트렌드와<br>업무 활용 가이드를 정리하여 이메일로 보내드립니다.</p>'
    + '</td></tr>'
    // 본문
    + '<tr><td style="padding:24px 44px 32px;">'
    + '<p style="font-size:15px;color:#374151;line-height:1.8;margin:0 0 20px;">안녕하세요!<br><br>AX추진팀에서 <strong style="color:#059669;">AI 뉴스레터 「THE AI INSIGHT」</strong>를 발행하고 있습니다. 월~금 매일 + 월말/연말 스페셜 에디션으로 AI 트렌드를 정리해 드립니다.</p>'
    + '<div style="text-align:center;margin:28px 0;"><a href="' + subscribeUrl + '" style="display:inline-block;background:linear-gradient(135deg,#059669,#10b981);color:#fff;font-size:16px;font-weight:700;padding:16px 52px;border-radius:14px;text-decoration:none;">지금 구독하기</a></div>'
    + '<p style="text-align:center;margin:0;font-size:13px;color:#94a3b8;"><a href="' + archiveUrl + '" style="color:#94a3b8;text-decoration:none;">지난 뉴스레터 보기 →</a></p>'
    + '</td></tr>'
    // 푸터
    + '<tr><td style="background:#f8fafc;padding:22px 44px;text-align:center;border-top:1px solid #f1f5f9;">'
    + '<p style="margin:0;font-size:12px;color:#94a3b8;">THE AI INSIGHT by AX추진팀 (서연이화)</p>'
    + '</td></tr>'
    + '</table></td></tr></table></body></html>';
}

// ─── 이메일 발송 ────────────────────────────────────────────
function sendNewsletter(password, subject, htmlContent) {
  if (!authenticate(password).success) {
    return { success: false, message: '인증에 실패했습니다.' };
  }

  if (!subject || !htmlContent) {
    return { success: false, message: '제목과 내용을 입력해주세요.' };
  }

  let sentCount = 0;
  const errors = [];
  const sentEmails = []; // 추적용: 발송 성공 이메일 목록

  // 뉴스레터 ID 미리 생성 (추적 픽셀에 사용)
  const nlSheet = getSheet('newsletters');
  const nlId = 'NL-' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-') +
               String(nlSheet.getLastRow()).padStart(3, '0');

  // 한국어 어절 단위 줄바꿈 래퍼 (본문에 keep-all 강제 적용)
  var wrappedContent = wrapWithKeepAll(htmlContent);

  // 1) 내부 구독자 - Gmail 발송 (구독자별 추적 픽셀 삽입)
  CONFIG.INTERNAL_EMAILS.forEach(email => {
    try {
      var trackedHtml = insertTrackingPixel(wrappedContent, email, nlId);
      GmailApp.sendEmail(email, '[THE AI INSIGHT] ' + subject, '', {
        htmlBody: trackedHtml,
        name: 'THE AI INSIGHT - AX추진팀'
      });
      sentCount++;
      sentEmails.push(email);
    } catch (e) {
      errors.push(email + ': ' + e.message);
    }
  });

  // 2) 외부 구독자 발송 (구독자별 추적 픽셀 삽입)
  const sheet = getSheet('subscribers');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === 'active') {
      const email = data[i][0].toString();
      // 내부 이메일은 이미 Gmail로 발송했으므로 스킵
      if (CONFIG.INTERNAL_EMAILS.includes(email)) continue;

      try {
        var trackedHtml = insertTrackingPixel(wrappedContent, email, nlId);
        sendExternalEmail(email, subject, trackedHtml);
        sentCount++;
        sentEmails.push(email);
      } catch (e) {
        errors.push(email + ': ' + e.message);
      }
    }
  }

  // 발송 카운트 업데이트
  if (sentEmails.length > 0) {
    try { incrementSentCount(sentEmails); } catch (e) { /* 카운트 실패해도 발송은 성공 */ }
  }

  // 뉴스레터 기록 저장
  nlSheet.appendRow([nlId, subject, new Date().toISOString(), '', '', sentCount]);

  addLog('SEND_NEWSLETTER', subject + ' (' + sentCount + '명 발송, ' + errors.length + '건 실패)');

  return {
    success: true,
    message: sentCount + '명에게 발송 완료' + (errors.length > 0 ? ' (' + errors.length + '건 실패)' : ''),
    sentCount,
    errors
  };
}

function sendExternalEmail(to, subject, htmlContent) {
  const provider = CONFIG.EXTERNAL_MAILER.provider;

  switch (provider) {
    case 'gmail':
      // Gmail로 외부도 발송 (일 100통 제한 주의)
      GmailApp.sendEmail(to, '[THE AI INSIGHT] ' + subject, '', {
        htmlBody: htmlContent,
        name: 'THE AI INSIGHT'
      });
      break;

    case 'woorimail':
      sendViaWoorimail(to, subject, htmlContent);
      break;

    case 'mailgun':
      sendViaMailgun(to, subject, htmlContent);
      break;

    case 'sendgrid':
      sendViaSendgrid(to, subject, htmlContent);
      break;

    default:
      throw new Error('지원하지 않는 메일 발송 서비스: ' + provider);
  }
}

function sendViaWoorimail(to, subject, htmlContent) {
  const cfg = CONFIG.EXTERNAL_MAILER.woorimail;
  const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMddHHmmss');

  const payload = {
    type: 'api',
    mid: 'auth_woorimail',
    act: 'dispWwapimanagerMailApi',
    authkey: cfg.authkey,
    domain: cfg.domain,
    sender_email: cfg.senderEmail,
    sender_nickname: cfg.senderName,
    receiver_email: to,
    receiver_nickname: to.split('@')[0],
    member_regdate: now,
    title: '[THE AI INSIGHT] ' + subject,
    content: htmlContent
  };

  const response = UrlFetchApp.fetch('https://woorimail.com/index.php', {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());
  if (result.result !== 'OK') {
    throw new Error(result.error_msg || '우리메일 발송 실패');
  }
}

function sendViaMailgun(to, subject, htmlContent) {
  const cfg = CONFIG.EXTERNAL_MAILER.mailgun;
  const response = UrlFetchApp.fetch(
    'https://api.mailgun.net/v3/' + cfg.domain + '/messages',
    {
      method: 'post',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode('api:' + cfg.apiKey)
      },
      payload: {
        from: 'THE AI INSIGHT <' + cfg.senderEmail + '>',
        to: to,
        subject: '[THE AI INSIGHT] ' + subject,
        html: htmlContent
      }
    }
  );
}

function sendViaSendgrid(to, subject, htmlContent) {
  const cfg = CONFIG.EXTERNAL_MAILER.sendgrid;
  UrlFetchApp.fetch('https://api.sendgrid.com/v3/mail/send', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + cfg.apiKey },
    payload: JSON.stringify({
      personalizations: [{ to: [{ email: to }] }],
      from: { email: cfg.senderEmail, name: 'THE AI INSIGHT' },
      subject: '[THE AI INSIGHT] ' + subject,
      content: [{ type: 'text/html', value: htmlContent }]
    })
  });
}

function sendWelcomeEmail(email, name) {
  const greeting = name ? name + '님' : '구독자님';
  const html = `
    <div style="max-width:600px;margin:0 auto;font-family:'Apple SD Gothic Neo',sans-serif;word-break:keep-all;overflow-wrap:break-word;">
      <div style="background:#065f46;padding:24px;text-align:center;">
        <h1 style="color:#fff;margin:0;font-size:22px;">THE AI INSIGHT</h1>
      </div>
      <div style="padding:32px 24px;background:#fff;">
        <h2 style="color:#065f46;">환영합니다, ${greeting}!</h2>
        <p style="color:#374151;line-height:1.8;">
          THE AI INSIGHT 뉴스레터를 구독해주셔서 감사합니다.<br>
          서연이화 AX추진팀이 엄선한 AI 트렌드, 실전 가이드, 업무 자동화 팁을
          정기적으로 보내드리겠습니다.
        </p>
        <p style="color:#6b7280;font-size:14px;margin-top:24px;">
          구독 취소를 원하시면 뉴스레터 하단의 구독취소 링크를 이용해주세요.
        </p>
      </div>
      <div style="background:#f3f4f6;padding:16px;text-align:center;">
        <p style="color:#9ca3af;font-size:12px;margin:0;">
          &copy; 2026 THE AI INSIGHT by AX추진팀 (서연이화)
        </p>
      </div>
    </div>
  `;

  const provider = CONFIG.EXTERNAL_MAILER.provider;
  if (provider === 'gmail' || CONFIG.INTERNAL_EMAILS.includes(email.toLowerCase())) {
    GmailApp.sendEmail(email, '[THE AI INSIGHT] 구독을 환영합니다!', '', {
      htmlBody: html,
      name: 'THE AI INSIGHT'
    });
  } else {
    sendExternalEmail(email, '구독을 환영합니다!', html);
  }
}

// ─── 유틸리티 ──────────────────────────────────────────────

/** 관리자 이메일 (인증 메일 수신 대상) */
var ADMIN_EMAIL = '110316@seoyoneh.com';

/**
 * 인증 함수 — Brute-force 방어 포함
 * 5회 연속 실패 시 계정 잠금 + 관리자에게 비밀번호 재설정 인증 메일 발송
 */
function authenticate(password) {
  var cache = CacheService.getScriptCache();
  var failKey = 'auth_fail_count';
  var lockKey = 'auth_locked';

  // 잠금 확인 — 5회 실패 후에는 인증 메일로만 해제 가능
  if (cache.get(lockKey) === 'true') {
    return { success: false, locked: true, message: '계정이 잠겼습니다. 인증 메일을 확인해주세요.' };
  }

  if (password === CONFIG.ADMIN_PASSWORD) {
    cache.remove(failKey);
    return { success: true };
  }

  // 실패 카운터
  var failCount = parseInt(cache.get(failKey) || '0') + 1;
  cache.put(failKey, String(failCount), 3600); // 1시간 유지

  if (failCount >= 5) {
    cache.put(lockKey, 'true', 3600); // 1시간 잠금 (인증 메일로 즉시 해제 가능)
    cache.remove(failKey);

    // 관리자에게 비밀번호 재설정 인증 메일 발송
    sendPasswordResetEmail();

    addLog('AUTH_LOCKED', '5회 실패 → 계정 잠금 + 인증 메일 발송');
    return { success: false, locked: true, message: '5회 인증 실패. ' + ADMIN_EMAIL + '로 비밀번호 재설정 메일을 발송했습니다.' };
  }

  addLog('AUTH_FAIL', '인증 실패 (' + failCount + '/5)');
  return { success: false, message: '비밀번호가 올바르지 않습니다. (' + failCount + '/5)' };
}

/**
 * 비밀번호 재설정 인증 메일 발송
 * 6자리 인증 토큰을 생성하여 관리자 이메일로 전송
 */
function sendPasswordResetEmail() {
  // 6자리 랜덤 토큰 생성
  var token = '';
  for (var i = 0; i < 6; i++) {
    token += Math.floor(Math.random() * 10);
  }

  // 토큰을 캐시에 저장 (30분 유효)
  var cache = CacheService.getScriptCache();
  cache.put('reset_token', token, 1800);

  var html = '<div style="max-width:480px;margin:0 auto;font-family:\'Apple SD Gothic Neo\',sans-serif;word-break:keep-all;overflow-wrap:break-word;">' +
    '<div style="background:#065f46;padding:20px;text-align:center;border-radius:12px 12px 0 0;">' +
      '<h2 style="color:#fff;margin:0;font-size:18px;">THE AI INSIGHT</h2>' +
      '<p style="color:rgba(255,255,255,0.7);margin:4px 0 0;font-size:12px;">관리자 비밀번호 재설정</p>' +
    '</div>' +
    '<div style="background:#fff;padding:32px;border:1px solid #e5e7eb;border-top:none;">' +
      '<p style="color:#374151;font-size:14px;line-height:1.8;margin:0 0 20px;">관리자 로그인 5회 연속 실패로 계정이 잠겼습니다.<br>아래 인증 코드를 입력하여 새 비밀번호를 설정하세요.</p>' +
      '<div style="background:#f0fdf4;border:2px solid #065f46;border-radius:12px;padding:24px;text-align:center;margin-bottom:20px;">' +
        '<p style="color:#6b7280;font-size:12px;margin:0 0 8px;">인증 코드</p>' +
        '<p style="color:#065f46;font-size:36px;font-weight:900;letter-spacing:8px;margin:0;">' + token + '</p>' +
      '</div>' +
      '<p style="color:#9ca3af;font-size:12px;line-height:1.6;margin:0;">' +
        '⏱ 이 코드는 <strong>30분간</strong> 유효합니다.<br>' +
        '본인이 요청하지 않았다면 이 메일을 무시하세요.' +
      '</p>' +
    '</div>' +
    '<div style="background:#f9fafb;padding:16px;text-align:center;border:1px solid #e5e7eb;border-top:none;border-radius:0 0 12px 12px;">' +
      '<p style="color:#9ca3af;font-size:11px;margin:0;">&copy; 2026 THE AI INSIGHT by AX추진팀</p>' +
    '</div>' +
  '</div>';

  GmailApp.sendEmail(ADMIN_EMAIL, '[THE AI INSIGHT] 관리자 비밀번호 재설정 인증 코드', '', {
    htmlBody: html,
    name: 'THE AI INSIGHT 보안'
  });
}

/**
 * 비밀번호 재설정 요청 (수동 — 관리자가 직접 요청)
 */
function requestPasswordReset() {
  sendPasswordResetEmail();
  addLog('PWD_RESET_REQUEST', '비밀번호 재설정 인증 메일 수동 요청');
  return { success: true, message: ADMIN_EMAIL + '로 인증 코드를 발송했습니다.' };
}

/**
 * 비밀번호 재설정 실행 (인증 코드 + 새 비밀번호)
 * 새 비밀번호: 최소 12자, 영문+숫자+특수문자 포함 필수
 */
function resetPassword(token, newPassword) {
  var cache = CacheService.getScriptCache();
  var savedToken = cache.get('reset_token');

  if (!savedToken) {
    return { success: false, message: '인증 코드가 만료되었습니다. 다시 요청해주세요.' };
  }

  if (token !== savedToken) {
    addLog('PWD_RESET_FAIL', '잘못된 인증 코드 입력');
    return { success: false, message: '인증 코드가 올바르지 않습니다.' };
  }

  // 비밀번호 강도 검증 (최소 12자, 영문+숫자+특수문자)
  if (!newPassword || newPassword.length < 12) {
    return { success: false, message: '비밀번호는 최소 12자 이상이어야 합니다.' };
  }
  if (!/[a-zA-Z]/.test(newPassword)) {
    return { success: false, message: '비밀번호에 영문자를 포함해주세요.' };
  }
  if (!/[0-9]/.test(newPassword)) {
    return { success: false, message: '비밀번호에 숫자를 포함해주세요.' };
  }
  if (!/[!@#$%^&*()_+\-=\[\]{};':"\\|,.<>\/?]/.test(newPassword)) {
    return { success: false, message: '비밀번호에 특수문자를 포함해주세요.' };
  }

  // 비밀번호 변경
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', newPassword);

  // 잠금 해제 + 토큰 삭제
  cache.remove('reset_token');
  cache.remove('auth_locked');
  cache.remove('auth_fail_count');

  addLog('PWD_RESET_SUCCESS', '비밀번호 변경 완료');

  // 확인 메일 발송
  GmailApp.sendEmail(ADMIN_EMAIL, '[THE AI INSIGHT] 관리자 비밀번호가 변경되었습니다', '', {
    htmlBody: '<div style="max-width:480px;margin:0 auto;font-family:sans-serif;padding:24px;word-break:keep-all;overflow-wrap:break-word;">' +
      '<h3 style="color:#065f46;">비밀번호 변경 완료</h3>' +
      '<p style="color:#374151;font-size:14px;line-height:1.8;">' +
        '관리자 비밀번호가 성공적으로 변경되었습니다.<br>' +
        '변경 시각: ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss') + '<br><br>' +
        '본인이 변경하지 않았다면 즉시 Apps Script에서 setupSecureProperties를 실행하여 비밀번호를 재설정하세요.' +
      '</p></div>',
    name: 'THE AI INSIGHT 보안'
  });

  return { success: true, message: '비밀번호가 변경되었습니다. 새 비밀번호로 로그인하세요.' };
}

function validateEmail(email) {
  const re = /^[^\s@]+@[^\s@]+\.[a-zA-Z]{2,}$/;
  return re.test(email);
}

// ─── 초기 설정 (최초 1회 실행) ─────────────────────────────
function setupSheets() {
  getSheet('subscribers');
  getSheet('newsletters');
  getSheet('config');
  getSheet('logs');
  console.log('시트 초기화 완료!');
}

// ─── 정기 발송 트리거 (선택사항) ────────────────────────────
function createWeeklyTrigger() {
  ScriptApp.newTrigger('weeklyDigest')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();
}

function weeklyDigest() {
  // 자동 발송 로직 (필요 시 커스터마이즈)
  addLog('WEEKLY_TRIGGER', '주간 다이제스트 트리거 실행');
}

// ═══════════════════════════════════════════════════════════════
// ─── [월말 스페셜] 발송일 계산 및 가드 로직 ──────────────────
// ═══════════════════════════════════════════════════════════════

/**
 * 휴일 목록 (주말 제외)
 * - 2026: 서연이화 공장 생산예정표 기준 (법정공휴일 + 사내휴일 + 대체조정 반영)
 * - 2027 이후: 한국 법정 공휴일만 기본 적용 (사내 휴일 별도 공지 시 추가)
 */
var KOREAN_PUBLIC_HOLIDAYS = [
  // ═══════════════════════════════════════════════
  // ══ 2026 서연이화 공장 생산예정표 ══
  // ═══════════════════════════════════════════════

  // [신정 휴무] 1/1~1/2
  '2026-01-01',  // 목, 신정
  '2026-01-02',  // 금, 신정 연휴

  // [설날 연휴] 2/16~2/20 (※ 2/20은 5/25 대체공휴일을 앞당겨 조정)
  '2026-02-16',  // 월
  '2026-02-17',  // 화, 설날
  '2026-02-18',  // 수
  '2026-02-19',  // 목
  '2026-02-20',  // 금 (5/25 → 2/20 조정)

  // [삼일절] 3/1은 일요일 → 3/2(월) 대체공휴일
  '2026-03-02',

  // [노동절]
  '2026-05-01',  // 금

  // [어린이날]
  '2026-05-05',  // 화

  // ※ 5/24 석가탄신일(일), 5/25 대체공휴일은 2/20으로 조정되어 정상근무

  // [제8회 전국동시지방선거]
  '2026-06-03',  // 수

  // ※ 6/6 현충일은 토요일 (주말 처리)

  // [하기휴무] 8/3~8/7
  '2026-08-03',
  '2026-08-04',
  '2026-08-05',
  '2026-08-06',
  '2026-08-07',

  // ※ 8/15 광복절은 토요일 → 8/17(월) 대체공휴일
  '2026-08-17',

  // [추석 연휴] 9/24~9/28 (9/26·27은 주말)
  '2026-09-24',  // 목
  '2026-09-25',  // 금, 추석
  '2026-09-28',  // 월, 대체공휴일

  // ※ 10/3 개천절은 토요일 → 10/5(월) 대체공휴일
  '2026-10-05',

  // [한글날]
  '2026-10-09',  // 금

  // [성탄절]
  '2026-12-25',  // 금

  // ※ 12/29 회사 창립일은 정상근무, 대체 휴무일을 12/31로 조정
  '2026-12-31',  // 목

  // ═══════════════════════════════════════════════
  // ══ 2027 한국 법정 공휴일 (사내 휴일 별도 공지 시 추가) ══
  // ═══════════════════════════════════════════════
  '2027-01-01',
  '2027-02-06', '2027-02-07', '2027-02-08', '2027-02-09',  // 설날
  '2027-03-01',
  '2027-05-05',
  '2027-05-13',                      // 석가탄신일
  '2027-06-06', '2027-06-07',        // 현충일 + 대체
  '2027-08-15', '2027-08-16',
  '2027-09-14', '2027-09-15', '2027-09-16',  // 추석
  '2027-10-03', '2027-10-04',
  '2027-10-09', '2027-10-11',
  '2027-12-25'
];

/**
 * 추가 사내 휴일 (정기 휴일표에 없는 임시 휴일 발생 시 추가)
 * 형식: 'YYYY-MM-DD'
 */
var SEOYON_COMPANY_HOLIDAYS = [
  // '2026-MM-DD',
];

/**
 * 주어진 Date가 "쉬는 날"인지 판정
 *   - 토/일요일
 *   - 한국 법정 공휴일
 *   - 서연이화 사내 휴일
 */
function isNonWorkingDay(date) {
  if (!date || !(date instanceof Date)) return false;
  var dow = date.getDay(); // 0=일, 6=토
  if (dow === 0 || dow === 6) return true;

  var y = date.getFullYear();
  var m = String(date.getMonth() + 1).padStart(2, '0');
  var d = String(date.getDate()).padStart(2, '0');
  var key = y + '-' + m + '-' + d;

  if (KOREAN_PUBLIC_HOLIDAYS.indexOf(key) > -1) return true;
  if (SEOYON_COMPANY_HOLIDAYS.indexOf(key) > -1) return true;
  return false;
}

/**
 * 월말 스페셜 발송일 계산
 *   - 해당 월의 마지막 날부터 거꾸로 내려오면서
 *   - 공휴일·회사휴일·주말이 아닌 첫 평일을 반환
 *
 * @param {number} year  예: 2026
 * @param {number} month 1-12
 * @return {Date} 발송 예정일
 */
function getMonthlySpecialSendDate(year, month) {
  // new Date(year, month, 0) → 해당 월의 마지막 날
  var d = new Date(year, month, 0);
  // 안전장치: 최대 15일 거슬러 올라감 (무한루프 방지)
  for (var i = 0; i < 15; i++) {
    if (!isNonWorkingDay(d)) return d;
    d.setDate(d.getDate() - 1);
  }
  return d;
}

/**
 * 오늘이 이번 달의 월말 스페셜 발송일인지 확인
 * @param {Date} [date] 테스트용 날짜 (기본값: 오늘)
 * @return {boolean}
 */
function isMonthlySpecialDay(date) {
  date = date || new Date();
  var sendDate = getMonthlySpecialSendDate(date.getFullYear(), date.getMonth() + 1);
  return date.getFullYear() === sendDate.getFullYear()
      && date.getMonth() === sendDate.getMonth()
      && date.getDate() === sendDate.getDate();
}

/**
 * 오늘 월~금 정기 뉴스레터를 발송해야 하는지 여부
 *   - 쉬는 날이면 false
 *   - 월말 스페셜 발송일이면 false (월말 스페셜만 발송)
 *   - 그 외 평일이면 true
 *
 * @param {Date} [date] 테스트용 날짜 (기본값: 오늘)
 * @return {{ send: boolean, reason: string }}
 */
function shouldSendDailyNewsletterToday(date) {
  date = date || new Date();
  if (isNonWorkingDay(date)) {
    return { send: false, reason: 'HOLIDAY' };
  }
  if (isMonthlySpecialDay(date)) {
    return { send: false, reason: 'MONTHLY_SPECIAL_DAY' };
  }
  return { send: true, reason: 'WEEKDAY' };
}

/**
 * [자동 실행용] 매일 아침 호출 — 오늘 발송할 뉴스레터 종류 결정
 *   - 월말 스페셜 발송일 → 월말 스페셜 발송
 *   - 평일 → 월~금 뉴스레터 발송
 *   - 휴일 → 발송하지 않음
 *
 * Apps Script 시간 트리거에 이 함수를 등록 (매일 오전 8시 추천)
 */
function dailyNewsletterDispatcher() {
  var today = new Date();

  if (isMonthlySpecialDay(today)) {
    addLog('MONTHLY_SPECIAL', '월말 스페셜 발송일 감지 → 월말 스페셜만 발송 (월~금 스킵)');
    // TODO: sendMonthlySpecial() 함수 호출 (별도 구현)
    // sendMonthlySpecial();
    return { sent: 'monthly_special' };
  }

  var decision = shouldSendDailyNewsletterToday(today);
  if (!decision.send) {
    addLog('SKIP_DAILY', '오늘 발송 생략 — 사유: ' + decision.reason);
    return { sent: 'none', reason: decision.reason };
  }

  // 평일 → 기존 월~금 뉴스레터 발송 로직 호출
  // TODO: sendDailyNewsletter() 함수 호출 (별도 구현/연결)
  addLog('DAILY_DISPATCH', '평일 뉴스레터 발송 진행');
  return { sent: 'daily' };
}

/**
 * 월말 스페셜 트리거 설정 (1회 실행)
 * — 매일 오전 8시에 dailyNewsletterDispatcher 실행
 * — dispatcher 내부에서 "오늘 뭘 보낼지" 판단
 */
function createDailyDispatcherTrigger() {
  // 기존 트리거 제거
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'dailyNewsletterDispatcher') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('dailyNewsletterDispatcher')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  addLog('TRIGGER_CREATED', '일간 뉴스레터 디스패처 트리거 등록 (매일 08시)');
}

/**
 * [테스트] 올해 각 월의 월말 스페셜 발송 예정일을 로그로 출력
 * Apps Script 편집기에서 직접 실행하여 확인
 */
function testMonthlySpecialSchedule() {
  var y = new Date().getFullYear();
  var days = ['일', '월', '화', '수', '목', '금', '토'];
  for (var m = 1; m <= 12; m++) {
    var d = getMonthlySpecialSendDate(y, m);
    var label = y + '-' + String(m).padStart(2, '0') + ' → '
              + d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0')
              + ' (' + days[d.getDay()] + ')';
    console.log(label);
  }
}

// ─── Google Form 자동 생성 (1회 실행) ──────────────────────
function createSubscribeForm() {
  const form = FormApp.create('THE AI INSIGHT 뉴스레터 구독');
  form.setDescription('서연이화 AX추진팀이 전하는 AI 트렌드 뉴스레터');
  form.setConfirmationMessage('구독이 완료되었습니다! THE AI INSIGHT 뉴스레터를 통해 최신 AI 트렌드를 만나보세요.');

  // 이메일 필드 (필수)
  form.addTextItem()
    .setTitle('이메일')
    .setHelpText('뉴스레터를 받을 이메일 주소')
    .setRequired(true);

  // 이름 필드 (선택)
  form.addTextItem()
    .setTitle('이름')
    .setHelpText('선택사항')
    .setRequired(false);

  // 유입 경로 (숨김용 - 기본값 web)
  form.addTextItem()
    .setTitle('유입경로')
    .setRequired(false);

  // 응답을 Google Sheets에 연결
  form.setDestination(FormApp.DestinationType.SPREADSHEET, CONFIG.SPREADSHEET_ID);

  // 폼 제출 트리거 생성
  ScriptApp.newTrigger('onFormSubmit')
    .forForm(form)
    .onFormSubmit()
    .create();

  const formUrl = form.getPublishedUrl();
  const formId = form.getId();
  const editUrl = form.getEditUrl();

  // config 시트에 저장
  const configSheet = getSheet('config');
  configSheet.appendRow(['form_url', formUrl]);
  configSheet.appendRow(['form_id', formId]);
  configSheet.appendRow(['form_edit_url', editUrl]);

  console.log('=== 폼 생성 완료 ===');
  console.log('폼 URL: ' + formUrl);
  console.log('폼 ID: ' + formId);
  console.log('수정 URL: ' + editUrl);

  return { formUrl, formId };
}

// ─── 폼 제출 시 자동 처리 ──────────────────────────────────
function onFormSubmit(e) {
  try {
    const responses = e.response.getItemResponses();
    const email = responses[0].getResponse().trim().toLowerCase();
    const name = responses.length > 1 ? responses[1].getResponse().trim() : '';
    const source = responses.length > 2 ? responses[2].getResponse().trim() || 'web' : 'web';

    // subscribers 시트에 직접 추가
    const result = subscribe(email, name, source);
    addLog('FORM_SUBMIT', email + ' - ' + result.message);
  } catch (err) {
    addLog('FORM_ERROR', err.message);
  }
}

// ─── Google Sheets 웹 발행 설정 ────────────────────────────
function publishSheet() {
  console.log('Google Sheets 웹 발행 안내:');
  console.log('1. Google Sheets 열기: https://docs.google.com/spreadsheets/d/' + CONFIG.SPREADSHEET_ID);
  console.log('2. 파일 → 공유 → 웹에 게시');
  console.log('3. "전체 문서" 선택 → "게시" 클릭');
}

// ─── 테스트 뉴스레터 발송 (1명에게만) ─────────────────────
function testSendNewsletter() {
  var subject = '이번 주 AI 하이라이트 - 테스트 발송';
  var html = ''
    + '<div style="max-width:640px;margin:0 auto;font-family:Apple SD Gothic Neo,Noto Sans KR,sans-serif;word-break:keep-all;overflow-wrap:break-word;">'
    + '  <div style="background:#065f46;padding:28px 24px;text-align:center;">'
    + '    <h1 style="color:#fff;margin:0;font-size:22px;letter-spacing:-0.5px;">THE AI INSIGHT</h1>'
    + '    <p style="color:rgba(255,255,255,0.8);margin:8px 0 0;font-size:13px;">by 서연이화 AX추진팀</p>'
    + '  </div>'
    + '  <div style="padding:32px 24px;background:#fff;">'
    + '    <p style="color:#6b7280;font-size:13px;margin:0 0 4px;">2026년 4월 9일 · 제6호</p>'
    + '    <h2 style="color:#065f46;font-size:20px;margin:0 0 20px;line-height:1.4;">AI Agent 시대, 업무 자동화의 새로운 패러다임</h2>'
    + '    <p style="color:#374151;font-size:15px;line-height:1.8;margin:0 0 24px;">'
    + '      이번 주 가장 주목할 AI 뉴스는 Anthropic의 Claude Code 업데이트입니다. '
    + '      MCP(Model Context Protocol)를 통해 사내 시스템과 AI를 직접 연결할 수 있게 되면서, '
    + '      개발자뿐만 아니라 비개발 직군도 AI 자동화를 구축할 수 있는 시대가 열리고 있습니다.'
    + '    </p>'
    + '    <div style="background:#ecfdf5;border-left:4px solid #065f46;padding:16px 20px;border-radius:0 8px 8px 0;margin:0 0 24px;">'
    + '      <h3 style="color:#065f46;font-size:15px;margin:0 0 8px;">핵심 인사이트</h3>'
    + '      <ul style="color:#374151;font-size:14px;line-height:1.8;margin:0;padding-left:20px;">'
    + '        <li>Claude Code Cowork 모드로 비개발자도 자동화 구축 가능</li>'
    + '        <li>Google Sheets + Apps Script 조합이 사내 경량 시스템의 표준으로 부상</li>'
    + '        <li>RAG 도입 시 초기 비용 대비 ROI가 6개월 내 200% 이상 달성 가능</li>'
    + '      </ul>'
    + '    </div>'
    + '    <h3 style="color:#1f2937;font-size:16px;margin:0 0 12px;">이번 주 추천 도구</h3>'
    + '    <p style="color:#374151;font-size:14px;line-height:1.8;margin:0 0 24px;">'
    + '      <strong>Claude Code</strong> - 터미널 기반 AI 코딩 도구. MCP 서버 연동으로 Slack, Notion, JIRA 등과 통합 가능.<br>'
    + '      <strong>NotebookLM</strong> - Google의 AI 문서 분석 도구. 사내 매뉴얼 검색에 유용.'
    + '    </p>'
    + '  </div>'
    + '  <div style="background:#f3f4f6;padding:20px 24px;text-align:center;">'
    + '    <p style="color:#9ca3af;font-size:12px;margin:0;">'
    + '      &copy; 2026 THE AI INSIGHT by AX추진팀 (서연이화)<br>'
    + '      <a href="https://aa200526seoyoneh-star.github.io/ai-insight-newsletter/?action=unsubscribe" style="color:#065f46;">구독 취소</a>'
    + '    </p>'
    + '  </div>'
    + '</div>';

  // 내부 이메일(테스트: 110316@seoyoneh.com)에게만 발송
  CONFIG.INTERNAL_EMAILS.forEach(function(email) {
    GmailApp.sendEmail(email, '[THE AI INSIGHT] ' + subject, '', {
      htmlBody: html,
      name: 'THE AI INSIGHT - AX추진팀'
    });
  });

  // 발송 기록 저장
  var nlSheet = getSheet('newsletters');
  nlSheet.appendRow(['NL-2026-TEST', subject, new Date().toISOString(), '테스트 발송', '', CONFIG.INTERNAL_EMAILS.length]);

  Logger.log(CONFIG.INTERNAL_EMAILS.length + '명에게 발송 완료!');
}

// ─── 관리자 액션 폼 생성 (1회 실행) ────────────────────────
function createAdminForm() {
  var form = FormApp.create('THE AI INSIGHT 관리자');
  form.setDescription('관리자 전용 - 뉴스레터 발송 및 구독자 관리');

  // 비밀번호
  form.addTextItem()
    .setTitle('비밀번호')
    .setRequired(true);

  // 액션 선택
  form.addMultipleChoiceItem()
    .setTitle('액션')
    .setChoiceValues(['뉴스레터 발송', '구독자 상태 변경', '구독자 삭제'])
    .setRequired(true);

  // 뉴스레터 제목
  form.addTextItem()
    .setTitle('뉴스레터 제목')
    .setRequired(false);

  // 뉴스레터 HTML 내용
  form.addParagraphTextItem()
    .setTitle('뉴스레터 내용(HTML)')
    .setRequired(false);

  // 대상 이메일 (구독자 관리용)
  form.addTextItem()
    .setTitle('대상 이메일')
    .setRequired(false);

  // 상태 변경값
  form.addMultipleChoiceItem()
    .setTitle('변경할 상태')
    .setChoiceValues(['active', 'inactive'])
    .setRequired(false);

  // 트리거 연결
  ScriptApp.newTrigger('onAdminFormSubmit')
    .forForm(form)
    .onFormSubmit()
    .create();

  var configSheet = getSheet('config');
  configSheet.appendRow(['admin_form_id', form.getId()]);
  configSheet.appendRow(['admin_form_url', form.getPublishedUrl()]);

  Logger.log('=== 관리자 폼 생성 완료 ===');
  Logger.log('폼 ID: ' + form.getId());
  Logger.log('Published URL: ' + form.getPublishedUrl());

  // entry ID 출력
  var items = form.getItems();
  for (var i = 0; i < items.length; i++) {
    Logger.log(items[i].getTitle() + ' → entry.' + items[i].getId());
  }
}

// ─── 관리자 폼 제출 처리 ───────────────────────────────────
function onAdminFormSubmit(e) {
  try {
    var responses = e.response.getItemResponses();
    var password = responses[0].getResponse().trim();
    var action = responses[1].getResponse().trim();
    var nlSubject = responses.length > 2 ? responses[2].getResponse().trim() : '';
    var nlContent = responses.length > 3 ? responses[3].getResponse().trim() : '';
    var targetEmail = responses.length > 4 ? responses[4].getResponse().trim() : '';
    var newStatus = responses.length > 5 ? responses[5].getResponse().trim() : '';

    // 비밀번호 확인
    if (password !== CONFIG.ADMIN_PASSWORD) {
      addLog('ADMIN_DENIED', '잘못된 비밀번호로 관리자 액션 시도');
      return;
    }

    switch (action) {
      case '뉴스레터 발송':
        if (nlSubject) {
          var result = sendNewsletter(password, nlSubject, nlContent || '<p>' + nlSubject + '</p>');
          addLog('ADMIN_SEND', result.message);
        }
        break;

      case '구독자 상태 변경':
        if (targetEmail && newStatus) {
          var sheet = getSheet('subscribers');
          var data = sheet.getDataRange().getValues();
          for (var i = 1; i < data.length; i++) {
            if (data[i][0].toString().toLowerCase() === targetEmail.toLowerCase()) {
              sheet.getRange(i + 1, 4).setValue(newStatus);
              if (newStatus === 'inactive') {
                sheet.getRange(i + 1, 6).setValue(new Date().toISOString());
              }
              addLog('ADMIN_STATUS', targetEmail + ' → ' + newStatus);
              break;
            }
          }
        }
        break;

      case '구독자 삭제':
        if (targetEmail) {
          deleteSubscriber(password, targetEmail);
        }
        break;

      case '테스트 발송':
        if (nlSubject && targetEmail) {
          try {
            GmailApp.sendEmail(targetEmail, '[THE AI INSIGHT-테스트] ' + nlSubject, '', {
              htmlBody: nlContent || '<p>' + nlSubject + '</p>',
              name: 'THE AI INSIGHT - AX추진팀'
            });
            addLog('ADMIN_TEST_SEND', '테스트 발송 → ' + targetEmail + ' / ' + nlSubject);
          } catch (err2) {
            addLog('ADMIN_TEST_SEND_FAIL', targetEmail + ': ' + err2.message);
          }
        }
        break;

      case '구독자 추가':
        if (targetEmail) {
          var addResult = subscribe(targetEmail, newStatus || '', 'admin');
          addLog('ADMIN_ADD_SUBSCRIBER', targetEmail + ' → ' + addResult.message);
        }
        break;

      case '발송 주기 변경':
        if (nlSubject) {
          // nlSubject 필드에 JSON 형태로 주기 정보 전달: {"days":"mon,wed,fri","hour":"7"}
          try {
            var scheduleData = JSON.parse(nlSubject);
            var configSheet = getSheet('config');
            var configData = configSheet.getDataRange().getValues();
            var updated = {};

            // config 시트에 SEND_DAYS, SEND_HOUR 설정
            ['SEND_DAYS', 'SEND_HOUR'].forEach(function(key) {
              var val = key === 'SEND_DAYS' ? (scheduleData.days || 'mon,tue,wed,thu,fri') : (scheduleData.hour || '7');
              var found = false;
              for (var ci = 1; ci < configData.length; ci++) {
                if (configData[ci][0] === key) {
                  configSheet.getRange(ci + 1, 2).setValue(val);
                  found = true;
                  break;
                }
              }
              if (!found) configSheet.appendRow([key, val]);
              updated[key] = val;
            });
            addLog('ADMIN_SCHEDULE', 'SEND_DAYS=' + updated.SEND_DAYS + ', SEND_HOUR=' + updated.SEND_HOUR);
          } catch (parseErr) {
            addLog('ADMIN_SCHEDULE_ERROR', parseErr.message);
          }
        }
        break;
    }
  } catch (err) {
    addLog('ADMIN_ERROR', err.message);
  }
}

// ─── Gmail 보낸편지함에서 뉴스레터 아카이브 자동 수집 ──────────
/**
 * Gmail 보낸편지함에서 [THE AI INSIGHT] 제목의 정식 발송 메일을 찾아
 * newsletters 시트에 자동 기록합니다.
 * - 수신자가 2명 이상인 메일만 정식 발송으로 판단
 * - 이미 기록된 메일은 중복 추가하지 않음
 * - 매일 트리거로 자동 실행 권장
 */
function syncNewslettersFromGmail() {
  var nlSheet = getSheet('newsletters');
  var existingData = nlSheet.getDataRange().getValues();

  // 기존에 기록된 제목+날짜 조합을 Set으로 저장 (중복 방지)
  var existingKeys = {};
  for (var i = 1; i < existingData.length; i++) {
    var key = (existingData[i][1] || '') + '|' + (existingData[i][2] || '').toString().substring(0, 10);
    existingKeys[key] = true;
  }

  // Gmail에서 THE AI INSIGHT 발송 메일 검색 (최근 7일)
  var query = 'from:me subject:"[THE AI INSIGHT]" -subject:테스트 -subject:TEST newer_than:7d in:sent';
  var threads = GmailApp.search(query, 0, 50);
  var newCount = 0;

  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();
    var msg = messages[0];

    // 수신자 수 확인 (2명 이상이면 정식 발송)
    var toField = msg.getTo() || '';
    var recipients = toField.split(',');
    if (recipients.length < 2) continue;

    var subject = msg.getSubject().replace(/^\[THE AI INSIGHT\]\s*/, '');
    var sentDate = Utilities.formatDate(msg.getDate(), 'Asia/Seoul', 'yyyy-MM-dd');
    var snippet = msg.getPlainBody().substring(0, 200).replace(/\n/g, ' ').trim();

    // 중복 확인
    var checkKey = subject + '|' + sentDate;
    if (existingKeys[checkKey]) continue;

    // 뉴스레터 ID 생성
    var nlId = 'NL-' + Utilities.formatDate(msg.getDate(), 'Asia/Seoul', 'yyyy-') +
               String(nlSheet.getLastRow() + 1).padStart(3, '0');

    // newsletters 시트에 기록
    nlSheet.appendRow([nlId, subject, sentDate, snippet, '', recipients.length]);
    existingKeys[checkKey] = true;
    newCount++;
  }

  if (newCount > 0) {
    addLog('SYNC_ARCHIVE', newCount + '개 뉴스레터 아카이브에 추가');
  }

  // 공개 시트에도 newsletters 데이터 동기화
  syncNewslettersToPublicSheet();

  return { success: true, added: newCount };
}

/**
 * newsletters 시트 데이터를 공개용 시트에 동기화
 * 아카이브 페이지가 PUBLIC_SHEET_ID를 읽으므로 이쪽에도 데이터가 있어야 함
 */
function syncNewslettersToPublicSheet() {
  try {
    var source = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName('newsletters');
    if (!source) return;

    var publicSS = SpreadsheetApp.openById(CONFIG.PUBLIC_SHEET_ID);
    var target = publicSS.getSheetByName('newsletters');
    if (!target) {
      target = publicSS.insertSheet('newsletters');
    }

    // 전체 데이터 복사 (헤더 포함)
    var data = source.getDataRange().getValues();
    target.clearContents();
    if (data.length > 0) {
      target.getRange(1, 1, data.length, data[0].length).setValues(data);
    }

    addLog('SYNC_PUBLIC', 'newsletters → 공개 시트 동기화 완료 (' + (data.length - 1) + '건)');
  } catch (e) {
    addLog('SYNC_ERROR', '공개 시트 동기화 실패: ' + e.message);
  }
}

/**
 * 아카이브 수집 트리거 생성 (1회 실행)
 * 매일 오전 8시에 syncNewslettersFromGmail 실행
 */
function createArchiveSyncTrigger() {
  // 기존 트리거 중복 방지
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncNewslettersFromGmail') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('syncNewslettersFromGmail')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  addLog('TRIGGER_CREATED', '아카이브 자동 수집 트리거 (매일 08시)');
}

// ═══════════════════════════════════════════════════════════════
// 읽음 추적 (Open Tracking) 시스템
// ═══════════════════════════════════════════════════════════════

/**
 * 시트 구조: [tracking] email | nlId | opened_at
 * 시트 구조: [subscribers]에 컬럼 추가 → sent_count(7번째), open_count(8번째)
 */

/**
 * 이메일 열람 기록
 * 뉴스레터에 삽입된 추적 픽셀이 로드될 때 호출됨
 */
function recordOpen(email, nlId) {
  if (!email) return;
  email = email.toLowerCase().trim();

  try {
    var sheet = getSheet('tracking');
    // 중복 방지: 같은 이메일 + 같은 뉴스레터는 1회만 기록
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === email && data[i][1] === nlId) return;
    }
    sheet.appendRow([email, nlId || '', new Date().toISOString()]);

    // subscribers 시트의 open_count 증가
    updateOpenCount(email);
  } catch (e) {
    addLog('TRACK_ERROR', email + ': ' + e.message);
  }
}

/**
 * 구독자의 open_count를 +1 업데이트
 */
function updateOpenCount(email) {
  var sheet = getSheet('subscribers');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  // open_count, sent_count 컬럼 위치 찾기 (없으면 추가)
  var sentIdx = headers.indexOf('sent_count');
  var openIdx = headers.indexOf('open_count');

  if (sentIdx === -1) {
    sentIdx = headers.length;
    sheet.getRange(1, sentIdx + 1).setValue('sent_count');
  }
  if (openIdx === -1) {
    openIdx = headers.length + (sentIdx === headers.length ? 1 : 0);
    sheet.getRange(1, openIdx + 1).setValue('open_count');
  }

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === email) {
      var currentOpen = parseInt(data[i][openIdx]) || 0;
      sheet.getRange(i + 1, openIdx + 1).setValue(currentOpen + 1);
      break;
    }
  }
}

/**
 * 뉴스레터 발송 시 구독자의 sent_count를 +1 업데이트
 * 발송 함수에서 호출
 */
function incrementSentCount(emailList) {
  var sheet = getSheet('subscribers');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var sentIdx = headers.indexOf('sent_count');
  if (sentIdx === -1) {
    sentIdx = headers.length;
    sheet.getRange(1, sentIdx + 1).setValue('sent_count');
    data = sheet.getDataRange().getValues();
  }

  var emailSet = {};
  emailList.forEach(function(e) { emailSet[e.toLowerCase().trim()] = true; });

  for (var i = 1; i < data.length; i++) {
    if (emailSet[data[i][0].toString().toLowerCase()]) {
      var current = parseInt(data[i][sentIdx]) || 0;
      sheet.getRange(i + 1, sentIdx + 1).setValue(current + 1);
    }
  }
}

/**
 * 뉴스레터 HTML 본문에 한국어 어절 단위 줄바꿈(word-break:keep-all) 스타일을 강제 적용
 * - <body>가 있으면 style 속성에 주입
 * - <body>가 없으면 전체를 <div>로 감싸서 keep-all 적용
 */
function wrapWithKeepAll(htmlContent) {
  if (!htmlContent) return htmlContent;
  var KEEP_ALL = 'word-break:keep-all;overflow-wrap:break-word;';

  // 1) <body ...> 태그가 있는 경우 style 속성에 주입
  var bodyMatch = htmlContent.match(/<body\b[^>]*>/i);
  if (bodyMatch) {
    var openTag = bodyMatch[0];
    var newTag;
    if (/style\s*=\s*"/i.test(openTag)) {
      newTag = openTag.replace(/style\s*=\s*"([^"]*)"/i, function(_, s) {
        return 'style="' + s.replace(/;?\s*$/, ';') + KEEP_ALL + '"';
      });
    } else {
      newTag = openTag.replace(/<body/i, '<body style="' + KEEP_ALL + '"');
    }
    return htmlContent.replace(openTag, newTag);
  }

  // 2) <body>가 없으면 외곽 div로 감싸기
  return '<div style="' + KEEP_ALL + '">' + htmlContent + '</div>';
}

/**
 * 뉴스레터 HTML에 추적 픽셀을 삽입하는 헬퍼
 * 각 구독자별 고유 URL 생성
 */
function insertTrackingPixel(htmlContent, email, nlId) {
  var webappUrl = ScriptApp.getService().getUrl();
  var trackUrl = webappUrl + '?action=track&email=' + encodeURIComponent(email) + '&nlId=' + encodeURIComponent(nlId || '');
  var pixel = '<img src="' + trackUrl + '" width="1" height="1" style="display:none;" alt="">';
  // </body> 태그 앞에 삽입
  if (htmlContent.indexOf('</body>') > -1) {
    return htmlContent.replace('</body>', pixel + '</body>');
  }
  return htmlContent + pixel;
}

// ═══════════════════════════════════════════════════════════════
// 자동 구독취소 & 독려 메일 시스템
// ═══════════════════════════════════════════════════════════════

/**
 * 20회 발송 후 1회도 읽지 않은 구독자를 자동 구독취소하고 독려 메일 발송
 * 매일 트리거로 실행 (checkInactiveAndNotify)
 */
function checkInactiveAndNotify() {
  // [보안] Race condition 방어 — 동시 실행 방지
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    addLog('INACTIVE_CHECK', '다른 실행이 진행 중 — 건너뜀');
    return;
  }

  try {
    _doInactiveCheck();
  } finally {
    lock.releaseLock();
  }
}

function _doInactiveCheck() {
  var sheet = getSheet('subscribers');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var sentIdx = headers.indexOf('sent_count');
  var openIdx = headers.indexOf('open_count');
  var statusIdx = headers.indexOf('status');
  var emailIdx = 0;
  var nameIdx = 1;

  if (sentIdx === -1 || openIdx === -1) {
    addLog('INACTIVE_CHECK', 'sent_count 또는 open_count 컬럼이 없습니다. 추적이 아직 시작되지 않았습니다.');
    return;
  }

  var THRESHOLD = 20; // 연속 미독 기준
  var inactiveCount = 0;

  for (var i = 1; i < data.length; i++) {
    if (data[i][statusIdx] !== 'active') continue;

    var sent = parseInt(data[i][sentIdx]) || 0;
    var opened = parseInt(data[i][openIdx]) || 0;
    var email = data[i][emailIdx].toString().toLowerCase();
    var name = data[i][nameIdx] || '';

    // 발송 20회 이상인데 열람 0회인 구독자
    if (sent >= THRESHOLD && opened === 0) {
      // 1. 구독 상태를 inactive로 변경
      sheet.getRange(i + 1, statusIdx + 1).setValue('inactive');
      sheet.getRange(i + 1, headers.indexOf('unsubscribed_at') + 1).setValue(
        Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd')
      );

      // 2. 재구독 독려 메일 발송
      sendReEngagementEmail(email, name);

      inactiveCount++;
      addLog('AUTO_UNSUB', email + ' (발송 ' + sent + '회, 열람 ' + opened + '회)');
    }
  }

  if (inactiveCount > 0) {
    addLog('INACTIVE_CHECK', inactiveCount + '명 자동 구독취소 및 독려 메일 발송');
  }
}

/**
 * 재구독 독려 메일 발송
 */
function sendReEngagementEmail(email, name) {
  var displayName = name || '구독자';
  // [보안] 로컬 경로 대신 구독 페이지 상대 경로 사용
  var subscribeUrl = ScriptApp.getService().getUrl() + '?action=subscribe';

  var subject = '[THE AI INSIGHT] ' + displayName + '님, 다시 만나고 싶어요!';

  var htmlBody = '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f4f5f7;font-family:sans-serif;word-break:keep-all;overflow-wrap:break-word;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f5f7;padding:32px 16px;"><tr><td align="center">'
    + '<table width="560" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:16px;overflow:hidden;">'

    // Header
    + '<tr><td style="background:#065f46;padding:36px 32px;text-align:center;">'
    + '<h1 style="margin:0;font-size:24px;color:#fff;font-weight:900;">THE AI INSIGHT</h1>'
    + '</td></tr>'

    // Body
    + '<tr><td style="padding:36px 32px;">'
    + '<p style="font-size:15px;color:#374151;line-height:1.8;margin:0 0 20px;">'
    + displayName + '님, 안녕하세요.<br><br>'
    + '한동안 뉴스레터를 읽어주시지 않으셔서, 구독이 자동으로 해제되었습니다.<br>'
    + '혹시 바쁘셨거나 메일이 스팸함에 들어갔던 건 아닌지 걱정됩니다.</p>'

    + '<p style="font-size:15px;color:#374151;line-height:1.8;margin:0 0 24px;">'
    + '최근 AI 업계에는 흥미로운 변화들이 많았습니다.<br>'
    + '<strong style="color:#065f46;">다시 구독하시면</strong> 매일 아침 핵심 AI 인사이트를 받아보실 수 있어요!</p>'

    // CTA
    + '<table width="100%" cellpadding="0" cellspacing="0"><tr><td align="center">'
    + '<a href="' + subscribeUrl + '" style="display:inline-block;background:#065f46;color:#fff;font-size:16px;font-weight:700;'
    + 'padding:14px 40px;border-radius:10px;text-decoration:none;">다시 구독하기 →</a>'
    + '</td></tr></table>'

    + '<p style="text-align:center;margin:20px 0 0;font-size:13px;color:#9ca3af;">'
    + '이 메일은 자동 발송된 안내 메일입니다.</p>'
    + '</td></tr>'

    // Footer
    + '<tr><td style="background:#111827;padding:20px 32px;text-align:center;">'
    + '<p style="margin:0;font-size:12px;color:rgba(255,255,255,0.4);">THE AI INSIGHT by AX추진팀 (서연이화)</p>'
    + '</td></tr>'

    + '</table></td></tr></table></body></html>';

  try {
    GmailApp.sendEmail(email, subject, '뉴스레터 구독이 해제되었습니다. 다시 구독하려면 사이트를 방문해주세요.', {
      htmlBody: htmlBody,
      name: 'THE AI INSIGHT'
    });
    addLog('RE_ENGAGE_SENT', email);
  } catch (e) {
    addLog('RE_ENGAGE_ERROR', email + ': ' + e.message);
  }
}

/**
 * 비활성 구독자 체크 + 독려 메일 트리거 생성 (1회 실행)
 * 매일 오전 9시에 checkInactiveAndNotify 실행
 */
function createInactiveCheckTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkInactiveAndNotify') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('checkInactiveAndNotify')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  addLog('TRIGGER_CREATED', '비활성 구독자 체크 트리거 (매일 09시)');
}

/**
 * tracking 시트 초기화 (헤더 자동 생성)
 */
function initTrackingSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('tracking');
  if (!sheet) {
    sheet = ss.insertSheet('tracking');
    sheet.getRange(1, 1, 1, 3).setValues([['email', 'nlId', 'opened_at']]);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  // subscribers 시트에 sent_count, open_count 컬럼 추가
  var subSheet = getSheet('subscribers');
  var headers = subSheet.getRange(1, 1, 1, subSheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('sent_count') === -1) {
    subSheet.getRange(1, headers.length + 1).setValue('sent_count');
  }
  if (headers.indexOf('open_count') === -1) {
    var newHeaders = subSheet.getRange(1, 1, 1, subSheet.getLastColumn()).getValues()[0];
    subSheet.getRange(1, newHeaders.length + 1).setValue('open_count');
  }
  addLog('INIT', 'tracking 시트 및 subscribers 컬럼 초기화 완료');
}

// ═══════════════════════════════════════════════════════════════
// 피드백 시스템
// ═══════════════════════════════════════════════════════════════

/**
 * 시트 구조: [feedback] timestamp | email | rating | category | message | status
 */

/**
 * 피드백 저장
 */
function submitFeedback(email, rating, category, message) {
  if (!message) {
    return { success: false, message: '의견을 입력해주세요.' };
  }

  // [보안] 피드백 Rate limiting — 동일 이메일/세션에서 10분 내 중복 방지
  var cache = CacheService.getScriptCache();
  var fbKey = 'fb_' + (email || 'anon').toLowerCase().replace(/[^a-z0-9]/g, '_');
  if (cache.get(fbKey)) {
    return { success: false, message: '피드백은 10분에 1회만 가능합니다.' };
  }
  cache.put(fbKey, '1', 600); // 10분

  // 입력값 살균 — HTML 태그 및 스크립트 제거
  var allowedCategories = ['content','frequency','topic','design','bug','other'];
  var safeEmail = (email || '').toLowerCase().trim().substring(0, 254);
  var safeRating = Math.min(Math.max(parseInt(rating) || 0, 0), 5);
  var safeCategory = allowedCategories.indexOf(category) > -1 ? category : 'other';
  var safeMessage = stripHtml(message).substring(0, 2000);

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('feedback');
  if (!sheet) {
    sheet = ss.insertSheet('feedback');
    sheet.getRange(1, 1, 1, 6).setValues([['timestamp', 'email', 'rating', 'category', 'message', 'status']]);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    new Date().toISOString(),
    safeEmail,
    safeRating,
    safeCategory,
    safeMessage,
    'new'
  ]);

  addLog('FEEDBACK', maskEmail(safeEmail) + ' - ' + safeCategory + ' (⭐' + safeRating + ')');

  return { success: true, message: '피드백이 접수되었습니다. 감사합니다!' };
}

/**
 * 피드백 조회 (관리자)
 */
function getFeedback(password) {
  if (!authenticate(password).success) {
    return { success: false, message: '인증에 실패했습니다.' };
  }

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('feedback');
  if (!sheet) {
    return { success: true, data: [] };
  }

  var data = sheet.getDataRange().getValues();
  var feedbacks = [];
  for (var i = 1; i < data.length; i++) {
    feedbacks.push({
      timestamp: data[i][0],
      email: data[i][1],
      rating: data[i][2],
      category: data[i][3],
      message: data[i][4],
      status: data[i][5]
    });
  }

  // 최신순 정렬
  feedbacks.reverse();

  return { success: true, data: feedbacks };
}
