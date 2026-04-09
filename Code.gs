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
const CONFIG = {
  SPREADSHEET_ID: '1XpalUZ3ap_97U4VPwCPD5603o107WAS9-B6ebxSPtFc',  // Google Sheets ID
  ADMIN_PASSWORD: 'ax123',                        // 관리자 비밀번호
  INTERNAL_EMAILS: [
    // 테스트용 1명 (정식 운영 시 내부 6명으로 복원)
    '110316@seoyoneh.com'
  ],
  // 외부 발송 API 설정 (우리메일)
  EXTERNAL_MAILER: {
    provider: 'gmail',  // 승인 완료 후 'woorimail'로 변경
    woorimail: {
      authkey: 'c7d7180b1dfb2bafca2d',
      domain: 'seoyoneh.com',
      senderEmail: 'aa200526@seoyoneh.com',
      senderName: 'THE AI INSIGHT'
    }
  }
};

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
    sheet.appendRow([new Date().toISOString(), action, detail]);
  } catch (e) {
    console.log('Log error: ' + e.message);
  }
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
      case 'auth':
        result = authenticate(params.password);
        break;
      default:
        result = { success: false, message: '알 수 없는 요청입니다.' };
    }
  } catch (error) {
    result = { success: false, message: error.message };
    addLog('ERROR', action + ': ' + error.message);
  }

  // JSONP 콜백이 있으면 JSONP로 응답 (file:// CORS 우회)
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  return doGet(e);
}

// ─── 구독 관리 ──────────────────────────────────────────────
function subscribe(email, name, source) {
  if (!email || !validateEmail(email)) {
    return { success: false, message: '올바른 이메일 주소를 입력해주세요.' };
  }

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
  for (let i = 1; i < data.length; i++) {
    subscribers.push({
      email: data[i][0],
      name: data[i][1],
      subscribedAt: data[i][2],
      status: data[i][3],
      source: data[i][4],
      unsubscribedAt: data[i][5]
    });
  }

  return { success: true, data: subscribers };
}

function deleteSubscriber(password, email) {
  if (!authenticate(password).success) {
    return { success: false, message: '인증에 실패했습니다.' };
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

  // 1) 내부 구독자 - Gmail 발송 (기존 유지)
  CONFIG.INTERNAL_EMAILS.forEach(email => {
    try {
      GmailApp.sendEmail(email, '[THE AI INSIGHT] ' + subject, '', {
        htmlBody: htmlContent,
        name: 'THE AI INSIGHT - AX추진팀'
      });
      sentCount++;
    } catch (e) {
      errors.push(email + ': ' + e.message);
    }
  });

  // 2) 외부 구독자 발송
  const sheet = getSheet('subscribers');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === 'active') {
      const email = data[i][0].toString();
      // 내부 이메일은 이미 Gmail로 발송했으므로 스킵
      if (CONFIG.INTERNAL_EMAILS.includes(email)) continue;

      try {
        sendExternalEmail(email, subject, htmlContent);
        sentCount++;
      } catch (e) {
        errors.push(email + ': ' + e.message);
      }
    }
  }

  // 뉴스레터 기록 저장
  const nlSheet = getSheet('newsletters');
  const nlId = 'NL-' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-') +
               String(nlSheet.getLastRow()).padStart(3, '0');
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
    <div style="max-width:600px;margin:0 auto;font-family:'Apple SD Gothic Neo',sans-serif;">
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
function authenticate(password) {
  if (password === CONFIG.ADMIN_PASSWORD) {
    return { success: true };
  }
  return { success: false, message: '비밀번호가 올바르지 않습니다.' };
}

function validateEmail(email) {
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
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
    + '<div style="max-width:640px;margin:0 auto;font-family:Apple SD Gothic Neo,Noto Sans KR,sans-serif;">'
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
