/**
 * Cloudflare Worker - Apps Script Proxy
 *
 * 회사 방화벽이 script.google.com을 차단할 때,
 * 이 Worker를 경유하여 Apps Script Web App에 접근합니다.
 *
 * 배포 방법:
 * 1. https://dash.cloudflare.com 가입 (무료)
 * 2. Workers & Pages → Create Worker
 * 3. 이 코드를 붙여넣기 → Deploy
 * 4. 생성된 URL (예: xxx.workers.dev)을 config.js의 PROXY_URL에 입력
 */

const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxX56OBdahXeBP3wACO5AqzlhmPA5m74lYTYtgTZYCsd8EWJk6R6Wla_q7Lpf9A4LkVaA/exec';

// 허용할 Origin (GitHub Pages 주소)
const ALLOWED_ORIGINS = [
  'https://aa200526seoyoneh-star.github.io',
  'http://localhost',
  'http://127.0.0.1'
];

// Rate limit (2-tier):
//  - 일반 액션: IP당 분당 10회
//  - 관리자 액션(인증·시트 조회·발송): IP당 분당 3회, 실패 시 1시간 장기 ban
// Worker 인스턴스 warm 동안 유효. 완전 분산 일관성은 Durable Objects 필요.
const RATE_LIMIT_PUBLIC = 10;
const RATE_LIMIT_ADMIN = 3;
const RATE_LIMIT_WINDOW_MS = 60 * 1000;
const BAN_DURATION_MS = 60 * 60 * 1000;  // 1 hour
const rateLimitMap = new Map();    // ip → { pubCount, adminCount, resetAt }
const banMap = new Map();          // ip → banUntil (ms)

const ADMIN_ACTIONS = new Set([
  'auth', 'getStats', 'getSubscribers', 'deleteSubscriber', 'addSubscriber',
  'changeStatus', 'saveSchedule', 'sendNewsletter', 'testSend', 'sendPromo',
  'getFeedback', 'requestPasswordReset', 'resetPassword'
]);

function detectAction(request, bodyText) {
  // GET: ?action=xxx, POST: body JSON.action
  try {
    const u = new URL(request.url);
    const a = u.searchParams.get('action');
    if (a) return a;
  } catch (_) {}
  if (bodyText) {
    try { return JSON.parse(bodyText).action || ''; } catch (_) {}
  }
  return '';
}

function checkBan(ip) {
  const until = banMap.get(ip);
  if (!until) return false;
  if (Date.now() >= until) { banMap.delete(ip); return false; }
  return true;
}

function checkRateLimit(ip, isAdmin) {
  const now = Date.now();
  let entry = rateLimitMap.get(ip);
  if (!entry || now >= entry.resetAt) {
    entry = { pubCount: 0, adminCount: 0, resetAt: now + RATE_LIMIT_WINDOW_MS };
  }
  if (isAdmin) entry.adminCount += 1; else entry.pubCount += 1;
  rateLimitMap.set(ip, entry);

  if (rateLimitMap.size > 5000) {
    for (const [k, v] of rateLimitMap) {
      if (now >= v.resetAt) rateLimitMap.delete(k);
    }
  }

  const count = isAdmin ? entry.adminCount : entry.pubCount;
  const max = isAdmin ? RATE_LIMIT_ADMIN : RATE_LIMIT_PUBLIC;

  // 관리자 액션을 과도하게 시도하면 장기 ban
  if (isAdmin && entry.adminCount > max * 3) {
    banMap.set(ip, now + BAN_DURATION_MS);
  }

  return {
    allowed: count <= max,
    retryAfter: Math.max(1, Math.ceil((entry.resetAt - now) / 1000))
  };
}

function getCorsHeaders(request) {
  const origin = request.headers.get('Origin') || '';
  const isAllowed = ALLOWED_ORIGINS.some(o => origin.startsWith(o));

  return {
    'Access-Control-Allow-Origin': isAllowed ? origin : ALLOWED_ORIGINS[0],
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Max-Age': '86400',
  };
}

export default {
  async fetch(request, env, ctx) {
    // CORS preflight
    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: getCorsHeaders(request) });
    }

    // GET/POST만 허용 (GET: 구독 상태 조회/구독 요청, POST: 관리자 액션)
    if (request.method !== 'POST' && request.method !== 'GET') {
      return new Response(JSON.stringify({ error: 'Method not allowed' }), {
        status: 405,
        headers: { ...getCorsHeaders(request), 'Content-Type': 'application/json' }
      });
    }

    const clientIp = request.headers.get('cf-connecting-ip') || 'unknown';

    if (checkBan(clientIp)) {
      return new Response(JSON.stringify({
        success: false,
        message: '과도한 인증 시도로 접근이 일시 차단되었습니다. 잠시 후 다시 시도해주세요.'
      }), {
        status: 429,
        headers: { ...getCorsHeaders(request), 'Content-Type': 'application/json' }
      });
    }

    try {
      let bodyText = null;
      if (request.method === 'POST') {
        bodyText = await request.text();
      }
      const action = detectAction(request, bodyText);
      const isAdmin = ADMIN_ACTIONS.has(action);

      const rl = checkRateLimit(clientIp, isAdmin);
      if (!rl.allowed) {
        return new Response(JSON.stringify({
          success: false,
          message: '요청이 너무 많습니다. 잠시 후 다시 시도해주세요.',
          retryAfter: rl.retryAfter
        }), {
          status: 429,
          headers: {
            ...getCorsHeaders(request),
            'Content-Type': 'application/json',
            'Retry-After': String(rl.retryAfter)
          }
        });
      }

      let appsScriptResponse;
      if (request.method === 'GET') {
        const url = new URL(request.url);
        url.searchParams.set('_clientIp', clientIp);
        const forwardUrl = APPS_SCRIPT_URL + '?' + url.searchParams.toString();
        appsScriptResponse = await fetch(forwardUrl, { method: 'GET', redirect: 'follow' });
      } else {
        let forwardBody = bodyText || '';
        try {
          const parsed = JSON.parse(bodyText);
          parsed._clientIp = clientIp;
          forwardBody = JSON.stringify(parsed);
        } catch (_) {}
        appsScriptResponse = await fetch(APPS_SCRIPT_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain' },
          body: forwardBody,
          redirect: 'follow'
        });
      }

      const result = await appsScriptResponse.text();

      return new Response(result, {
        status: 200,
        headers: {
          ...getCorsHeaders(request),
          'Content-Type': 'application/json'
        }
      });
    } catch (err) {
      return new Response(JSON.stringify({ error: err.message }), {
        status: 500,
        headers: { ...getCorsHeaders(request), 'Content-Type': 'application/json' }
      });
    }
  }
};
