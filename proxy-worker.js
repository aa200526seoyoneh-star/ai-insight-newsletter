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

// Rate limit: IP당 분당 10 POST.
// Cloudflare Worker 인스턴스가 warm인 동안 유효 — 단일 공격자 대상 brute-force/DoS 차단 목적.
// 완전한 분산 일관성이 필요하면 Durable Objects로 옮길 것.
const RATE_LIMIT_MAX = 10;
const RATE_LIMIT_WINDOW_MS = 60 * 1000;
const rateLimitMap = new Map();

function checkRateLimit(ip) {
  const now = Date.now();
  let entry = rateLimitMap.get(ip);
  if (!entry || now >= entry.resetAt) {
    entry = { count: 0, resetAt: now + RATE_LIMIT_WINDOW_MS };
  }
  entry.count += 1;
  rateLimitMap.set(ip, entry);

  // 메모리 누수 방지: 맵이 커지면 만료된 항목 정리
  if (rateLimitMap.size > 5000) {
    for (const [k, v] of rateLimitMap) {
      if (now >= v.resetAt) rateLimitMap.delete(k);
    }
  }

  return {
    allowed: entry.count <= RATE_LIMIT_MAX,
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

    const rl = checkRateLimit(clientIp);
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

    try {
      let appsScriptResponse;
      if (request.method === 'GET') {
        // GET은 쿼리스트링 그대로 Apps Script에 전달 + _clientIp 주입
        const url = new URL(request.url);
        url.searchParams.set('_clientIp', clientIp);
        const forwardUrl = APPS_SCRIPT_URL + '?' + url.searchParams.toString();
        appsScriptResponse = await fetch(forwardUrl, { method: 'GET', redirect: 'follow' });
      } else {
        const originalBody = await request.text();
        let forwardBody = originalBody;
        try {
          const parsed = JSON.parse(originalBody);
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
