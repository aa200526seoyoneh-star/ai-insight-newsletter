/**
 * THE AI INSIGHT - 클라이언트 설정
 *
 * 보안 정책 (2026-04-18):
 * - 이 파일은 공개 GitHub Pages에 그대로 배포됨. 민감 정보 포함 금지.
 * - 비공개 SHEET_ID, WEBAPP_URL, ADMIN_PW_HASH는 서버·PropertiesService로 이동.
 * - 관리자 인증은 서버 authenticate()가 전담 (클라 사전 해시 검증 제거).
 * - 남은 값들은 공개 가능:
 *     PUBLIC_SHEET_ID — 공개 시트(뉴스레터 아카이브) ID
 *     PROXY_URL       — Cloudflare Worker 프록시 (rate limit·CORS로 보호)
 *     FORM_ID / SUB_FORM_ACTION — 구독 폼 fallback 경로 (폼 자체가 공개)
 */
window.ADMIN_CONFIG = {
  PUBLIC_SHEET_ID: '1bPc6XBR6qo7eGFUybuSNp8-fmeIFWT55PJdnE4LQ_QM',
  FORM_ID:  '1a2KI-p3c55Oy9xMJ9seQn_FKQ8Alx3gGRvuBjtleZ9U',
  SUB_FORM_ACTION: 'https://docs.google.com/forms/d/e/1GSfXbBWkE4f5eqqnW1LRfwj6V9snexJTSRbUR2NwOyo/formResponse',
  PROXY_URL: 'https://ai-insight-proxy.aa200526seoyoneh.workers.dev'
};
