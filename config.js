/**
 * THE AI INSIGHT - 설정 파일
 *
 * ★ 보안 주의 ★
 * - 이 파일은 GitHub 등에 올리지 마세요 (.gitignore에 추가 필수)
 * - SHEET_ID는 비공개 (구독자 데이터 포함). admin.html에서만 사용됨.
 * - PUBLIC_SHEET_ID는 공개용 (뉴스레터 아카이브만). index/archive에서 사용.
 * - 시트 권한을 반드시 확인하세요 (비공개 시트는 본인만 접근 가능해야 함)
 */
window.ADMIN_CONFIG = {
  // ── 비공개 (관리자 전용 — gviz 접근은 시트 권한으로 차단됨) ──
  SHEET_ID: '1XpalUZ3ap_97U4VPwCPD5603o107WAS9-B6ebxSPtFc',
  // ── 공개용 (뉴스레터 아카이브만 — 구독자 데이터 없음) ──
  PUBLIC_SHEET_ID: '1bPc6XBR6qo7eGFUybuSNp8-fmeIFWT55PJdnE4LQ_QM',
  FORM_ID:  '1a2KI-p3c55Oy9xMJ9seQn_FKQ8Alx3gGRvuBjtleZ9U',
  SUB_FORM_ACTION: 'https://docs.google.com/forms/d/e/1GSfXbBWkE4f5eqqnW1LRfwj6V9snexJTSRbUR2NwOyo/formResponse',
  WEBAPP_URL: 'https://script.google.com/macros/s/AKfycbxX56OBdahXeBP3wACO5AqzlhmPA5m74lYTYtgTZYCsd8EWJk6R6Wla_q7Lpf9A4LkVaA/exec',
  PROXY_URL: 'https://ai-insight-proxy.aa200526seoyoneh.workers.dev',
  // ── 로그인용 비밀번호 해시 (SHA-256) ── 비밀번호 변경 시 해시도 업데이트 필요
  ADMIN_PW_HASH: '796e27a63d95b6b88a733e3daaa362f6a06382774147419ebf195a957f20f32f'
};
