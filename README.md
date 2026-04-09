# THE AI INSIGHT - 뉴스레터 구독 플랫폼

서연이화 AX추진팀의 AI 트렌드 뉴스레터 구독 관리 시스템입니다.

## 아키텍처

```
┌─────────────────────────────────────────────────┐
│  GitHub Pages (프론트엔드)                        │
│  ├ index.html      구독/구독취소                  │
│  ├ archive.html    지난 뉴스레터 아카이브          │
│  └ admin.html      관리자 대시보드                │
└────────────────┬────────────────────────────────┘
                 │ fetch API
┌────────────────▼────────────────────────────────┐
│  Google Apps Script (백엔드 API)                  │
│  ├ 구독/구독취소 처리                              │
│  ├ 구독자 CRUD                                    │
│  ├ 뉴스레터 아카이브                               │
│  └ 이메일 발송 (Gmail + 외부 API)                  │
└────────────────┬────────────────────────────────┘
                 │
┌────────────────▼────────────────────────────────┐
│  Google Sheets (DB)                              │
│  ├ subscribers    구독자 목록                     │
│  ├ newsletters    뉴스레터 아카이브               │
│  ├ config         설정                           │
│  └ logs           활동 로그                      │
└─────────────────────────────────────────────────┘
```

## 설정 방법

### 1단계: Google Sheets 생성
1. [Google Sheets](https://sheets.google.com)에서 새 스프레드시트 생성
2. 스프레드시트 ID 복사 (URL의 `/d/` 와 `/edit` 사이 문자열)

### 2단계: Google Apps Script 배포
1. [Apps Script](https://script.google.com)에서 새 프로젝트 생성
2. `apps-script/Code.gs` 내용을 붙여넣기
3. `CONFIG.SPREADSHEET_ID`에 위에서 복사한 ID 입력
4. `CONFIG.ADMIN_PASSWORD` 변경 (기본값: `aiinsight2026`)
5. 내부 이메일 6명 목록 `CONFIG.INTERNAL_EMAILS` 수정
6. **배포** → **새 배포** → **웹 앱** 선택
   - 실행 계정: **나**
   - 액세스 권한: **모든 사용자**
7. 배포 URL 복사

### 3단계: 프론트엔드에 API URL 연결
아래 3개 파일에서 `YOUR_APPS_SCRIPT_WEB_APP_URL`을 배포 URL로 교체:
- `index.html` (18번째 줄 근처)
- `archive.html`
- `admin.html`

### 4단계: GitHub Pages 배포
```bash
# 저장소 생성 후
git init
git add .
git commit -m "initial: THE AI INSIGHT newsletter platform"
git remote add origin https://github.com/YOUR_ORG/ai-insight-newsletter.git
git push -u origin main
```
GitHub → Settings → Pages → Source: `main` / `/ (root)` → Save

### 5단계: 초기 데이터 설정
Apps Script 에디터에서 `setupSheets` 함수를 한 번 실행하면 시트가 자동 생성됩니다.

## 이메일 발송 설정

### 현재 구성
| 대상 | 방식 | 비고 |
|------|------|------|
| 내부 6명 | Gmail (GmailApp) | 기존 자동 발송 유지 |
| 외부 구독자 | Gmail (기본) | 일 100통 제한 |

### 외부 API로 전환하기
`Code.gs`의 `CONFIG.EXTERNAL_MAILER.provider`를 변경:

**우리메일** (회사 차단 해제 시):
```javascript
provider: 'woorimail',
woorimail: {
  apiKey: 'YOUR_API_KEY',
  apiUrl: 'https://api.woorimail.com/v1/send',
  senderEmail: 'newsletter@seoyoneh.com',
  senderName: 'THE AI INSIGHT'
}
```

**Mailgun** (대안 1 - 무료 월 1,000통):
```javascript
provider: 'mailgun',
mailgun: {
  apiKey: 'YOUR_API_KEY',
  domain: 'mg.yourdomain.com',
  senderEmail: 'newsletter@yourdomain.com'
}
```

**SendGrid** (대안 2 - 무료 일 100통):
```javascript
provider: 'sendgrid',
sendgrid: {
  apiKey: 'YOUR_API_KEY',
  senderEmail: 'newsletter@yourdomain.com'
}
```

## 파일 구조
```
ai-insight-newsletter/
├── index.html           # 구독/구독취소 페이지
├── archive.html         # 뉴스레터 아카이브
├── admin.html           # 관리자 대시보드 (비밀번호 보호)
├── css/
│   └── shared.css       # 공유 스타일
├── apps-script/
│   └── Code.gs          # Google Apps Script 백엔드
└── README.md            # 이 파일
```

## 관리자 기능
- **통계 대시보드**: 활성/비활성 구독자, 월별 추이, 유입 경로
- **구독자 관리**: 검색, 필터, 삭제
- **뉴스레터 발송**: HTML 에디터로 직접 발송
- **기본 비밀번호**: `AX123`

## 데모 모드
API URL을 연결하지 않으면 샘플 데이터로 동작합니다. 레이아웃과 기능을 미리 확인할 수 있습니다.
