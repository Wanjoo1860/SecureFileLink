# SecureFileLink 제품 계획서 v3.0 (PKCE 전환)

> **버전**: v3.0 | **날짜**: 2026-09-14  
> **핵심 변경**: Client Secret 완전 배제, PKCE + NAA(Nested App Auth) 방식 전환  
> **이전 버전**: v2.3 (Client Secret 기반)

---

## 목차

1. [비전 & 핵심 변경 사항](#1-비전--핵심-변경-사항)
2. [아키텍처 비교 (v2.3 → v3.0)](#2-아키텍처-비교-v23--v30)
3. [사전 준비 체크리스트](#3-사전-준비-체크리스트)
4. [STEP 1: Entra ID 앱 등록 (Public Client)](#4-step-1-entra-id-앱-등록-public-client)
5. [STEP 2: SharePoint 사이트 & 저장소 구성](#5-step-2-sharepoint-사이트--저장소-구성)
6. [STEP 3: Cloudflare Workers 배포](#6-step-3-cloudflare-workers-배포)
7. [STEP 4: Outlook Add-in 개발 (MSAL.js + NAA + PKCE)](#7-step-4-outlook-add-in-개발-msaljs--naa--pkce)
8. [STEP 5: 다운로드 토큰 & Worker 처리 흐름](#8-step-5-다운로드-토큰--worker-처리-흐름)
9. [STEP 6: 관리자 포털](#9-step-6-관리자-포털)
10. [STEP 7: 자동화 (Cron)](#10-step-7-자동화-cron)
11. [STEP 8: 테스트 체크리스트](#11-step-8-테스트-체크리스트)
12. [STEP 9: 선택적 M365 보안 강화](#12-step-9-선택적-m365-보안-강화)
13. [STEP 10: 운영 가이드](#13-step-10-운영-가이드)
14. [API 엔드포인트 전체 목록](#14-api-엔드포인트-전체-목록)
15. [SharePoint Lists 스키마](#15-sharepoint-lists-스키마)
16. [시스템 제한 & 비용](#16-시스템-제한--비용)
17. [로드맵](#17-로드맵)
18. [버전 히스토리](#18-버전-히스토리)

---

## 1. 비전 & 핵심 변경 사항

### 제품 비전 (변경 없음)

Outlook Add-in 기반 대용량 파일 전송 솔루션. 파일은 SharePoint/OneDrive에 저장하고, 수신자에게는 보안 다운로드 링크만 메일에 삽입한다. 보안은 M365 인프라(IRM, Sensitivity Labels, DLP, Defender, Conditional Access)에 100% 위임하고, SecureFileLink는 편의 기능(대용량 전송, 만료, 로그)만 담당한다.

### v3.0 핵심 변경: Client Secret 완전 배제

| 항목 | v2.3 (이전) | v3.0 (현재) |
|------|-------------|-------------|
| **Outlook Add-in 인증** | MSAL.js + Delegated (PKCE 미사용) | MSAL.js NAA + **PKCE 필수** |
| **Worker → Graph API** | App-only 토큰 (Client Secret) | **Pre-Auth URL 방식** (Secret 불필요) |
| **앱 등록 유형** | Confidential Client (Web) | **Public Client (SPA)** |
| **비밀 관리 대상** | `CLIENT_SECRET` + `TOKEN_SECRET` | **`TOKEN_SECRET`만** (HMAC 서명용) |
| **Entra 인증서/시크릿** | 필수 (24개월 갱신) | **불필요** |
| **다운로드 URL 획득 시점** | Worker가 실시간 Graph 호출 | **Add-in이 미리 획득 → 토큰에 암호화 포함** |

### 왜 Client Secret을 배제하는가?

1. **보안**: Secret 유출 위험 제거 (Workers 환경 변수에 저장 불필요)
2. **운영 부담 감소**: 24개월마다 Secret 갱신 작업 불필요
3. **권한 최소화**: Application 권한 (Files.Read.All 등) 불필요, Delegated 권한만 사용
4. **제로 트러스트**: 모든 Graph API 호출이 사용자 컨텍스트에서만 발생

### 새로운 다운로드 흐름 개요

```
[발신자 Outlook] → MSAL.js PKCE로 Delegated 토큰 획득
                → Graph API로 파일 업로드
                → Graph API로 @microsoft.graph.downloadUrl 획득
                → downloadUrl을 AES-256 암호화하여 토큰에 포함
                → 수신자에게 링크 메일 전송

[수신자 클릭]   → Cloudflare Worker가 토큰 디코딩
                → HMAC 서명 검증
                → 만료/횟수 확인
                → 암호화된 downloadUrl 복호화
                → 302 리다이렉트 (또는 HTML 자동 다운로드)
```

> **핵심 포인트**: Worker는 더 이상 Graph API를 직접 호출하지 않는다.  
> 대신 Add-in이 미리 획득한 pre-authenticated URL을 복호화하여 리다이렉트한다.

---

## 2. 아키텍처 비교 (v2.3 → v3.0)

### v2.3 아키텍처 (Client Secret 기반)

```
Outlook Add-in ─── MSAL.js (Delegated) ───→ Graph API (업로드)
                                            │
수신자 클릭 ─→ Worker ─── Client Secret ───→ Graph API (downloadUrl 획득)
                     └──→ 302 리다이렉트 ──→ SharePoint CDN
```

### v3.0 아키텍처 (PKCE + Pre-Auth URL)

```
Outlook Add-in ─── MSAL.js NAA + PKCE ───→ Graph API (업로드)
                                           │
              ←── @microsoft.graph.downloadUrl 획득
                                           │
              ───→ downloadUrl AES-256 암호화 → 토큰에 포함 → 메일 전송
                                           
수신자 클릭 ─→ Worker ─── 토큰 디코딩 ───→ downloadUrl AES-256 복호화
                     └──→ 302 리다이렉트 ──→ SharePoint CDN (pre-auth URL)
```

### Pre-Auth URL의 수명 문제와 해결

| 문제 | 설명 | 해결 방안 |
|------|------|----------|
| **downloadUrl 만료** | `@microsoft.graph.downloadUrl`은 **수분~수시간** 후 만료 | 토큰에 `refreshUrl` 필드 추가 — 만료 시 사용자에게 "발신자에게 재전송 요청" 안내 |
| **장기 다운로드 링크** | 7일~30일 유효한 링크 필요 | **SharePoint 공유 링크** (Anonymous/Organization) 생성 → 이 URL은 정책에 따라 장기 유효 |
| **공유 링크 보안** | Anonymous 링크 유출 우려 | HMAC 토큰으로 1차 검증 + 만료/횟수 제한 + SharePoint 공유 링크 만료 설정 병행 |

### 최종 결정: 하이브리드 방식

```
┌─────────────────────────────────────────────────────┐
│  토큰 생성 시 (Add-in에서):                          │
│                                                      │
│  1. Graph API로 공유 링크 생성                        │
│     POST /drives/{id}/items/{id}/createLink          │
│     { type: "view", scope: "anonymous",              │
│       expirationDateTime: "2026-10-01T00:00:00Z" }   │
│                                                      │
│  2. 반환된 공유 URL을 AES-256 암호화                  │
│                                                      │
│  3. 암호화된 URL을 HMAC 토큰에 포함                   │
│                                                      │
│  → 결과: 장기 유효 + Worker에서 Secret 불필요         │
└─────────────────────────────────────────────────────┘
```

---

## 3. 사전 준비 체크리스트

시작하기 전에 아래 항목을 모두 확인하세요.

### 필수 계정 & 라이선스

| 항목 | 필요 여부 | 확인 방법 |
|------|-----------|----------|
| Microsoft 365 Business Basic 이상 | **필수** | https://admin.microsoft.com → 라이선스 |
| SharePoint Online 포함 | **필수** | 위 관리센터에서 SharePoint 활성 확인 |
| 전역 관리자 또는 SharePoint 관리자 | **필수** | 앱 등록, 사이트 생성, 권한 동의에 필요 |
| Cloudflare 계정 (무료) | **필수** | https://dash.cloudflare.com 가입 |
| GitHub 계정 | 선택 | 소스 코드 관리용 |

### 필수 소프트웨어

터미널(명령 프롬프트)을 열고 아래 명령으로 버전을 확인하세요:

```bash
# Node.js v18 이상
node --version

# npm v9 이상
npm --version

# Git
git --version

# Cloudflare Wrangler CLI
wrangler --version
```

**설치가 안 되어 있다면:**

```bash
# Node.js: https://nodejs.org 에서 LTS 버전 다운로드

# Wrangler 설치
npm install -g wrangler

# Yeoman & Office Add-in 제너레이터 설치
npm install -g yo generator-office
```

### 필수 VS Code 확장

- **Microsoft Office Add-in Debugger**
- **Azure Account** (앱 등록 확인용)
- **Cloudflare Workers** (선택)

---

## 4. STEP 1: Entra ID 앱 등록 (Public Client)

> **v2.3과의 차이**: Confidential Client → **Public Client (SPA)**,  
> Client Secret **생성하지 않음**, Application 권한 **추가하지 않음**

### 4.1 앱 등록 생성

1. **https://entra.microsoft.com** 접속 → 좌측 메뉴 **앱 등록** → **새 등록**
2. 아래와 같이 입력:

| 항목 | 값 |
|------|-----|
| 이름 | `SecureFileLink` |
| 지원되는 계정 유형 | **이 조직 디렉터리의 계정만** (Single tenant) |
| 리디렉션 URI 플랫폼 | **SPA (단일 페이지 애플리케이션)** |
| 리디렉션 URI 값 | `brk-multihub://localhost:3000` |

3. **등록** 클릭

### 4.2 필수 정보 기록

등록 후 **개요** 페이지에서 다음 값을 복사하여 메모장에 저장:

```
TENANT_ID     = xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_ID     = yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy
```

> **⚠️ 주의**: Client Secret은 **생성하지 마세요!**  
> "인증서 및 비밀" 메뉴는 건드리지 않습니다.

### 4.3 추가 리디렉션 URI 등록

**인증** 메뉴에서 SPA 리디렉션 URI를 추가합니다:

```
brk-multihub://localhost:3000
brk-multihub://yourdomain.com
https://localhost:3000/taskpane.html       ← 개발용
https://yourdomain.com/taskpane.html       ← 배포용
```

### 4.4 Public Client 활성화

1. **인증** 메뉴 → 하단 **고급 설정**
2. **퍼블릭 클라이언트 흐름 허용** → **예** 선택
3. **저장**

### 4.5 API 권한 추가 (Delegated만!)

**API 권한** 메뉴 → **권한 추가** → **Microsoft Graph** → **위임된 권한**:

| 권한 | 용도 |
|------|------|
| `Files.Read` | 내 파일 읽기 |
| `Files.ReadWrite` | 내 파일 쓰기 |
| `Files.ReadWrite.All` | 공동 폴더 파일 관리 |
| `Sites.ReadWrite.All` | SharePoint 리스트 읽기/쓰기 |
| `Mail.Send` | 메일 전송 |
| `User.Read` | 사용자 프로필 읽기 |
| `offline_access` | Refresh Token 획득 |
| `openid` | OIDC 인증 |
| `profile` | 사용자 이름 등 |

> **⚠️ Application 권한은 추가하지 마세요!**  
> Files.Read.All (Application) 같은 것은 Client Secret/인증서가 필요하므로 사용하지 않습니다.

**관리자 동의 부여** 버튼 클릭 → 확인

### 4.6 검증: 올바르게 설정되었는지 확인

| 확인 항목 | 기대값 |
|-----------|--------|
| 인증서 및 비밀 | **비어 있음** (Secret 없음) |
| 리디렉션 URI 유형 | **SPA** |
| 퍼블릭 클라이언트 흐름 | **예** |
| API 권한 유형 | 모두 **위임된 권한** |
| 관리자 동의 상태 | 모두 **부여됨** |

---

## 5. STEP 2: SharePoint 사이트 & 저장소 구성

> 이 단계는 v2.3과 동일합니다.

### 5.1 SharePoint 사이트 생성

1. **https://admin.microsoft.com** → SharePoint 관리 센터
2. **사이트** → **활성 사이트** → **만들기** → **팀 사이트**
3. 사이트 이름: `SecureFileLink`, URL: `https://회사명.sharepoint.com/sites/securefilelink`

### 5.2 사이트 ID 확인

브라우저에서 아래 URL 접속 (전역 관리자 로그인 상태):

```
https://graph.microsoft.com/v1.0/sites/회사명.sharepoint.com:/sites/securefilelink
```

응답에서 `id` 값을 복사:

```
SITE_ID = 회사명.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy
```

### 5.3 문서 라이브러리 생성

1. 사이트 접속 → **사이트 콘텐츠** → **새로 만들기** → **문서 라이브러리**
2. 이름: `SecureFileLink`

### 5.4 폴더 구조 생성

라이브러리 안에 다음 폴더를 만듭니다:

```
SecureFileLink/              ← 문서 라이브러리
├── _shared/                 ← 공동 폴더 (부서 공유 파일)
├── _mail/                   ← 메일 첨부 업로드
│   └── 2026/
│       └── 09/
│           └── 14/          ← 날짜별 자동 생성
└── _system/                 ← 시스템 설정 백업
```

### 5.5 드라이브 ID 확인

**Graph Explorer** (https://developer.microsoft.com/graph/graph-explorer)에서:

```
GET https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives
```

`SecureFileLink` 라이브러리의 `id` 값을 복사:

```
DRIVE_ID = xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```

### 5.6 SharePoint Lists 생성 (6개)

사이트의 **사이트 콘텐츠** → **새로 만들기** → **목록**에서 각각 생성:

#### ① TransferRecords (전송 기록)

| 열 이름 | 유형 | 설명 |
|---------|------|------|
| Title | 한 줄 텍스트 | 전송 ID (TR-20260914-001) |
| FileName | 한 줄 텍스트 | 파일명 |
| FileSize | 숫자 | 바이트 단위 |
| Sender | 한 줄 텍스트 | 발신자 이메일 |
| Recipients | 여러 줄 텍스트 | 수신자 목록 (JSON) |
| Source | 한 줄 텍스트 | shared / onedrive / mail |
| ShareUrl | 여러 줄 텍스트 | 암호화된 공유 URL |
| ExpireDate | 날짜 및 시간 | 링크 만료일 |
| MaxDownloads | 숫자 | 최대 다운로드 횟수 (0=무제한) |
| Status | 한 줄 텍스트 | active / expired / deleted |
| CreatedDate | 날짜 및 시간 | 생성일 |

#### ② DownloadLogs (다운로드 로그)

| 열 이름 | 유형 | 설명 |
|---------|------|------|
| Title | 한 줄 텍스트 | 로그 ID |
| TransferId | 한 줄 텍스트 | TransferRecords 참조 |
| DownloadIP | 한 줄 텍스트 | 다운로드 IP |
| UserAgent | 여러 줄 텍스트 | 브라우저 정보 |
| Status | 한 줄 텍스트 | started / completed / failed |
| DownloadDate | 날짜 및 시간 | 다운로드 시간 |

#### ③ SharedAssets (공동 폴더 자산)

| 열 이름 | 유형 | 설명 |
|---------|------|------|
| Title | 한 줄 텍스트 | 파일 표시명 |
| DriveItemId | 한 줄 텍스트 | Graph driveItem ID |
| FileName | 한 줄 텍스트 | 원본 파일명 |
| FileSize | 숫자 | 바이트 |
| UploadedBy | 한 줄 텍스트 | 업로더 이메일 |
| Category | 한 줄 텍스트 | 카테고리 분류 |

#### ④ SystemConfig (시스템 설정)

| 열 이름 | 유형 | 설명 |
|---------|------|------|
| Title | 한 줄 텍스트 | 설정 키 |
| Value | 여러 줄 텍스트 | 설정 값 |
| Description | 여러 줄 텍스트 | 설명 |

**초기 데이터 입력:**

| Title (키) | Value (값) | Description |
|------------|-----------|-------------|
| MaxFileSize_Shared | 10737418240 | 공동 폴더 최대 파일 크기 (10 GB) |
| MaxFileSize_Mail | 2147483648 | 메일 업로드 최대 크기 (2 GB) |
| DefaultExpireDays_Mail | 30 | 메일 파일 기본 만료일 |
| MaxDownloadsPerLink | 0 | 링크당 최대 다운로드 (0=무제한) |
| WorkerBaseUrl | https://sfl.회사명.workers.dev | Worker URL |
| ShareLinkExpireDays | 30 | 공유 링크 기본 만료일 |
| ShareLinkScope | anonymous | 공유 링크 범위 (anonymous/organization) |
| AESEncryptionKey | (자동 생성) | AES-256 암호화 키 (공유 URL 암호화용) |

#### ⑤ AdminUsers (관리자)

| 열 이름 | 유형 | 설명 |
|---------|------|------|
| Title | 한 줄 텍스트 | 관리자 이름 |
| Email | 한 줄 텍스트 | 관리자 이메일 |
| Role | 한 줄 텍스트 | superadmin / admin / viewer |

#### ⑥ AuditLog (감사 로그)

| 열 이름 | 유형 | 설명 |
|---------|------|------|
| Title | 한 줄 텍스트 | 이벤트 ID |
| Action | 한 줄 텍스트 | 동작 (upload/download/delete/config_change) |
| Actor | 한 줄 텍스트 | 실행자 이메일 |
| Detail | 여러 줄 텍스트 | 상세 내용 (JSON) |
| EventDate | 날짜 및 시간 | 이벤트 시간 |

---

## 6. 저장소 구조 & 제한

| 저장소 | 위치 | 최대 파일 크기 | 자동 삭제 | 공유 링크 |
|--------|------|---------------|----------|----------|
| `_shared` | SharePoint Library | 10 GB (설정 가능, 최대 250 GB) | 무기한 | Organization 링크 |
| `_mail` | SharePoint `_mail/YYYY/MM/DD/` | 2 GB (설정 가능, 최대 10 GB) | 7/30/60일 선택 | Anonymous 링크 (만료 포함) |
| 개인 폴더 | OneDrive `/SecureFileLink/` | 라이선스 따름 (1-5 TB) | 사용자 관리 | Anonymous 링크 (만료 포함) |
## 6. STEP 3: Cloudflare Workers 배포

### 6.1 Wrangler 로그인

```bash
wrangler login
```

브라우저가 열리면 Cloudflare 계정으로 로그인합니다.

### 6.2 프로젝트 생성

```bash
npm create cloudflare@latest securefilelink-worker -- --type=typescript
cd securefilelink-worker
```

### 6.3 KV Namespace 생성

```bash
wrangler kv namespace create SFL_CACHE
```

출력된 `id` 값을 복사합니다:

```
SFL_CACHE_ID = xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```

### 6.4 wrangler.toml 설정

```toml
name = "securefilelink-worker"
main = "src/index.ts"
compatibility_date = "2026-09-01"

# KV 바인딩
[[kv_namespaces]]
binding = "SFL_CACHE"
id = "위에서_복사한_SFL_CACHE_ID"

# Cron 트리거 (자동화)
[triggers]
crons = [
  "0 * * * *",        # 매시간: 만료 파일 삭제
  "0 2 * * *",        # 매일 02:00 UTC: 빈 폴더 정리
  "0 3 * * 0"         # 매주 일요일 03:00 UTC: 통계 집계
]
```

### 6.5 Secrets 등록

> **v2.3과의 차이**: `CLIENT_SECRET` 없음!

```bash
# HMAC 서명용 비밀 키 (64자 hex 생성)
# PowerShell: -join ((1..64) | ForEach-Object { '{0:x}' -f (Get-Random -Max 16) })
# 또는 Linux/Mac: openssl rand -hex 32

wrangler secret put TOKEN_SECRET
# → 생성한 64자 hex 값 붙여넣기

wrangler secret put AES_KEY
# → 32바이트 (64자 hex) AES-256 키 붙여넣기
# openssl rand -hex 32 로 생성

wrangler secret put SITE_ID
# → 5.2에서 복사한 SITE_ID 값

wrangler secret put DRIVE_ID
# → 5.5에서 복사한 DRIVE_ID 값
```

### 6.6 등록된 Secrets 확인

```bash
wrangler secret list
```

출력 예시:

```
[
  { "name": "TOKEN_SECRET", "type": "secret_text" },
  { "name": "AES_KEY", "type": "secret_text" },
  { "name": "SITE_ID", "type": "secret_text" },
  { "name": "DRIVE_ID", "type": "secret_text" }
]
```

> **확인**: `CLIENT_SECRET`이 목록에 **없어야** 합니다!

### 6.7 배포

```bash
wrangler deploy
```

배포 후 출력된 URL을 기록합니다:

```
WORKER_URL = https://securefilelink-worker.회사명.workers.dev
```

SharePoint SystemConfig 리스트의 `WorkerBaseUrl` 값을 이 URL로 업데이트합니다.

---

## 7. STEP 4: Outlook Add-in 개발 (MSAL.js + NAA + PKCE)

### 7.1 NAA(Nested App Authentication)란?

NAA는 Office Add-in이 호스트 앱(Outlook, Word 등)의 SSO를 활용하여 MSAL.js로 토큰을 획득하는 최신 인증 방식입니다.

| 특징 | 설명 |
|------|------|
| **PKCE 자동 적용** | `createNestablePublicClientApplication`이 내부적으로 PKCE 수행 |
| **Client Secret 불필요** | Public Client Application이므로 Secret 없이 동작 |
| **SSO 지원** | Outlook에 로그인한 계정으로 자동 인증 |
| **Popup Fallback** | Silent 실패 시 팝업으로 인터랙티브 인증 |

### 7.2 프로젝트 생성

```bash
yo office
```

프롬프트에서 선택:

```
? Choose a project type:    Office Add-in Task Pane project
? Choose a script type:     TypeScript
? What do you want to name your add-in?    SecureFileLink
? Which Office client application would you like to support?    Outlook
```

### 7.3 MSAL.js 설치

```bash
cd SecureFileLink
npm install @azure/msal-browser
```

### 7.4 MSAL 초기화 코드

`src/taskpane/taskpane.ts` 파일을 열고 최상단에 추가:

```typescript
import {
  createNestablePublicClientApplication,
  type IPublicClientApplication,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";

// ──────────────────────────────────────────────
// 1. MSAL 인스턴스 (NAA + PKCE)
// ──────────────────────────────────────────────
let msalInstance: IPublicClientApplication | undefined;

const TENANT_ID = "여기에_TENANT_ID_입력";
const CLIENT_ID = "여기에_CLIENT_ID_입력";

async function initMsal(): Promise<IPublicClientApplication> {
  if (!msalInstance) {
    msalInstance = await createNestablePublicClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: `https://login.microsoftonline.com/${TENANT_ID}`,
      },
      cache: {
        cacheLocation: "localStorage",
      },
    });
  }
  return msalInstance;
}
```

> **핵심**: `createNestablePublicClientApplication`은 내부적으로  
> **Authorization Code Flow + PKCE**를 자동 수행합니다.  
> `clientSecret` 파라미터가 **존재하지 않습니다**.

### 7.5 토큰 획득 함수

```typescript
// ──────────────────────────────────────────────
// 2. Graph API 토큰 획득 (Silent → Popup Fallback)
// ──────────────────────────────────────────────
async function getGraphToken(): Promise<string> {
  const msal = await initMsal();
  const scopes = [
    "Files.ReadWrite.All",
    "Sites.ReadWrite.All",
    "Mail.Send",
    "User.Read",
  ];

  try {
    // Silent 시도 (SSO)
    const result = await msal.acquireTokenSilent({ scopes });
    console.log("✅ 토큰 획득 성공 (Silent)");
    return result.accessToken;
  } catch (silentError) {
    if (silentError instanceof InteractionRequiredAuthError) {
      // Popup Fallback
      const result = await msal.acquireTokenPopup({ scopes });
      console.log("✅ 토큰 획득 성공 (Popup)");
      return result.accessToken;
    }
    throw silentError;
  }
}
```

### 7.6 파일 업로드 함수

```typescript
// ──────────────────────────────────────────────
// 3. 파일 업로드 (SharePoint에 직접 업로드)
// ──────────────────────────────────────────────
const DRIVE_ID = "여기에_DRIVE_ID_입력";

async function uploadFile(
  file: File,
  folder: string // "_shared" | "_mail/2026/09/14"
): Promise<{ itemId: string; fileName: string; fileSize: number }> {
  const token = await getGraphToken();
  const uploadPath = `/drives/${DRIVE_ID}/root:/${folder}/${file.name}:/content`;

  if (file.size <= 4 * 1024 * 1024) {
    // 4MB 이하: 단순 PUT
    const res = await fetch(
      `https://graph.microsoft.com/v1.0${uploadPath}`,
      {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": file.type,
        },
        body: file,
      }
    );
    const data = await res.json();
    return { itemId: data.id, fileName: file.name, fileSize: file.size };
  } else {
    // 4MB 초과: Upload Session (chunked)
    const sessionRes = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/root:/${folder}/${file.name}:/createUploadSession`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          item: { "@microsoft.graph.conflictBehavior": "rename" },
        }),
      }
    );
    const session = await sessionRes.json();
    const uploadUrl = session.uploadUrl;

    // 10MB 청크로 분할 업로드
    const chunkSize = 10 * 1024 * 1024;
    let offset = 0;
    let result: any;

    while (offset < file.size) {
      const end = Math.min(offset + chunkSize, file.size);
      const chunk = file.slice(offset, end);

      const chunkRes = await fetch(uploadUrl, {
        method: "PUT",
        headers: {
          "Content-Range": `bytes ${offset}-${end - 1}/${file.size}`,
        },
        body: chunk,
      });
      result = await chunkRes.json();
      offset = end;
    }

    return { itemId: result.id, fileName: file.name, fileSize: file.size };
  }
}
```

### 7.7 공유 링크 생성 (핵심 — Client Secret 대체)

```typescript
// ──────────────────────────────────────────────
// 4. 공유 링크 생성 (downloadUrl 대체)
//    → Worker가 Graph API를 호출할 필요 없음!
// ──────────────────────────────────────────────
async function createShareLink(
  itemId: string,
  expireDays: number = 30,
  scope: "anonymous" | "organization" = "anonymous"
): Promise<string> {
  const token = await getGraphToken();

  const expireDate = new Date();
  expireDate.setDate(expireDate.getDate() + expireDays);

  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/items/${itemId}/createLink`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        type: "view",                          // 읽기 전용
        scope: scope,                          // anonymous = 인증 불필요
        expirationDateTime: expireDate.toISOString(),
      }),
    }
  );

  const data = await res.json();

  // data.link.webUrl = SharePoint 공유 링크 (장기 유효)
  return data.link.webUrl;
}
```

> **이것이 v3.0의 핵심 변경입니다!**
>
> - v2.3: Worker가 `CLIENT_SECRET`으로 App-only 토큰 획득 → Graph API 호출 → `downloadUrl` 획득
> - v3.0: Add-in이 사용자의 **Delegated 토큰**으로 공유 링크 생성 → 토큰에 암호화 포함 → Worker는 복호화만

### 7.8 다운로드 토큰 생성

```typescript
// ──────────────────────────────────────────────
// 5. 다운로드 토큰 생성
//    shareUrl을 AES-256 암호화하여 포함
// ──────────────────────────────────────────────
async function generateDownloadToken(params: {
  transferId: string;
  fileId: string;
  fileName: string;
  fileSize: number;
  sender: string;
  source: "shared" | "onedrive" | "mail";
  shareUrl: string;           // ← createShareLink()의 반환값
  expireDays: number;
  maxDownloads: number;
}): Promise<string> {
  // Worker의 /api/token/generate 엔드포인트 호출
  // Worker가 AES 암호화 + HMAC 서명 수행
  const res = await fetch(`${WORKER_BASE_URL}/api/token/generate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      tid: params.transferId,
      fid: params.fileId,
      fn: params.fileName,
      fs: params.fileSize,
      fr: params.sender,
      src: params.source,
      url: params.shareUrl,     // 평문 공유 URL → Worker가 암호화
      exp_days: params.expireDays,
      mc: params.maxDownloads,
    }),
  });

  const data = await res.json();
  return data.token; // Base64URL 인코딩된 토큰
}
```

### 7.9 메일에 링크 삽입

```typescript
// ──────────────────────────────────────────────
// 6. Outlook 메일 본문에 다운로드 링크 삽입
// ──────────────────────────────────────────────
function insertDownloadLink(
  token: string,
  fileName: string,
  fileSize: number,
  workerBaseUrl: string
): void {
  const downloadUrl = `${workerBaseUrl}/d/${token}`;
  const fileSizeMB = (fileSize / (1024 * 1024)).toFixed(1);

  const html = `
    <div style="border:1px solid #0078d4; border-radius:8px; padding:16px;
                margin:8px 0; max-width:480px; font-family:Segoe UI,sans-serif;">
      <div style="display:flex; align-items:center; gap:8px; margin-bottom:8px;">
        <span style="font-size:24px;">📎</span>
        <strong style="font-size:14px; color:#0078d4;">SecureFileLink</strong>
      </div>
      <div style="font-size:13px; color:#333; margin-bottom:12px;">
        <strong>${fileName}</strong> (${fileSizeMB} MB)
      </div>
      <a href="${downloadUrl}"
         style="display:inline-block; background:#0078d4; color:#fff;
                padding:8px 20px; border-radius:4px; text-decoration:none;
                font-size:13px;">
        📥 파일 다운로드
      </a>
      <div style="font-size:11px; color:#999; margin-top:8px;">
        Powered by SecureFileLink
      </div>
    </div>
  `;

  Office.context.mailbox.item?.body.setSelectedDataAsync(
    html,
    { coercionType: Office.CoercionType.Html },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("✅ 다운로드 링크 삽입 완료");
      }
    }
  );
}
```

### 7.10 NAA 미지원 환경 Fallback

```typescript
// ──────────────────────────────────────────────
// 7. NAA 지원 여부 확인 & Fallback
// ──────────────────────────────────────────────
async function initAuth(): Promise<IPublicClientApplication> {
  const isNAASupported = Office.context.requirements.isSetSupported(
    "NestedAppAuth",
    "1.1"
  );

  if (isNAASupported) {
    console.log("✅ NAA 지원됨 → PKCE 인증");
    return await initMsal();
  } else {
    // 구버전 Outlook: Office Dialog API 사용
    console.log("⚠️ NAA 미지원 → Dialog API Fallback");
    return await initMsalWithDialogFallback();
  }
}

// Dialog Fallback (구버전 Outlook용)
async function initMsalWithDialogFallback(): Promise<IPublicClientApplication> {
  // Office.context.ui.displayDialogAsync를 사용하여
  // 별도 팝업에서 MSAL 인증 수행
  // 상세 구현은 Microsoft 공식 샘플 참조:
  // https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth
  throw new Error("Dialog Fallback 구현 필요");
}
```

### 7.11 전체 흐름 통합 (메인 함수)

```typescript
// ──────────────────────────────────────────────
// 8. 메인: 파일 선택 → 업로드 → 링크 생성 → 삽입
// ──────────────────────────────────────────────
async function handleSendFile(file: File, source: "shared" | "mail"): Promise<void> {
  try {
    // ① MSAL 초기화
    await initAuth();

    // ② 파일 업로드
    const folder = source === "shared" ? "_shared" : `_mail/${getTodayPath()}`;
    const uploaded = await uploadFile(file, folder);
    console.log(`📤 업로드 완료: ${uploaded.itemId}`);

    // ③ 공유 링크 생성 (Graph API - Delegated)
    const shareUrl = await createShareLink(uploaded.itemId, 30, "anonymous");
    console.log(`🔗 공유 링크: ${shareUrl}`);

    // ④ 전송 ID 생성
    const transferId = `TR-${new Date().toISOString().slice(0,10).replace(/-/g,"")}-${String(Math.random()).slice(2,5)}`;

    // ⑤ 다운로드 토큰 생성 (Worker 호출)
    const token = await generateDownloadToken({
      transferId,
      fileId: uploaded.itemId,
      fileName: uploaded.fileName,
      fileSize: uploaded.fileSize,
      sender: Office.context.mailbox.userProfile.emailAddress,
      source,
      shareUrl,
      expireDays: 30,
      maxDownloads: 0,
    });

    // ⑥ 메일에 링크 삽입
    insertDownloadLink(token, uploaded.fileName, uploaded.fileSize, WORKER_BASE_URL);

  } catch (error) {
    console.error("❌ 파일 전송 실패:", error);
  }
}

function getTodayPath(): string {
  const d = new Date();
  return `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,"0")}/${String(d.getDate()).padStart(2,"0")}`;
}
```

### 7.12 Manifest 등록 & 배포

**개발 테스트:**

```bash
npm start
```

**배포 (관리자):**

1. https://admin.microsoft.com → **설정** → **통합 앱** → **사용자 지정 앱 업로드**
2. `manifest.xml` (또는 Unified Manifest JSON) 업로드
3. 배포 대상 사용자/그룹 선택

---

## 8. STEP 5: 다운로드 토큰 & Worker 처리 흐름

### 8.1 토큰 구조 (v3.0 — 변경됨)

```json
{
  "tid": "TR-20260914-001",
  "fid": "01ABCDEF12345678",
  "fn": "설계도면_v3.pdf",
  "fs": 52428800,
  "fr": "kim@company.com",
  "src": "shared",
  "eurl": "AES-256-GCM 암호화된 공유 URL (Base64)",
  "exp": 1697328000,
  "iat": 1726272000,
  "mc": 0,
  "sig": "HMAC-SHA256 서명"
}
```

| 필드 | v2.3 | v3.0 | 변경 사항 |
|------|------|------|----------|
| `eurl` | 없음 | **새로 추가** | 암호화된 SharePoint 공유 URL |
| `fid` | Graph driveItem ID | 동일 | 기록용 (Graph 직접 호출 안 함) |
| `sig` | HMAC-SHA256 | 동일 | 변경 없음 |

### 8.2 토큰 생성 API (Worker)

```
POST /api/token/generate
```

Worker 처리 로직:

```
입력: { tid, fid, fn, fs, fr, src, url, exp_days, mc }
  ↓
① shareUrl(url)을 AES-256-GCM으로 암호화 → eurl
  ↓
② exp = now + exp_days (Unix timestamp)
  ↓
③ JSON 조립: { tid, fid, fn, fs, fr, src, eurl, exp, iat, mc }
  ↓
④ sig 제외한 JSON을 HMAC-SHA256(TOKEN_SECRET)으로 서명 → sig
  ↓
⑤ 전체 JSON을 Base64URL 인코딩 → 토큰 문자열 반환
```

### 8.3 다운로드 흐름 (Worker)

```
수신자가 https://sfl.company.com/d/{token} 클릭
  ↓
① 토큰 Base64URL 디코딩 → JSON 파싱
  ↓
② HMAC-SHA256 서명 검증
   ✗ → 403 Forbidden ("유효하지 않은 링크")
  ↓
③ 만료 확인 (exp < now)
   ✗ → 410 Gone ("링크가 만료되었습니다")
  ↓
④ 다운로드 횟수 확인 (mc > 0 && 현재 횟수 >= mc)
   ✗ → 429 Too Many ("다운로드 횟수 초과")
  ↓
⑤ AES-256-GCM 복호화 → 원본 SharePoint 공유 URL
  ↓
⑥ DownloadLogs에 기록 (SharePoint List API → Delegated 불필요, Worker에서 직접 REST 호출)
  ↓
⑦ HTML 응답 반환:
   - 파일명, 크기, 발신자 표시
   - 자동 다운로드 JS (window.location = shareUrl)
   - 수동 다운로드 버튼 (백업)
  ↓
⑧ 사용자 브라우저가 SharePoint 공유 URL로 리다이렉트
   → SharePoint가 파일 직접 제공 (인증 불필요 — anonymous 링크)
```

### 8.4 Worker에서 SharePoint List 기록 문제 & 해결

> **문제**: Worker에 Client Secret이 없으므로 Graph API(Application)를 직접 호출할 수 없다.  
> 로그를 어떻게 기록하는가?

**해결 방안: SharePoint REST API + API Key**

SharePoint 리스트는 별도의 **Azure Function**이나 **Power Automate**를 통해 기록할 수도 있지만,  
가장 단순한 방법은 **Worker KV에 임시 저장 → 관리자 포털에서 일괄 동기화**입니다.

```
┌─────────────────────────────────────────────┐
│  다운로드 로그 기록 흐름                      │
│                                              │
│  수신자 클릭 → Worker                         │
│    ↓                                          │
│  KV에 로그 임시 저장                           │
│    key: "log:{timestamp}:{tid}"              │
│    value: { ip, ua, status, date }           │
│    ↓                                          │
│  Cron (매시간) or 관리자 포털 접속 시          │
│    ↓                                          │
│  관리자의 Delegated 토큰으로                   │
│  KV 로그 → SharePoint List 일괄 기록          │
│    ↓                                          │
│  기록 완료 후 KV에서 삭제                      │
└─────────────────────────────────────────────┘
```

**대안 (즉시 기록이 필요한 경우):**

Worker에서 SharePoint REST API를 **API Key 없이** 호출하는 방법:

```
POST https://회사명.sharepoint.com/sites/securefilelink/_api/web/lists/getbytitle('DownloadLogs')/items
```

이 호출에는 Bearer 토큰이 필요하므로, 다음 중 하나를 선택:

| 방법 | 장점 | 단점 |
|------|------|------|
| **KV 임시 저장 + 일괄 동기화** (권장) | Secret 완전 불필요, 가장 단순 | 실시간 아님 (최대 1시간 지연) |
| **Power Automate HTTP Webhook** | Secret 불필요, 실시간 | Power Automate 라이선스 필요 |
| **Azure Function + Managed Identity** | Secret 불필요, 실시간 | Azure 인프라 추가 필요 |

> **이 계획서에서는 "KV 임시 저장 + 일괄 동기화"를 기본 방식으로 채택합니다.**

---

## 9. STEP 6: 관리자 포털

> v2.3과 대부분 동일하나, 관리자의 **Delegated 토큰**으로 모든 Graph 호출을 수행합니다.

### 9.1 프로젝트 생성

```bash
npm create vite@latest securefilelink-admin -- --template react-ts
cd securefilelink-admin
npm install @azure/msal-browser @azure/msal-react
```

### 9.2 MSAL 설정 (관리자 포털용)

```typescript
// src/authConfig.ts
export const msalConfig = {
  auth: {
    clientId: "동일한_CLIENT_ID",
    authority: "https://login.microsoftonline.com/TENANT_ID",
    redirectUri: "https://admin.sfl.회사명.com",
  },
  cache: { cacheLocation: "localStorage" },
};

export const graphScopes = [
  "Sites.ReadWrite.All",
  "Files.ReadWrite.All",
  "User.Read",
];
```

### 9.3 관리자 포털 핵심 기능

| 메뉴 | 기능 | Graph API 호출 |
|------|------|---------------|
| 대시보드 | 전송 통계, 최근 활동 | SharePoint List 조회 |
| 파일 관리 | 공동 폴더 파일 목록, 삭제 | `GET /drives/{id}/root:/_shared:/children` |
| 전송 기록 | TransferRecords 조회, 필터 | SharePoint List 조회 |
| 다운로드 로그 | DownloadLogs 조회, CSV 내보내기 | SharePoint List 조회 |
| KV 로그 동기화 | Worker KV → SharePoint List | Worker API + Graph API |
| 시스템 설정 | SystemConfig 편집 | SharePoint List CRUD |
| 관리자 관리 | AdminUsers 추가/삭제 | SharePoint List CRUD |
| 보안 가이드 | M365 보안 설정 안내 | 정적 페이지 |

### 9.4 Cloudflare Pages 배포

```bash
npm run build
wrangler pages deploy dist --project-name=securefilelink-admin
```

### 9.5 Cloudflare Access 보호

1. Cloudflare Dashboard → **Zero Trust** → **Access** → **Applications**
2. **Add an Application** → Self-hosted
3. 도메인: `admin.sfl.회사명.com`
4. Policy: **Allow** → Identity Provider: **Microsoft Azure AD**
5. 이메일 조건: AdminUsers 리스트의 이메일만 허용

---

## 10. STEP 7: 자동화 (Cron)

### 10.1 매시간 — 만료 파일 삭제 & KV 로그 동기화

```
Trigger: 0 * * * *

작업:
1. KV에서 "log:*" 키 스캔 → 리스트 수집
2. (관리자 포털의 캐시된 토큰 또는 별도 메커니즘으로)
   SharePoint DownloadLogs에 일괄 기록
3. 기록 완료된 KV 키 삭제
4. SystemConfig에서 만료일 설정 읽기
5. TransferRecords에서 만료된 항목 조회
6. 해당 파일의 SharePoint 공유 링크 제거
7. TransferRecords Status → "expired" 업데이트
```

> **참고**: Cron에서 SharePoint 기록을 위해서는 관리자가 미리 refresh token을  
> Worker KV에 안전하게 저장해두는 방식 또는 Power Automate Webhook을 사용합니다.
>
> **가장 단순한 방식**: 관리자 포털 접속 시 "KV 동기화" 버튼을 수동 클릭

### 10.2 매일 02:00 UTC — 빈 폴더 정리

```
Trigger: 0 2 * * *

작업:
1. _mail/ 하위의 빈 날짜 폴더 정리
```

### 10.3 매주 일요일 03:00 UTC — 통계 집계

```
Trigger: 0 3 * * 0

작업:
1. 주간 전송 건수, 총 용량, 다운로드 횟수 집계
2. KV에 "stats:weekly:YYYY-WW" 키로 저장
3. 관리자 포털 대시보드에서 표시
```

---

## 11. STEP 8: 테스트 체크리스트

### 11.1 인증 테스트

| # | 테스트 항목 | 기대 결과 | 확인 |
|---|-----------|----------|------|
| 1 | Add-in에서 MSAL 초기화 | NAA 성공, PKCE 사용 확인 | ☐ |
| 2 | acquireTokenSilent | SSO로 자동 토큰 획득 | ☐ |
| 3 | acquireTokenPopup | 팝업 인증 성공 | ☐ |
| 4 | Client Secret 미사용 확인 | Entra 앱 등록에 Secret 없음 | ☐ |
| 5 | Application 권한 미사용 확인 | Delegated 권한만 존재 | ☐ |

### 11.2 파일 업로드 테스트

| # | 테스트 항목 | 기대 결과 | 확인 |
|---|-----------|----------|------|
| 6 | 1 MB 파일 업로드 | 단순 PUT 성공 | ☐ |
| 7 | 100 MB 파일 업로드 | Upload Session 성공 | ☐ |
| 8 | 2 GB 파일 업로드 (mail) | Upload Session 성공 | ☐ |
| 9 | 10 GB 파일 업로드 (shared) | Upload Session 성공 | ☐ |

### 11.3 공유 링크 & 토큰 테스트

| # | 테스트 항목 | 기대 결과 | 확인 |
|---|-----------|----------|------|
| 10 | 공유 링크 생성 | anonymous view 링크 반환 | ☐ |
| 11 | 토큰 생성 | AES 암호화 + HMAC 서명 포함 | ☐ |
| 12 | 토큰 크기 | URL 길이 2,000자 미만 | ☐ |

### 11.4 다운로드 테스트

| # | 테스트 항목 | 기대 결과 | 확인 |
|---|-----------|----------|------|
| 13 | 유효한 토큰 다운로드 | HTML 페이지 → 자동 다운로드 | ☐ |
| 14 | 만료된 토큰 | 410 Gone 에러 페이지 | ☐ |
| 15 | 변조된 토큰 | 403 Forbidden 에러 페이지 | ☐ |
| 16 | 횟수 초과 토큰 | 429 에러 페이지 | ☐ |
| 17 | 외부 수신자 (로그인 없이) | 즉시 다운로드 성공 | ☐ |

### 11.5 관리자 포털 테스트

| # | 테스트 항목 | 기대 결과 | 확인 |
|---|-----------|----------|------|
| 18 | Cloudflare Access 로그인 | Entra ID SSO 성공 | ☐ |
| 19 | 대시보드 로딩 | 통계 표시 | ☐ |
| 20 | KV 로그 동기화 | KV → SharePoint 기록 성공 | ☐ |
| 21 | 시스템 설정 변경 | SystemConfig 업데이트 성공 | ☐ |
## 12. STEP 9: 선택적 M365 보안 강화

> v2.3과 동일. 기본은 "인증 없이 즉시 다운로드", 필요 시 아래 보안 옵션 적용.

### 12.1 보안 옵션 비교표

| 보안 수준 | 방법 | 설정 위치 | 효과 |
|-----------|------|----------|------|
| 기본 | HMAC 토큰 + 만료 | SecureFileLink 자체 | 링크 변조 방지, 기간/횟수 제한 |
| 강화 1 | Do-Not-Forward | Outlook 작성 시 선택 | 메일 전달 금지 |
| 강화 2 | Sensitivity Labels | Outlook → 민감도 레이블 | 파일 암호화 + 전달/인쇄/복사 금지 |
| 강화 3 | Exchange Transport Rules | Exchange 관리센터 | 키워드 기반 자동 암호화 |
| 강화 4 | DLP 정책 | Purview 준수 포털 | 주민번호/카드번호 등 자동 차단 |
| 강화 5 | Defender Safe Links | Microsoft 365 Defender | 악성 링크 실시간 검사 |
| 강화 6 | Conditional Access | Entra 관리센터 | 디바이스/IP/위치 기반 접근 제한 |

### 12.2 조직 공유 링크 정책 확인

PKCE 방식에서는 **공유 링크**가 핵심이므로, 조직의 SharePoint 공유 정책을 확인합니다:

1. **SharePoint 관리센터** → **정책** → **공유**
2. 확인 항목:

| 설정 | 권장값 | 이유 |
|------|--------|------|
| 외부 공유 수준 | "새 및 기존 게스트" 이상 | Anonymous 링크 허용 필요 |
| 익명 링크 만료 | 30일 (조직 정책에 맞게) | 무기한 링크 방지 |
| 익명 링크 권한 | 보기만 | 편집 방지 |
| 링크 유형 기본값 | "특정 사용자" | 실수 방지 (Add-in이 명시적으로 anonymous 선택) |

> **⚠️ 중요**: 조직에서 Anonymous 링크가 **비활성화**되어 있으면  
> SecureFileLink v3.0의 외부 전송이 작동하지 않습니다.  
> 이 경우 `scope: "organization"`으로 변경하고, 외부 수신자에게는  
> 게스트 초대가 필요합니다.

---

## 13. STEP 10: 운영 가이드

### 13.1 일상 운영

| 주기 | 작업 | 방법 |
|------|------|------|
| 매일 | 대시보드 확인 | 관리자 포털 접속 |
| 매일 | KV 로그 동기화 확인 | 관리자 포털 → 로그 동기화 버튼 |
| 주간 | 다운로드 로그 검토 | 관리자 포털 → 다운로드 로그 |
| 월간 | Cloudflare 사용량 확인 | dash.cloudflare.com → Workers 분석 |
| 월간 | SharePoint 저장소 용량 확인 | SharePoint 관리센터 |
| 분기 | AdminUsers 목록 점검 | 관리자 포털 → 관리자 관리 |

### 13.2 비밀 키 관리 (v3.0 — 대폭 간소화)

| 비밀 키 | 저장 위치 | 갱신 주기 | 갱신 방법 |
|---------|----------|----------|----------|
| `TOKEN_SECRET` | Worker Secret | 연 1회 권장 | `wrangler secret put TOKEN_SECRET` |
| `AES_KEY` | Worker Secret | 연 1회 권장 | `wrangler secret put AES_KEY` |
| ~~`CLIENT_SECRET`~~ | ~~Worker Secret~~ | ~~24개월~~ | **❌ 삭제됨 — 관리 불필요!** |

> **갱신 시 주의**: `TOKEN_SECRET`이나 `AES_KEY`를 변경하면  
> 기존에 발급된 모든 토큰이 무효화됩니다.  
> 갱신 전 활성 토큰의 만료를 기다리거나, 이전 키로도 검증하는 로직을 추가하세요.

### 13.3 트러블슈팅

| 증상 | 원인 | 해결 |
|------|------|------|
| Add-in에서 "interaction_required" 반복 | 토큰 캐시 만료 | 사용자가 팝업 인증 허용 |
| 다운로드 시 "403 Forbidden" | HMAC 서명 불일치 | TOKEN_SECRET 일치 여부 확인 |
| 다운로드 시 "410 Gone" | 토큰 만료 또는 공유 링크 만료 | 발신자에게 재전송 요청 |
| "공유 링크 생성 실패" | 조직 공유 정책에서 Anonymous 차단 | SharePoint 관리센터에서 정책 변경 |
| 관리자 포털 접근 불가 | Cloudflare Access 정책 불일치 | AdminUsers 이메일 확인 |

---

## 14. API 엔드포인트 전체 목록

### 14.1 공개 엔드포인트 (인증 불필요)

| 메서드 | 경로 | 설명 |
|--------|------|------|
| `GET` | `/d/{token}` | 다운로드 페이지 (토큰 검증 + 리다이렉트) |
| `POST` | `/api/track/{trackId}` | 다운로드 완료 비콘 |

### 14.2 Add-in 전용 엔드포인트 (토큰 검증)

| 메서드 | 경로 | 설명 |
|--------|------|------|
| `POST` | `/api/token/generate` | 다운로드 토큰 생성 (AES 암호화 + HMAC 서명) |
| `GET` | `/api/shared/files` | 공동 폴더 파일 목록 |
| `POST` | `/api/mail/upload-complete` | 메일 업로드 완료 기록 |
| `GET` | `/api/transfers` | 내 전송 기록 조회 |

### 14.3 관리자 전용 엔드포인트 (Cloudflare Access)

| 메서드 | 경로 | 설명 |
|--------|------|------|
| `GET` | `/api/admin/dashboard` | 대시보드 통계 |
| `GET` | `/api/admin/logs` | KV 로그 조회 |
| `POST` | `/api/admin/logs/sync` | KV → SharePoint 로그 동기화 |
| `GET` | `/api/admin/config` | 시스템 설정 조회 |
| `PUT` | `/api/admin/config/{key}` | 시스템 설정 변경 |
| `GET` | `/api/admin/admins` | 관리자 목록 |
| `POST` | `/api/admin/admins` | 관리자 추가 |
| `DELETE` | `/api/admin/admins/{email}` | 관리자 삭제 |
| `GET` | `/api/admin/transfers` | 전체 전송 기록 |
| `GET` | `/api/admin/downloads` | 전체 다운로드 로그 |

---

## 15. SharePoint Lists 스키마

> 섹션 5.6에서 생성한 6개 리스트의 전체 스키마 참조.  
> v3.0에서 추가/변경된 필드:

| 리스트 | 추가 필드 | 설명 |
|--------|----------|------|
| TransferRecords | `ShareUrl` (여러 줄 텍스트) | 암호화된 공유 URL (백업) |
| TransferRecords | `ShareLinkId` (한 줄 텍스트) | Graph 공유 링크 ID (삭제용) |
| SystemConfig | `ShareLinkExpireDays` | 공유 링크 기본 만료일 |
| SystemConfig | `ShareLinkScope` | anonymous / organization |
| SystemConfig | `AESEncryptionKey` | AES-256 키 (Worker와 동일값) |

---

## 16. 시스템 제한 & 비용

### 16.1 SharePoint 제한

| 항목 | 제한 |
|------|------|
| 사이트 저장소 | 25 TB |
| 단일 파일 크기 | 250 GB |
| 리스트 항목 수 | 30,000,000개 |
| 리스트 뷰 임계값 | 5,000개 (인덱싱된 열 사용으로 해결) |
| 익명 공유 링크 만료 | 조직 정책에 따름 (기본 30일) |

### 16.2 Cloudflare 제한 (무료 플랜)

| 항목 | 제한 |
|------|------|
| Workers 요청 | 100,000건/일 |
| Worker 스크립트 크기 | 10 MB |
| KV 저장소 | 1 GB |
| KV 읽기 | 100,000건/일 |
| KV 쓰기 | 1,000건/일 |
| Worker CPU 시간 | 10 ms/요청 |
| Pages 빌드 | 500회/월 |

### 16.3 비용

| 항목 | 비용 |
|------|------|
| Microsoft 365 | 기존 라이선스 사용 (**추가 비용 없음**) |
| Cloudflare Workers/Pages/KV | 무료 플랜으로 충분 (**$0**) |
| Entra ID 앱 등록 | **무료** |
| Client Secret 관리 | **불필요** (v3.0에서 제거) |
| **총 추가 비용** | **$0/월** |

### 16.4 v2.3 대비 운영 비용 변화

| 운영 항목 | v2.3 | v3.0 |
|-----------|------|------|
| Client Secret 갱신 (24개월) | 필요 | **불필요** |
| Application 권한 감사 | 필요 | **불필요** |
| Worker Secret 수 | 5개 | **4개** (1개 감소) |
| 보안 감사 범위 | App + Delegated | **Delegated만** |

---

## 17. 로드맵

### Phase 1: 인프라 & 인증 (4주)

| 주차 | 작업 | 체크 |
|------|------|------|
| 1주 | Entra ID 앱 등록 (Public Client, PKCE) | ☐ |
| 1주 | SharePoint 사이트, 라이브러리, 리스트 생성 | ☐ |
| 2주 | Cloudflare Workers 프로젝트 생성, KV, Secrets | ☐ |
| 2주 | Worker 토큰 생성/검증 API 구현 | ☐ |
| 3주 | AES-256-GCM 암호화/복호화 구현 | ☐ |
| 3주 | HMAC-SHA256 서명/검증 구현 | ☐ |
| 4주 | 다운로드 페이지 (`/d/{token}`) 구현 | ☐ |
| 4주 | 통합 테스트 (토큰 생성 → 다운로드) | ☐ |

### Phase 2: Outlook Add-in 핵심 기능 (6주)

| 주차 | 작업 | 체크 |
|------|------|------|
| 5주 | Yeoman 프로젝트 생성, MSAL NAA 초기화 | ☐ |
| 5주 | PKCE 인증 + 토큰 획득 구현 | ☐ |
| 6주 | 파일 업로드 (단순 PUT + Upload Session) | ☐ |
| 6주 | 공유 링크 생성 (`createLink`) 구현 | ☐ |
| 7주 | 토큰 생성 API 연동 | ☐ |
| 7주 | 메일 본문 링크 삽입 | ☐ |
| 8주 | Task Pane UI (공동 폴더, 내 OneDrive, 업로드) | ☐ |
| 9주 | NAA 미지원 Fallback (Dialog API) | ☐ |
| 10주 | Manifest 작성 & 사이드로딩 테스트 | ☐ |

### Phase 3: 관리자 포털 & 자동화 (4주)

| 주차 | 작업 | 체크 |
|------|------|------|
| 11주 | Vite React 프로젝트, MSAL 연동 | ☐ |
| 11주 | Cloudflare Pages 배포, Access 설정 | ☐ |
| 12주 | 대시보드, 파일 관리, 전송 기록 화면 | ☐ |
| 12주 | KV 로그 동기화 기능 | ☐ |
| 13주 | 시스템 설정, 관리자 관리 화면 | ☐ |
| 13주 | Cron 자동화 3개 구현 | ☐ |
| 14주 | 보안 가이드 페이지 작성 | ☐ |

### Phase 4: 테스트 & 배포 (4주)

| 주차 | 작업 | 체크 |
|------|------|------|
| 15주 | 전체 통합 테스트 (11.1~11.5 체크리스트) | ☐ |
| 16주 | 파일럿 사용자 배포 (5~10명) | ☐ |
| 17주 | 피드백 수집 & 수정 | ☐ |
| 18주 | 전사 배포 & 사용자 교육 | ☐ |

**총 기간: 18주**

---

## 18. 버전 히스토리

| 버전 | 날짜 | 주요 변경 |
|------|------|----------|
| v1.0 | 2026-08-01 | 초기 버전 (자체 인증 체계) |
| v2.0 | 2026-08-15 | 외부 인증 제거, 토큰 방식 전환 |
| v2.1 | 2026-08-20 | 토큰 필드 간소화, M365 보안 위임 |
| v2.2 | 2026-08-25 | Unified Manifest, UI 업데이트 |
| v2.3 | 2026-09-01 | 초보자 가이드, 상세 사양 확정 |
| **v3.0** | **2026-09-14** | **Client Secret 완전 배제, PKCE + NAA 전환, Pre-Auth URL 방식, AES-256 암호화 토큰, KV 로그 임시 저장** |

### v2.3 → v3.0 마이그레이션 체크리스트

기존 v2.3을 운영 중이라면 아래 순서로 전환합니다:

| # | 작업 | 상세 |
|---|------|------|
| 1 | Entra 앱 등록 변경 | SPA 리디렉션 추가, Public Client 활성화 |
| 2 | Application 권한 제거 | Files.Read.All (App), Sites.ReadWrite.All (App) 등 삭제 |
| 3 | Client Secret 삭제 | Entra → 인증서 및 비밀 → 모든 Secret 삭제 |
| 4 | Worker Secrets 변경 | `CLIENT_SECRET` 삭제, `AES_KEY` 추가 |
| 5 | Worker 코드 업데이트 | Graph API 호출 제거 → AES 복호화 + 공유 URL 리다이렉트 |
| 6 | Add-in 코드 업데이트 | MSAL NAA 전환, `createShareLink` 추가 |
| 7 | SharePoint Lists 업데이트 | TransferRecords에 ShareUrl, ShareLinkId 열 추가 |
| 8 | SystemConfig 업데이트 | ShareLinkExpireDays, ShareLinkScope, AESEncryptionKey 추가 |
| 9 | 테스트 | 전체 체크리스트 (섹션 11) 수행 |

---

> **📋 이 문서를 파일로 저장하려면:**
>
> 파트 1 + 파트 2 + 파트 3의 코드 블록 내용을 순서대로 복사하여
> 하나의 텍스트 파일에 붙여넣고 `SecureFileLink_Plan_v3.0_PKCE.md`로 저장하세요.
>
> VS Code, Typora, GitHub 등 마크다운 뷰어에서 정상 렌더링됩니다.
