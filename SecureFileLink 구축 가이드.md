# SecureFileLink 구축 가이드

> 이 문서는 개발 경험이 적은 담당자도 처음부터 끝까지 따라 할 수 있도록
> 모든 단계를 "왜 하는지 → 어디서 하는지 → 무엇을 클릭하는지" 수준으로 설명합니다.

| 항목 | 내용 |
|------|------|
| **버전** | v2.3 (Final) |
| **작성일** | 2026-09-01 |
| **상태** | 설계 확정 |
| **대상 독자** | Microsoft 365 기본 사용 경험이 있는 IT 담당자 |

---

## 목차

1. [제품 개요](#1-제품-개요)
2. [전체 구조 한눈에 보기](#2-전체-구조-한눈에-보기)
3. [용어 사전](#3-용어-사전)
4. [사전 준비 — 시작 전 확인 사항](#4-사전-준비--시작-전-확인-사항)
5. [STEP 1 — Entra ID에 앱 등록하기](#5-step-1--entra-id에-앱-등록하기)
6. [STEP 2 — SharePoint 사이트 만들기](#6-step-2--sharepoint-사이트-만들기)
7. [STEP 3 — 문서 라이브러리와 폴더 만들기](#7-step-3--문서-라이브러리와-폴더-만들기)
8. [STEP 4 — SharePoint Lists 만들기](#8-step-4--sharepoint-lists-만들기)
9. [STEP 5 — Cloudflare 계정 및 Workers 설정](#9-step-5--cloudflare-계정-및-workers-설정)
10. [STEP 6 — Worker 코드 배포 (다운로드 처리)](#10-step-6--worker-코드-배포-다운로드-처리)
11. [STEP 7 — Outlook Add‑in 개발 및 배포](#11-step-7--outlook-add-in-개발-및-배포)
12. [STEP 8 — 관리자 포털 만들기](#12-step-8--관리자-포털-만들기)
13. [STEP 9 — Cron 자동화 설정](#13-step-9--cron-자동화-설정)
14. [STEP 10 — 테스트 체크리스트](#14-step-10--테스트-체크리스트)
15. [STEP 11 — M365 보안 설정 (선택 사항)](#15-step-11--m365-보안-설정-선택-사항)
16. [운영 가이드](#16-운영-가이드)
17. [문제 해결 (Troubleshooting)](#17-문제-해결-troubleshooting)
18. [참고 자료 — 기술 사양 요약](#18-참고-자료--기술-사양-요약)
19. [버전 이력](#19-버전-이력)

---

## 1. 제품 개요

### 1.1 이 제품이 무엇인가요?

SecureFileLink는 **이메일로 대용량 파일을 보내는 도구**입니다.

일반적으로 Outlook에서 이메일에 파일을 첨부하면 25 MB 정도가 한계입니다. SecureFileLink를 사용하면 수 GB 크기의 파일도 간편하게 보낼 수 있습니다. 파일 자체를 메일에 첨부하는 것이 아니라, 파일을 회사의 SharePoint 또는 OneDrive에 저장한 뒤 **다운로드 링크**만 메일에 넣어서 보내는 방식입니다.

### 1.2 어떻게 작동하나요?

```
[발신자]                          [수신자]
   │                                │
   │  1. Outlook에서 파일 선택       │
   │  2. SharePoint/OneDrive에 저장  │
   │  3. 다운로드 링크를 메일에 삽입  │
   │  4. 메일 전송 ──────────────▶  │
   │                                │  5. 메일에서 "다운로드" 클릭
   │                                │  6. 파일이 바로 다운로드됨
   │                                │     (로그인 불필요)
```

### 1.3 왜 필요한가요?

**기존 방법의 문제점**

- 이메일 첨부: 25 MB 이상 불가
- 외부 클라우드(구글 드라이브, 드롭박스 등): 회사 데이터가 외부 서버에 저장됨
- USB/외장하드: 분실 위험, 추적 불가

**SecureFileLink의 장점**

- 파일이 회사의 M365 안에만 저장됩니다 (데이터 유출 방지)
- 누가, 언제, 어디서 다운로드했는지 모두 기록됩니다
- 별도 서버를 구매하거나 운영할 필요가 없습니다
- 추가 비용이 $0입니다 (기존 M365 + Cloudflare 무료 플랜)

### 1.4 세 가지 파일 저장소

SecureFileLink는 용도에 따라 3개의 저장 공간을 사용합니다.

| 저장소 | 쉬운 설명 | 사용 예 |
|--------|-----------|---------|
| **공동 폴더** | 부서 전체가 공유하는 파일 보관함 | 카탈로그, 제품 소개서, 가격표 |
| **개인 OneDrive** | 내 OneDrive에 있는 파일을 링크로 전송 | 고객 제안서, 개인 작업 파일 |
| **메일 업로드** | 이 메일을 위해 임시로 올리는 파일 | 설계 도면, 계약서 (자동 삭제됨) |

---

## 2. 전체 구조 한눈에 보기

### 2.1 구성 요소

SecureFileLink는 4개의 구성 요소로 이루어져 있습니다.

```
┌─────────────────────────────────────────────────────────┐
│  ① Outlook Add‑in (사용자가 직접 사용하는 화면)          │
│     - 파일 선택, 업로드, 링크 생성, 메일에 삽입          │
└────────────────────────┬────────────────────────────────┘
                         │ Microsoft Graph API
                         ▼
┌─────────────────────────────────────────────────────────┐
│  ② Microsoft 365 (파일과 데이터 저장)                    │
│     - SharePoint: 공동 폴더, 메일 업로드, 설정/로그      │
│     - OneDrive: 개인 파일                               │
│     - Entra ID: 로그인 인증                             │
└────────────────────────┬────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────┐
│  ③ Cloudflare Workers (다운로드 링크 처리)               │
│     - 링크 클릭 → 토큰 검증 → 파일 URL 획득 → 다운로드  │
│     - 다운로드 로그 기록                                 │
└─────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────┐
│  ④ 관리자 포털 (관리자용 웹사이트)                       │
│     - 대시보드, 파일 관리, 로그 조회, 설정 변경           │
└─────────────────────────────────────────────────────────┘
```

### 2.2 데이터 흐름 (파일 보내기 ~ 받기)

```
발신자(Outlook)                Cloudflare Worker              M365(SharePoint/OneDrive)
     │                              │                              │
     │  1. Add-in에서 파일 선택      │                              │
     │──── 2. 파일 업로드 ──────────────────────────────────────▶  │
     │                              │                              │  3. 파일 저장 완료
     │──── 4. 토큰 생성 요청 ──────▶│                              │
     │◀─── 5. 토큰(=다운로드 링크) ──│                              │
     │  6. 메일에 링크 삽입 후 전송   │                              │
     │                              │                              │
     ·  ·  ·  ·  ·  ·  ·  ·  ·  ·  ·  ·  ·  ·  ·  ·  ·  ·  ·  ·
     │                              │                              │
수신자(브라우저)                      │                              │
     │                              │                              │
     │──── 7. 링크 클릭 ──────────▶ │                              │
     │                              │──── 8. 파일 URL 요청 ──────▶ │
     │                              │◀─── 9. downloadUrl 반환 ──── │
     │◀─── 10. 파일 다운로드 시작 ── │                              │
     │  11. 다운로드 완료             │  12. 로그 기록               │
```

---

## 3. 용어 사전

이 문서에서 자주 사용되는 용어를 정리합니다. 모르는 단어가 나오면 이곳을 참고하세요.

| 용어 | 쉬운 설명 |
|------|-----------|
| **SharePoint Online** | Microsoft 365에 포함된 회사용 파일 저장소. 팀 사이트를 만들어 파일과 목록(리스트)을 관리할 수 있음 |
| **OneDrive for Business** | Microsoft 365에 포함된 개인용 클라우드 저장소. 각 직원에게 1 TB~5 TB가 할당됨 |
| **Entra ID** | Microsoft의 클라우드 로그인 시스템 (이전 이름: Azure Active Directory, Azure AD) |
| **Graph API** | Microsoft 365의 데이터(파일, 사용자, 메일 등)를 프로그램으로 다루는 통로 |
| **Cloudflare Workers** | 코드를 서버 없이 인터넷에서 실행하는 서비스 (서버리스). 무료로 하루 10만 건 처리 가능 |
| **Workers KV** | Cloudflare에서 제공하는 간단한 저장소 (키‑값 쌍). 캐시나 임시 데이터에 사용 |
| **Cloudflare Pages** | 정적 웹사이트를 무료로 호스팅하는 서비스 |
| **Cloudflare Access** | 웹사이트에 로그인 벽을 세우는 서비스 (50명 이하 무료) |
| **HMAC‑SHA256** | 데이터가 변조되지 않았음을 증명하는 서명(디지털 도장) 기술 |
| **토큰 (Token)** | 다운로드 권한 정보를 담은 암호화된 문자열. 링크 URL에 포함됨 |
| **downloadUrl** | Microsoft가 발급하는 사전 인증된 파일 다운로드 주소. 클릭 시점에 생성되며 약 1시간 유효 |
| **Delegated 권한** | 사용자를 대신하여 작업하는 권한 (사용자가 로그인해야 함) |
| **Application 권한** | 앱이 독립적으로 작업하는 권한 (사용자 로그인 없이 동작) |
| **MSAL.js** | Microsoft 로그인을 웹앱에 쉽게 추가해 주는 JavaScript 라이브러리 |
| **SSO** | Single Sign‑On. 한 번 로그인하면 다시 로그인하지 않아도 되는 방식 |
| **청크 업로드** | 큰 파일을 작은 조각(청크)으로 나누어 올리는 방식. 중간에 끊겨도 이어서 올릴 수 있음 |
| **Cron 트리거** | 정해진 시간에 자동으로 코드를 실행하는 예약 기능 |
| **IRM** | Information Rights Management. 메일이나 파일의 전달, 복사, 인쇄를 제한하는 M365 기능 |
| **Sensitivity Labels** | 민감도 레이블. 파일/메일에 "기밀" 등의 표시를 붙이고 자동으로 암호화/보호하는 M365 기능 |
| **DLP** | Data Loss Prevention. 주민번호 등 민감 정보가 외부로 나가는 것을 자동으로 막는 M365 기능 |

---

## 4. 사전 준비 — 시작 전 확인 사항

구축을 시작하기 전에 아래 항목을 확인하세요.

### 4.1 필수 조건

| # | 확인 항목 | 확인 방법 | 필요한 이유 |
|---|-----------|-----------|------------|
| 1 | Microsoft 365 Business 또는 Enterprise 라이선스 보유 | [admin.microsoft.com](https://admin.microsoft.com) → 결제 → 구독 | SharePoint, OneDrive, Graph API를 사용하기 위해 |
| 2 | M365 테넌트의 전역 관리자(Global Admin) 계정 접근 가능 | 관리센터 로그인 가능 여부 | Entra ID 앱 등록, SharePoint 사이트 생성에 관리자 권한 필요 |
| 3 | SharePoint 관리자 권한 | [admin.microsoft.com](https://admin.microsoft.com) → 관리 센터 → SharePoint | 사이트 생성, 문서 라이브러리 설정에 필요 |
| 4 | 개발 환경 준비 (PC) | 아래 소프트웨어 설치 확인 | Add‑in 개발 및 Worker 배포에 필요 |

### 4.2 설치해야 하는 소프트웨어

| # | 소프트웨어 | 용도 | 설치 방법 |
|---|-----------|------|-----------|
| 1 | **Node.js** (v18 이상) | JavaScript 실행 환경 | [nodejs.org](https://nodejs.org) → LTS 버전 다운로드 → 설치 |
| 2 | **npm** (Node.js에 포함) | 패키지(라이브러리) 관리 | Node.js 설치 시 자동 포함 |
| 3 | **Visual Studio Code** | 코드 편집기 | [code.visualstudio.com](https://code.visualstudio.com) → 다운로드 → 설치 |
| 4 | **Git** | 코드 버전 관리 | [git-scm.com](https://git-scm.com) → 다운로드 → 설치 |
| 5 | **Wrangler** (Cloudflare CLI) | Worker 배포 도구 | 터미널에서 `npm install -g wrangler` 실행 |

### 4.3 설치 확인 방법

터미널(명령 프롬프트)을 열고 아래 명령어를 입력하세요.

```bash
node --version
# v18.x.x 이상이 표시되면 정상

npm --version
# 9.x.x 이상이 표시되면 정상

git --version
# git version 2.x.x 이상이 표시되면 정상

wrangler --version
# 3.x.x 이상이 표시되면 정상
```

### 4.4 계정 준비

| # | 계정 | 가입 방법 | 비용 |
|---|------|-----------|------|
| 1 | **Cloudflare 계정** | [dash.cloudflare.com/sign-up](https://dash.cloudflare.com/sign-up) → 이메일/비밀번호 입력 | 무료 |
| 2 | **GitHub 계정** (선택) | [github.com](https://github.com) → Sign up | 무료 (코드 저장소, Pages 연동에 편리) |

### 4.5 사전 준비 체크리스트

아래 표의 모든 항목에 ✅가 되면 다음 단계로 진행하세요.

```
[ ] M365 라이선스 확인
[ ] 전역 관리자 계정 확인
[ ] SharePoint 관리자 권한 확인
[ ] Node.js 설치 완료 (v18+)
[ ] Visual Studio Code 설치 완료
[ ] Git 설치 완료
[ ] Wrangler 설치 완료
[ ] Cloudflare 계정 생성 완료
```

---

## 5. STEP 1 — Entra ID에 앱 등록하기

### 5.1 이 단계에서 하는 일

Microsoft 365에 "SecureFileLink라는 앱이 있어요, 이 앱이 파일을 읽고 쓸 수 있도록 허락해 주세요"라고 등록하는 단계입니다. 이 등록을 해야 Graph API로 SharePoint/OneDrive의 파일을 다룰 수 있습니다.

### 5.2 앱 등록 절차

**1단계: Entra 관리 센터 접속**

1. 브라우저에서 [entra.microsoft.com](https://entra.microsoft.com)에 접속합니다.
2. 전역 관리자 계정으로 로그인합니다.
3. 왼쪽 메뉴에서 **"애플리케이션"** → **"앱 등록"**을 클릭합니다.

**2단계: 새 앱 등록**

1. 상단의 **"+ 새 등록"** 버튼을 클릭합니다.
2. 아래와 같이 입력합니다:

| 입력 항목 | 입력 값 |
|-----------|---------|
| 이름 | `SecureFileLink` |
| 지원되는 계정 유형 | "이 조직 디렉터리의 계정만(단일 테넌트)" 선택 |
| 리디렉션 URI | 플랫폼: `단일 페이지 애플리케이션(SPA)` / URI: `https://localhost:3000` |

3. **"등록"** 버튼을 클릭합니다.

**3단계: 중요 정보 메모**

등록이 완료되면 "개요" 페이지가 표시됩니다. 아래 두 값을 반드시 메모하세요. 이후 모든 단계에서 사용됩니다.

```
애플리케이션(클라이언트) ID: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx  ← 메모!
디렉터리(테넌트) ID:        xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx  ← 메모!
```

### 5.3 API 권한 추가

**왜 하나요?** — 앱이 파일을 읽고, 업로드하고, 메일을 보내려면 각각에 해당하는 "권한"을 부여해야 합니다.

**1단계: 권한 설정 화면 이동**

1. 왼쪽 메뉴에서 **"API 사용 권한"**을 클릭합니다.
2. **"+ 권한 추가"** 버튼을 클릭합니다.
3. **"Microsoft Graph"**를 선택합니다.

**2단계: Delegated(위임) 권한 추가**

"위임된 권한"을 선택하고, 아래 권한을 하나씩 검색하여 체크한 뒤 "권한 추가"를 클릭합니다.

| 권한 이름 | 무엇을 하는 권한인가요? |
|-----------|----------------------|
| `Files.Read` | 사용자의 OneDrive 파일 목록 읽기 |
| `Files.ReadWrite.All` | SharePoint 파일 업로드/수정 |
| `Mail.Send` | 사용자 대신 메일 보내기 |
| `User.Read` | 로그인한 사용자의 이름/이메일 읽기 |

**3단계: Application(애플리케이션) 권한 추가**

"권한 추가" → "Microsoft Graph" → 이번에는 **"애플리케이션 권한"**을 선택합니다. 아래 권한을 검색하여 체크합니다.

| 권한 이름 | 무엇을 하는 권한인가요? |
|-----------|----------------------|
| `Files.Read.All` | 모든 파일의 다운로드 URL 획득 (Worker가 사용) |
| `Sites.ReadWrite.All` | SharePoint 사이트/리스트 읽기·쓰기 (Worker가 사용) |
| `User.Read.All` | 사용자 정보 조회 (관리자 기능) |

**4단계: 관리자 동의**

1. 권한 목록 위에 **"(조직이름)에 대한 관리자 동의 허용"** 버튼을 클릭합니다.
2. "예"를 클릭합니다.
3. 모든 권한의 상태가 **"(조직이름)에 대해 부여됨"**으로 바뀌면 완료입니다.

### 5.4 클라이언트 시크릿 생성

**왜 하나요?** — Cloudflare Worker가 사용자 로그인 없이 Graph API를 호출하려면 "앱 전용 토큰(App‑only Token)"이 필요합니다. 이 토큰을 발급받으려면 클라이언트 시크릿이 필요합니다.

1. 왼쪽 메뉴에서 **"인증서 및 암호"**를 클릭합니다.
2. **"+ 새 클라이언트 암호"**를 클릭합니다.
3. 설명: `SecureFileLink-Worker`, 만료: `24개월` → "추가" 클릭
4. **"값"** 열에 표시된 문자열을 즉시 복사하여 메모합니다.

> ⚠️ **주의**: 이 값은 이 화면을 벗어나면 다시 볼 수 없습니다. 반드시 지금 복사하세요!

```
클라이언트 시크릿: xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx  ← 메모!
```

### 5.5 메모 정리

이 단계에서 확보한 3개의 값을 안전한 곳에 보관하세요.

```
TENANT_ID      = xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_ID      = xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_SECRET  = xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```

### 5.6 완료 체크리스트

```
[ ] Entra ID에 "SecureFileLink" 앱 등록 완료
[ ] Delegated 권한 4개 추가 완료 (Files.Read, Files.ReadWrite.All, Mail.Send, User.Read)
[ ] Application 권한 3개 추가 완료 (Files.Read.All, Sites.ReadWrite.All, User.Read.All)
[ ] 관리자 동의 완료 (모든 권한 "부여됨" 상태)
[ ] TENANT_ID 메모 완료
[ ] CLIENT_ID 메모 완료
[ ] CLIENT_SECRET 메모 완료
```

---

## 6. STEP 2 — SharePoint 사이트 만들기

### 6.1 이 단계에서 하는 일

SecureFileLink의 모든 파일과 데이터를 저장할 전용 SharePoint 사이트를 만듭니다. 이 사이트 안에 파일 보관함(문서 라이브러리)과 데이터 저장소(리스트)가 들어갑니다.

### 6.2 사이트 생성 절차

**1단계: SharePoint 관리 센터 접속**

1. 브라우저에서 [admin.microsoft.com](https://admin.microsoft.com)에 접속합니다.
2. 왼쪽 메뉴에서 **"관리 센터"** → **"SharePoint"**를 클릭합니다.
3. SharePoint 관리 센터가 열립니다.

**2단계: 사이트 생성**

1. 왼쪽 메뉴에서 **"사이트"** → **"활성 사이트"**를 클릭합니다.
2. **"+ 만들기"** 버튼을 클릭합니다.
3. **"팀 사이트"**를 선택합니다.
4. 아래와 같이 입력합니다:

| 입력 항목 | 입력 값 |
|-----------|---------|
| 사이트 이름 | `SecureFileLink` |
| 사이트 주소 | `securefilelink` (자동으로 `https://회사명.sharepoint.com/sites/securefilelink`이 됨) |
| 소유자 | IT 관리자 이메일 입력 |
| 언어 | 한국어 또는 영어 |
| 개인 정보 설정 | "비공개 – 구성원만 이 사이트에 액세스할 수 있음" |

5. **"마침"**을 클릭합니다.

**3단계: 사이트 URL 메모**

사이트가 생성되면 URL을 메모합니다.

```
SHAREPOINT_SITE_URL = https://회사명.sharepoint.com/sites/securefilelink  ← 메모!
```

### 6.3 사이트 ID 확인

Worker가 Graph API로 이 사이트에 접근하려면 사이트 ID가 필요합니다. 브라우저에서 아래 URL을 방문하여 확인합니다.

```
https://graph.microsoft.com/v1.0/sites/회사명.sharepoint.com:/sites/securefilelink
```

> 💡 **팁**: 브라우저에서 직접 방문하면 로그인 후 JSON 데이터가 표시됩니다. `"id"` 항목의 값을 메모하세요. 또는 [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)에서 동일한 URL을 실행해도 됩니다.

```
SHAREPOINT_SITE_ID = 회사명.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx  ← 메모!
```

### 6.4 완료 체크리스트

```
[ ] SharePoint 팀 사이트 "SecureFileLink" 생성 완료
[ ] SHAREPOINT_SITE_URL 메모 완료
[ ] SHAREPOINT_SITE_ID 메모 완료
```

---

## 7. STEP 3 — 문서 라이브러리와 폴더 만들기

### 7.1 이 단계에서 하는 일

파일을 실제로 저장하는 "문서 라이브러리"를 만들고, 그 안에 용도별 폴더를 생성합니다.

### 7.2 문서 라이브러리 구조

```
SecureFileLink (문서 라이브러리)
├── _shared/        ← 공동 폴더: 모든 직원이 공유하는 파일
├── _mail/          ← 메일 업로드: 메일 발송용 임시 파일 (자동 삭제)
│   └── 2026/
│       └── 09/
│           └── 01/   ← 날짜별 하위 폴더 (자동 생성됨)
└── _system/        ← 시스템: 향후 템플릿 등
```

### 7.3 문서 라이브러리 만들기

1. 생성한 SharePoint 사이트(`https://회사명.sharepoint.com/sites/securefilelink`)에 접속합니다.
2. 왼쪽 메뉴에서 **"사이트 콘텐츠"**를 클릭합니다.
3. **"+ 새로 만들기"** → **"문서 라이브러리"**를 클릭합니다.
4. 이름: `SecureFileLink` 입력 → **"만들기"** 클릭

### 7.4 폴더 만들기

문서 라이브러리 안에서 폴더를 만듭니다.

1. 문서 라이브러리 "SecureFileLink"에 들어갑니다.
2. **"+ 새로 만들기"** → **"폴더"** → 이름: `_shared` → "만들기"
3. 같은 방법으로 `_mail` 폴더를 만듭니다.
4. 같은 방법으로 `_system` 폴더를 만듭니다.

### 7.5 문서 라이브러리 ID(Drive ID) 확인

Worker가 파일에 접근하려면 Drive ID가 필요합니다. Graph Explorer에서 아래 URL을 실행합니다.

```
GET https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/drives
```

응답에서 `name`이 `"SecureFileLink"`인 항목의 `"id"` 값을 메모합니다.

```
DRIVE_ID = b!xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx  ← 메모!
```

### 7.6 개인 OneDrive 폴더 안내

개인 OneDrive는 각 사용자가 직접 관리합니다. 사용자에게 아래와 같이 안내하세요:

> OneDrive에 `SecureFileLink` 폴더를 만들어 주세요.
> 이 폴더에 넣은 파일을 SecureFileLink Add‑in에서 선택하여 링크로 보낼 수 있습니다.
> 파일 추가, 삭제, 이름 변경은 OneDrive에서 직접 하시면 됩니다.

### 7.7 완료 체크리스트

```
[ ] 문서 라이브러리 "SecureFileLink" 생성 완료
[ ] _shared 폴더 생성 완료
[ ] _mail 폴더 생성 완료
[ ] _system 폴더 생성 완료
[ ] DRIVE_ID 메모 완료
```

---

## 8. STEP 4 — SharePoint Lists 만들기

### 8.1 이 단계에서 하는 일

SecureFileLink의 데이터(전송 기록, 다운로드 로그, 설정 등)를 저장하는 6개의 SharePoint 리스트를 만듭니다. 리스트는 "엑셀 표"와 비슷하다고 생각하면 됩니다. 각 리스트가 하나의 데이터 테이블 역할을 합니다.

### 8.2 리스트 목록 개요

| # | 리스트 이름 | 역할 | 쉬운 비유 |
|---|------------|------|-----------|
| 1 | TransferRecords | 누가 어떤 파일을 보냈는지 기록 | 택배 발송 대장 |
| 2 | DownloadLogs | 누가 어떤 파일을 다운로드했는지 기록 | 택배 수령 대장 |
| 3 | SharedAssets | 공동 폴더의 파일 정보 | 공유 자료실 목록 |
| 4 | SystemConfig | 시스템 설정값 | 환경 설정 파일 |
| 5 | AdminUsers | 관리자 계정 목록 | 관리자 명부 |
| 6 | AuditLog | 관리자가 한 모든 조작 기록 | 감사 일지 |

### 8.3 리스트 생성 방법 (공통)

모든 리스트를 같은 방법으로 만듭니다.

1. SharePoint 사이트에서 **"사이트 콘텐츠"**를 클릭합니다.
2. **"+ 새로 만들기"** → **"목록(List)"**을 클릭합니다.
3. **"빈 목록"**을 선택합니다.
4. 이름을 입력하고 **"만들기"**를 클릭합니다.
5. 리스트가 열리면 **"+ 열 추가"**를 클릭하여 아래 표의 컬럼을 하나씩 추가합니다.

> 💡 **팁**: "열 추가" 시 "텍스트 한 줄", "숫자", "날짜 및 시간", "선택" 등의 유형을 선택합니다. 아래 표의 "타입" 열을 참고하세요.

### 8.4 리스트 ① TransferRecords (전송 기록)

리스트 이름: `TransferRecords`

| 컬럼 이름 | 타입 | 설명 | 필수 |
|-----------|------|------|------|
| TransferId | 텍스트 한 줄 | 전송 고유 ID (예: `TR-20260901-001`) | ✅ |
| FileId | 텍스트 한 줄 | SharePoint/OneDrive 파일 ID | ✅ |
| FileName | 텍스트 한 줄 | 파일 이름 | ✅ |
| FileSize | 숫자 | 파일 크기 (바이트 단위) | ✅ |
| SenderEmail | 텍스트 한 줄 | 보낸 사람 이메일 | ✅ |
| SenderName | 텍스트 한 줄 | 보낸 사람 이름 | ✅ |
| Recipients | 여러 줄 텍스트 | 받는 사람 이메일 목록 (JSON 형식) | |
| Source | 선택 (shared/personal/mail) | 어떤 저장소에서 보냈는지 | ✅ |
| TokenExp | 날짜 및 시간 | 링크 만료 시각 | ✅ |
| MaxDownloads | 숫자 | 최대 다운로드 횟수 (0=무제한) | |
| CurrentDownloads | 숫자 | 현재까지 다운로드된 횟수 | |
| Status | 선택 (active/expired/deleted/file_deleted) | 현재 상태 | ✅ |
| CreatedAt | 날짜 및 시간 | 생성 시각 | ✅ |
| DriveId | 텍스트 한 줄 | 파일이 저장된 Drive의 ID | ✅ |

> 📌 **인덱스 설정**: 데이터가 많아지면 검색이 느려질 수 있습니다. `TransferId`, `CreatedAt`, `SenderEmail`, `Status` 컬럼에 인덱스를 추가하세요.
> 인덱스 추가 방법: 리스트 설정(⚙️) → "인덱싱된 열" → "새 인덱스 만들기" → 컬럼 선택

### 8.5 리스트 ② DownloadLogs (다운로드 로그)

리스트 이름: `DownloadLogs`

| 컬럼 이름 | 타입 | 설명 | 필수 |
|-----------|------|------|------|
| TrackId | 텍스트 한 줄 | 로그 고유 ID | ✅ |
| TransferRecordId | 텍스트 한 줄 | TransferRecords의 TransferId 참조 | ✅ |
| FileId | 텍스트 한 줄 | 파일 ID | ✅ |
| FileName | 텍스트 한 줄 | 파일 이름 | |
| FileSize | 숫자 | 파일 크기 | |
| SenderEmail | 텍스트 한 줄 | 보낸 사람 이메일 | |
| DownloaderIP | 텍스트 한 줄 | 다운로드한 사람의 IP 주소 | |
| Country | 텍스트 한 줄 | 다운로드 국가 (예: KR, US) | |
| UserAgent | 여러 줄 텍스트 | 브라우저 정보 | |
| RequestTime | 날짜 및 시간 | 다운로드 요청 시각 | ✅ |
| Status | 선택 (initiated/completed/failed/expired/limit_exceeded) | 다운로드 상태 | ✅ |
| DownloadMethod | 선택 (redirect/streaming) | 다운로드 방식 | |
| CompletedTime | 날짜 및 시간 | 다운로드 완료 시각 | |
| ErrorDetail | 텍스트 한 줄 | 오류 내용 (실패 시) | |

> 📌 **인덱스 설정**: `TrackId`, `TransferRecordId`, `RequestTime`, `Status` 컬럼에 인덱스를 추가하세요.

### 8.6 리스트 ③ SharedAssets (공동 폴더 메타데이터)

리스트 이름: `SharedAssets`

| 컬럼 이름 | 타입 | 설명 | 필수 |
|-----------|------|------|------|
| AssetId | 텍스트 한 줄 | 고유 ID | ✅ |
| FileId | 텍스트 한 줄 | SharePoint 파일 ID | ✅ |
| FileName | 텍스트 한 줄 | 파일 이름 | ✅ |
| FileSize | 숫자 | 파일 크기 | ✅ |
| UploaderEmail | 텍스트 한 줄 | 업로드한 사람 이메일 | ✅ |
| UploaderName | 텍스트 한 줄 | 업로드한 사람 이름 | |
| Description | 여러 줄 텍스트 | 파일 설명 | |
| Tags | 텍스트 한 줄 | 태그 (쉼표로 구분) | |
| DownloadCount | 숫자 | 누적 다운로드 횟수 | |
| CreatedAt | 날짜 및 시간 | 업로드 시각 | ✅ |
| DriveId | 텍스트 한 줄 | Drive ID | ✅ |

### 8.7 리스트 ④ SystemConfig (시스템 설정)

리스트 이름: `SystemConfig`

| 컬럼 이름 | 타입 | 설명 | 필수 |
|-----------|------|------|------|
| ConfigKey | 텍스트 한 줄 | 설정 이름 | ✅ |
| ConfigValue | 텍스트 한 줄 | 설정 값 | ✅ |
| Description | 텍스트 한 줄 | 설명 | |
| UpdatedAt | 날짜 및 시간 | 마지막 수정 시각 | |
| UpdatedBy | 텍스트 한 줄 | 수정한 사람 | |

리스트를 만든 후 아래 초기 데이터를 입력합니다. 리스트에서 **"+ 새 항목"**을 클릭하여 한 줄씩 추가하세요.

| ConfigKey | ConfigValue | Description |
|-----------|-------------|-------------|
| `MaxFileSize_Shared` | `10737418240` | 공동 폴더 최대 파일 크기 (10 GB) |
| `MaxFileSize_Mail` | `2147483648` | 메일 업로드 최대 파일 크기 (2 GB) |
| `DefaultExpireDays_Mail` | `30` | 메일 업로드 기본 만료일 |
| `ExpireOptions` | `[7, 30, 60]` | 만료 기간 선택지 (일) |
| `MaxDownloadsPerLink` | `0` | 링크당 최대 다운로드 횟수 (0=무제한) |
| `ChunkSizeMB` | `10` | 업로드 청크 크기 (MB) |
| `AdaptiveChunk` | `true` | 청크 크기 자동 조절 |
| `FallbackStreaming` | `true` | downloadUrl 실패 시 스트리밍 사용 |
| `StreamingMaxSize` | `5368709120` | 스트리밍 최대 크기 (5 GB) |
| `SiteName` | `SecureFileLink` | 사이트 표시 이름 |
| `WorkerBaseUrl` | `https://files.company.com` | Worker 기본 URL (나중에 수정) |

### 8.8 리스트 ⑤ AdminUsers (관리자 목록)

리스트 이름: `AdminUsers`

| 컬럼 이름 | 타입 | 설명 | 필수 |
|-----------|------|------|------|
| AdminId | 텍스트 한 줄 | 고유 ID | ✅ |
| Email | 텍스트 한 줄 | 관리자 이메일 | ✅ |
| DisplayName | 텍스트 한 줄 | 표시 이름 | ✅ |
| Role | 선택 (super_admin/admin/viewer) | 역할 | ✅ |
| AddedAt | 날짜 및 시간 | 등록 시각 | ✅ |
| AddedBy | 텍스트 한 줄 | 등록한 사람 | |

리스트를 만든 후 첫 번째 관리자(본인)를 추가합니다.

### 8.9 리스트 ⑥ AuditLog (감사 로그)

리스트 이름: `AuditLog`

| 컬럼 이름 | 타입 | 설명 | 필수 |
|-----------|------|------|------|
| AuditId | 텍스트 한 줄 | 고유 ID | ✅ |
| Action | 선택 (config_update/admin_add/admin_remove/file_delete/manual_expire) | 수행한 동작 | ✅ |
| ActorEmail | 텍스트 한 줄 | 수행한 사람 이메일 | ✅ |
| TargetType | 텍스트 한 줄 | 대상 유형 (config/admin/file/transfer) | |
| TargetId | 텍스트 한 줄 | 대상 ID | |
| Details | 여러 줄 텍스트 | 변경 내용 (JSON 형식) | |
| Timestamp | 날짜 및 시간 | 수행 시각 | ✅ |
| IP | 텍스트 한 줄 | 수행한 곳의 IP 주소 | |

### 8.10 완료 체크리스트

```
[ ] TransferRecords 리스트 생성 완료 (14개 컬럼)
[ ] DownloadLogs 리스트 생성 완료 (14개 컬럼)
[ ] SharedAssets 리스트 생성 완료 (11개 컬럼)
[ ] SystemConfig 리스트 생성 완료 (5개 컬럼) + 초기 데이터 11건 입력
[ ] AdminUsers 리스트 생성 완료 (6개 컬럼) + 첫 번째 관리자 등록
[ ] AuditLog 리스트 생성 완료 (8개 컬럼)
[ ] 인덱스 설정 완료 (TransferRecords 4개, DownloadLogs 4개)
```

---

## 9. STEP 5 — Cloudflare 계정 및 Workers 설정

### 9.1 이 단계에서 하는 일

외부 수신자가 다운로드 링크를 클릭했을 때 처리하는 서버 역할을 Cloudflare Workers가 담당합니다. 이 단계에서는 Cloudflare 프로젝트를 만들고 기본 설정을 합니다.

### 9.2 Cloudflare 로그인 및 Workers 프로젝트 생성

**1단계: Wrangler 로그인**

터미널을 열고 아래 명령어를 실행합니다.

```bash
wrangler login
```

브라우저가 열리면 Cloudflare 계정으로 로그인하고 "Allow"를 클릭합니다.

**2단계: 프로젝트 생성**

```bash
# 작업할 폴더로 이동
cd ~/projects

# Worker 프로젝트 생성
npm create cloudflare@latest securefilelink-worker

# 질문에 아래와 같이 답변
# - What type of application? → "Hello World" Worker
# - Do you want to use TypeScript? → Yes
# - Do you want to use git? → Yes
# - Do you want to deploy? → No (아직 배포하지 않음)
```

**3단계: 프로젝트 폴더로 이동**

```bash
cd securefilelink-worker
```

### 9.3 Workers KV 네임스페이스 생성

**왜 하나요?** — App‑only 토큰을 캐시하고, 다운로드 횟수를 임시 저장하는 데 사용합니다.

```bash
# KV 네임스페이스 생성
wrangler kv namespace create "SFL_CACHE"

# 결과로 표시되는 id를 메모합니다
# 예: { binding = "SFL_CACHE", id = "xxxxxxxxxxxxxxxxxxxxxxxxxxxx" }
```

### 9.4 wrangler.toml 설정

프로젝트 폴더의 `wrangler.toml` 파일을 VS Code로 열어 아래와 같이 수정합니다.

```toml
name = "securefilelink-worker"
main = "src/index.ts"
compatibility_date = "2024-01-01"

# KV 바인딩
[[kv_namespaces]]
binding = "SFL_CACHE"
id = "여기에_위에서_메모한_KV_id를_붙여넣으세요"

# Cron 트리거 (3개 — 무료 플랜 최대)
[triggers]
crons = [
  "0 * * * *",     # 매시간: 만료 파일 삭제
  "0 2 * * *",     # 매일 02:00 UTC: 정리 작업
  "0 3 * * 0"      # 매주 일요일 03:00 UTC: 통계 집계
]
```

### 9.5 환경 변수(시크릿) 등록

민감한 정보는 코드에 직접 쓰지 않고 환경 변수로 등록합니다.

```bash
# 하나씩 실행합니다. 각 명령 후 값을 입력하라는 프롬프트가 나옵니다.

wrangler secret put TENANT_ID
# → 5장에서 메모한 테넌트 ID 입력

wrangler secret put CLIENT_ID
# → 5장에서 메모한 클라이언트 ID 입력

wrangler secret put CLIENT_SECRET
# → 5장에서 메모한 클라이언트 시크릿 입력

wrangler secret put TOKEN_SECRET
# → 직접 만든 임의의 긴 문자열 입력 (예: openssl rand -hex 32 로 생성)
#   이 값은 다운로드 토큰의 서명(HMAC)에 사용됩니다

wrangler secret put DRIVE_ID
# → 7장에서 메모한 Drive ID 입력

wrangler secret put SITE_ID
# → 6장에서 메모한 SharePoint Site ID 입력
```

> 💡 **TOKEN_SECRET 생성 팁**: 터미널에서 `openssl rand -hex 32`를 실행하면 64자리 랜덤 문자열이 생성됩니다. 이 값을 복사하여 사용하세요. Windows에서 openssl이 없으면 아무 임의 문자열(최소 32자)을 직접 입력해도 됩니다.

### 9.6 완료 체크리스트

```
[ ] Wrangler 로그인 완료
[ ] securefilelink-worker 프로젝트 생성 완료
[ ] KV 네임스페이스 "SFL_CACHE" 생성 완료
[ ] wrangler.toml 설정 완료 (KV 바인딩, Cron 트리거)
[ ] 환경 변수 6개 등록 완료 (TENANT_ID, CLIENT_ID, CLIENT_SECRET, TOKEN_SECRET, DRIVE_ID, SITE_ID)
```

---

## 10. STEP 6 — Worker 코드 배포 (다운로드 처리)

### 10.1 이 단계에서 하는 일

외부 수신자가 다운로드 링크를 클릭했을 때 실행되는 핵심 코드를 작성하고 배포합니다. 이 Worker가 하는 일은 아래 세 가지입니다:

1. 토큰이 올바른지 확인 (변조 방지)
2. Microsoft에서 실제 파일 다운로드 URL을 받아옴
3. 수신자에게 파일을 내려줌

### 10.2 프로젝트 구조

```
securefilelink-worker/
├── src/
│   ├── index.ts          ← 메인 라우터 (요청을 각 핸들러로 분배)
│   ├── handlers/
│   │   ├── download.ts   ← GET /d/{token} 처리
│   │   ├── track.ts      ← POST /api/track/{trackId} 처리
│   │   ├── token.ts      ← POST /api/token/generate 처리
│   │   ├── shared.ts     ← /api/shared/* 처리
│   │   ├── mail.ts       ← /api/mail/upload 처리
│   │   ├── transfers.ts  ← /api/transfers/* 처리
│   │   ├── admin.ts      ← /api/admin/* 처리
│   │   └── cron.ts       ← Cron 트리거 처리
│   ├── lib/
│   │   ├── auth.ts       ← App-only 토큰 발급/캐시
│   │   ├── graph.ts      ← Graph API 호출 유틸리티
│   │   ├── hmac.ts       ← HMAC-SHA256 토큰 생성/검증
│   │   ├── sharepoint.ts ← SharePoint Lists CRUD 유틸리티
│   │   └── helpers.ts    ← 공통 유틸리티 (날짜, ID 생성 등)
│   └── types.ts          ← TypeScript 타입 정의
├── wrangler.toml
├── package.json
└── tsconfig.json
```

### 10.3 핵심 코드 설명 — 다운로드 처리 흐름

`src/handlers/download.ts` 파일이 가장 중요합니다. 이 파일의 동작을 단계별로 설명합니다.

```
수신자가 링크 클릭 (GET /d/{token})
         │
         ▼
┌─ 1. 토큰 디코딩 ────────────────────────────┐
│  Base64URL 문자열을 JSON으로 변환             │
│  { tid, fid, fn, fs, fr, src, exp, iat,     │
│    mc, sig }                                 │
└──────────────────────────┬───────────────────┘
                           │
                           ▼
┌─ 2. 서명 검증 ──────────────────────────────┐
│  sig를 제외한 JSON을 HMAC-SHA256으로 계산    │
│  계산 결과와 sig 비교                        │
│  불일치 → 403 Forbidden ("잘못된 링크")      │
└──────────────────────────┬───────────────────┘
                           │
                           ▼
┌─ 3. 만료 확인 ──────────────────────────────┐
│  exp < 현재 시각 → 410 Gone ("만료된 링크")  │
└──────────────────────────┬───────────────────┘
                           │
                           ▼
┌─ 4. 다운로드 횟수 확인 ─────────────────────┐
│  mc > 0 이고 currentDownloads ≥ mc          │
│  → 429 ("다운로드 횟수 초과")                │
└──────────────────────────┬───────────────────┘
                           │
                           ▼
┌─ 5. 로그 기록 ──────────────────────────────┐
│  DownloadLogs에 'initiated' 상태 기록        │
│  (IP, 국가, UserAgent, 시각)                 │
└──────────────────────────┬───────────────────┘
                           │
                           ▼
┌─ 6. Graph API 호출 ────────────────────────┐
│  App-only 토큰으로:                         │
│  GET /drives/{driveId}/items/{fid}          │
│      ?select=id,@microsoft.graph.downloadUrl │
│                                              │
│  → downloadUrl 획득 (약 1시간 유효)          │
│  → 비어 있으면 Fallback Streaming            │
└──────────────────────────┬───────────────────┘
                           │
                           ▼
┌─ 7. 다운로드 페이지 반환 ───────────────────┐
│  HTML 페이지를 생성하여 반환:                 │
│  - 파일명, 크기, 발신자 표시                  │
│  - JavaScript가 downloadUrl로 자동 다운로드  │
│  - 완료 후 비콘 전송 (로그 업데이트)          │
└──────────────────────────────────────────────┘
```

### 10.4 토큰 구조 (JSON)

다운로드 링크에 포함되는 토큰은 아래와 같은 JSON입니다.

```json
{
  "tid": "TR-20260901-001",
  "fid": "01ABCDEF12345678",
  "fn": "회사소개서_2026.pdf",
  "fs": 12500000,
  "fr": "kim@company.com",
  "src": "shared",
  "exp": 1696118400,
  "iat": 1693526400,
  "mc": 0,
  "sig": "HMAC-SHA256-서명값"
}
```

| 필드 | 의미 | 예시 |
|------|------|------|
| `tid` | 전송 기록 ID | `TR-20260901-001` |
| `fid` | 파일 ID (SharePoint/OneDrive) | `01ABCDEF12345678` |
| `fn` | 파일 이름 | `회사소개서_2026.pdf` |
| `fs` | 파일 크기 (바이트) | `12500000` (약 12 MB) |
| `fr` | 보낸 사람 이메일 | `kim@company.com` |
| `src` | 저장소 종류 | `shared` / `personal` / `mail` |
| `exp` | 만료 시각 (Unix 타임스탬프) | `1696118400` |
| `iat` | 생성 시각 (Unix 타임스탬프) | `1693526400` |
| `mc` | 최대 다운로드 횟수 (0=무제한) | `0` |
| `sig` | HMAC‑SHA256 서명 | (자동 생성됨) |

### 10.5 코드 배포

코드 작성이 완료되면 아래 명령어로 배포합니다.

```bash
# 배포
wrangler deploy

# 배포 완료 후 표시되는 URL을 메모합니다
# 예: https://securefilelink-worker.회사명.workers.dev
```

배포된 URL을 SystemConfig의 `WorkerBaseUrl`에 업데이트합니다.

> 💡 **커스텀 도메인 설정**: `회사명.workers.dev` 대신 `files.company.com` 같은 자체 도메인을 사용하려면 Cloudflare에 해당 도메인을 등록하고 Workers 라우트를 설정합니다. 이 설정은 선택 사항이며, workers.dev 도메인으로도 정상 동작합니다.

### 10.6 완료 체크리스트

```
[ ] 프로젝트 구조 생성 완료 (handlers/, lib/, types.ts)
[ ] 다운로드 핸들러 (download.ts) 구현 완료
[ ] HMAC 토큰 생성/검증 (hmac.ts) 구현 완료
[ ] App-only 토큰 발급/캐시 (auth.ts) 구현 완료
[ ] Graph API 유틸리티 (graph.ts) 구현 완료
[ ] SharePoint Lists CRUD (sharepoint.ts) 구현 완료
[ ] 모든 핸들러(16개 API 라우트) 구현 완료
[ ] Cron 핸들러 구현 완료
[ ] wrangler deploy 성공
[ ] Worker URL 메모 및 SystemConfig 업데이트 완료
```

---

## 11. STEP 7 — Outlook Add‑in 개발 및 배포

### 11.1 이 단계에서 하는 일

사용자가 Outlook에서 파일을 선택하고 다운로드 링크를 메일에 삽입하는 화면(Add‑in)을 만듭니다.

### 11.2 Add‑in 프로젝트 생성

```bash
# 작업 폴더로 이동
cd ~/projects

# Yeoman 제너레이터 설치 (처음 한 번만)
npm install -g yo generator-office

# Add-in 프로젝트 생성
yo office

# 질문에 아래와 같이 답변:
# - Choose a project type → "Office Add-in Task Pane project"
# - Choose a script type → "TypeScript"
# - What do you want to name your add-in? → "SecureFileLink"
# - Which Office client application? → "Outlook"
```

```bash
cd SecureFileLink
npm install
```

### 11.3 Add‑in 화면 구성

Add‑in은 Outlook 메일 작성 화면 오른쪽에 패널(Task Pane)로 표시됩니다. 3개의 탭으로 구성합니다.

```
┌──────────────────────────────────────┐
│  SecureFileLink                      │
│  ┌──────┬──────────┬───────────────┐ │
│  │ 공동  │ 내 OneDrive │ 새 파일 업로드 │ │
│  │ 폴더  │           │              │ │
│  └──────┴──────────┴───────────────┘ │
│                                      │
│  [탭 1: 공동 폴더]                    │
│  ┌──────────────────────────────────┐│
│  │ 📄 카탈로그_2026.pdf    50 MB    ││
│  │    [링크 삽입]                    ││
│  ├──────────────────────────────────┤│
│  │ 📄 제품소개서_v3.pptx   12 MB    ││
│  │    [링크 삽입]                    ││
│  ├──────────────────────────────────┤│
│  │ 📄 가격표_Q3.xlsx       2 MB     ││
│  │    [링크 삽입]                    ││
│  └──────────────────────────────────┘│
│                                      │
│  [탭 2: 내 OneDrive]                 │
│  ┌──────────────────────────────────┐│
│  │ 📄 고객제안서_A사.pdf   8 MB     ││
│  │    [링크 삽입]                    ││
│  ├──────────────────────────────────┤│
│  │ ℹ️ 파일 관리는 OneDrive에서      ││
│  │   직접 수행하세요.                ││
│  └──────────────────────────────────┘│
│                                      │
│  [탭 3: 새 파일 업로드]               │
│  ┌──────────────────────────────────┐│
│  │ 저장 위치: (●) 공동폴더 (○) 메일  ││
│  │                                  ││
│  │ [파일 선택]  선택된 파일 없음      ││
│  │                                  ││
│  │ 만료: [30일 ▼]  (메일 선택 시)    ││
│  │                                  ││
│  │ [업로드 및 링크 삽입]              ││
│  │                                  ││
│  │ ████████████░░░░ 75% (37/50 MB)  ││
│  └──────────────────────────────────┘│
└──────────────────────────────────────┘
```

### 11.4 MSAL.js SSO 설정

Add‑in이 사용자의 M365 계정으로 자동 로그인하도록 MSAL.js를 설정합니다.

```typescript
// src/auth/msalConfig.ts

const msalConfig = {
  auth: {
    clientId: "여기에_CLIENT_ID",          // 5장에서 메모한 값
    authority: "https://login.microsoftonline.com/여기에_TENANT_ID",
    redirectUri: "https://localhost:3000",  // 개발 시
  },
  cache: {
    cacheLocation: "sessionStorage",
  },
};

// 요청할 권한 범위
const loginScopes = {
  scopes: [
    "Files.Read",
    "Files.ReadWrite.All",
    "Mail.Send",
    "User.Read",
  ],
};
```

### 11.5 메일 본문에 삽입되는 HTML

사용자가 "링크 삽입"을 클릭하면, 메일 본문에 아래와 같은 형태의 카드가 삽입됩니다.

```
┌─────────────────────────────────────┐
│ 📎 SecureFileLink                   │
│                                     │
│ 파일명: 회사소개서_2026.pdf          │
│ 크기: 11.9 MB                       │
│ 만료: 2026-10-01                    │
│                                     │
│ ┌────────────┐                      │
│ │  다운로드   │  ← 클릭하면 다운로드  │
│ └────────────┘                      │
└─────────────────────────────────────┘
```

### 11.6 Unified App Manifest 작성

Outlook Add‑in을 배포하려면 매니페스트(manifest) 파일이 필요합니다. Unified App Manifest (v1.3+) 형식을 사용합니다.

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/teams/vDevPreview/MicrosoftTeams.schema.json",
  "manifestVersion": "devPreview",
  "version": "1.0.0",
  "id": "여기에_고유ID_생성",
  "name": {
    "short": "SecureFileLink",
    "full": "SecureFileLink - 대용량 파일 전송"
  },
  "description": {
    "short": "대용량 파일을 안전하게 전송합니다",
    "full": "SharePoint/OneDrive에 파일을 저장하고 보안 다운로드 링크를 메일에 삽입합니다"
  },
  "developer": {
    "name": "회사명",
    "websiteUrl": "https://company.com",
    "privacyUrl": "https://company.com/privacy",
    "termsOfUseUrl": "https://company.com/terms"
  },
  "extensions": [
    {
      "requirements": {
        "capabilities": [
          { "name": "Mailbox", "minVersion": "1.5" }
        ]
      },
      "runtimes": [
        {
          "requirements": { "capabilities": [{ "name": "Mailbox", "minVersion": "1.5" }] },
          "id": "TaskPaneRuntime",
          "type": "general",
          "code": { "page": "https://localhost:3000/taskpane.html" }
        }
      ],
      "ribbons": [
        {
          "contexts": ["mailCompose"],
          "tabs": [
            {
              "builtInTabId": "TabDefault",
              "groups": [
                {
                  "id": "SecureFileLinkGroup",
                  "label": "SecureFileLink",
                  "controls": [
                    {
                      "id": "TaskPaneButton",
                      "type": "button",
                      "label": "파일 전송",
                      "icons": [
                        { "size": 16, "url": "https://localhost:3000/assets/icon-16.png" },
                        { "size": 32, "url": "https://localhost:3000/assets/icon-32.png" },
                        { "size": 80, "url": "https://localhost:3000/assets/icon-80.png" }
                      ],
                      "supertip": {
                        "title": "SecureFileLink",
                        "description": "대용량 파일을 링크로 전송합니다"
                      },
                      "actionId": "TaskPaneAction"
                    }
                  ]
                }
              ]
            }
          ]
        }
      ]
    }
  ]
}
```

### 11.7 Add‑in 배포

**개발 중 테스트 (사이드로딩)**

```bash
npm start
```

Outlook이 열리면서 Add‑in이 자동으로 로드됩니다. 메일 작성 화면에서 "SecureFileLink" 버튼이 보이면 성공입니다.

**조직 전체 배포**

1. [admin.microsoft.com](https://admin.microsoft.com) → **설정** → **통합 앱** → **앱 업로드**
2. 매니페스트 파일을 업로드합니다.
3. 배포 대상을 "전체 조직" 또는 특정 그룹으로 선택합니다.
4. 배포를 완료하면, 대상 사용자의 Outlook에 Add‑in이 자동으로 나타납니다.

### 11.8 완료 체크리스트

```
[ ] Add-in 프로젝트 생성 완료
[ ] 3개 탭 UI 구현 완료 (공동 폴더, 내 OneDrive, 새 파일 업로드)
[ ] MSAL.js SSO 구현 완료
[ ] 파일 업로드 기능 구현 완료 (청크 업로드)
[ ] 토큰 생성 및 링크 삽입 기능 구현 완료
[ ] Unified App Manifest 작성 완료
[ ] 사이드로딩 테스트 성공
[ ] 조직 배포 완료 (또는 파일럿 그룹 배포)
```

---

## 12. STEP 8 — 관리자 포털 만들기

### 12.1 이 단계에서 하는 일

관리자가 전송 현황, 다운로드 로그, 시스템 설정 등을 관리하는 웹사이트를 만듭니다. 이 사이트는 Cloudflare Pages에 무료로 호스팅하고, Cloudflare Access로 접근을 제한합니다.

### 12.2 관리자 포털 프로젝트 생성

```bash
cd ~/projects

# React 프로젝트 생성
npm create vite@latest securefilelink-admin -- --template react-ts

cd securefilelink-admin
npm install
```

### 12.3 주요 화면

```
관리자 포털 화면 구성:

┌───────────────────────────────────────────────────┐
│  SecureFileLink Admin                    [로그아웃] │
├─────────┬─────────────────────────────────────────┤
│         │                                         │
│ 📊 대시보드 │  오늘 전송: 23건  다운로드: 67건      │
│         │  활성 링크: 145개  총 용량: 1.2 TB       │
│ 📁 파일관리 │                                      │
│         │  최근 다운로드                            │
│ 📋 전송기록 │  ┌────────────────────────────────┐  │
│         │  │ 09:23 kim@co.com → 카탈로그.pdf  │  │
│ 📝 로그    │  │ 09:15 lee@co.com → 제안서.pptx  │  │
│         │  │ 08:50 park@co.com → 도면.dwg    │  │
│ ⚙️ 설정   │  └────────────────────────────────┘  │
│         │                                         │
│ 👤 관리자  │                                       │
│         │                                         │
│ 🔒 보안가이드│                                     │
│         │                                         │
└─────────┴─────────────────────────────────────────┘
```

| 화면 | 기능 |
|------|------|
| **대시보드** | 전송/다운로드 건수, 활성 링크 수, 용량 사용량, 최근 로그 표시 |
| **파일 관리** | 공동 폴더 파일 추가/삭제/수정, 메일 업로드 파일 조회/삭제 |
| **전송 기록** | 모든 전송 기록 검색/필터링 (발신자, 파일명, 상태, 날짜) |
| **다운로드 로그** | 모든 다운로드 기록 검색, CSV 내보내기 |
| **시스템 설정** | SystemConfig 값 편집 (변경 시 AuditLog에 자동 기록) |
| **관리자 관리** | 관리자 추가/삭제/역할 변경 |
| **보안 설정 가이드** | M365 보안 기능 설정 방법 안내 (아래 15장 참조) |

### 12.4 Cloudflare Pages에 배포

```bash
# 빌드
npm run build

# Cloudflare Pages에 배포
wrangler pages deploy dist --project-name=securefilelink-admin
```

배포 후 `https://securefilelink-admin.pages.dev` 같은 URL이 생성됩니다.

### 12.5 Cloudflare Access 설정 (접근 제한)

**왜 하나요?** — 관리자 포털은 누구나 접근하면 안 됩니다. Cloudflare Access를 사용하여 M365 계정으로 로그인한 관리자만 접근할 수 있도록 합니다.

1. [dash.cloudflare.com](https://dash.cloudflare.com) → **Zero Trust** → **Access** → **Applications**
2. **"Add an application"** → **"Self-hosted"** 선택
3. 설정:

| 항목 | 값 |
|------|------|
| Application name | `SecureFileLink Admin` |
| Session duration | `24 hours` |
| Application domain | `securefilelink-admin.pages.dev` |

4. **Policy** 추가:

| 항목 | 값 |
|------|------|
| Policy name | `M365 Admins Only` |
| Action | `Allow` |
| Include | Login Methods → "One‑time PIN" 또는 "Azure AD" (OIDC) |
| Include | Emails → AdminUsers 리스트에 있는 이메일 입력 |

5. 설정 완료 후, 포털 URL에 접속하면 로그인 화면이 먼저 표시됩니다.

### 12.6 완료 체크리스트

```
[ ] 관리자 포털 프로젝트 생성 완료
[ ] 7개 화면 구현 완료
[ ] Cloudflare Pages 배포 완료
[ ] Cloudflare Access 설정 완료 (M365 연동)
[ ] 관리자 로그인 테스트 성공
```

---

## 13. STEP 9 — Cron 자동화 설정

### 13.1 이 단계에서 하는 일

아래 3가지 자동화 작업이 정해진 시간에 자동으로 실행되도록 합니다. Cron 트리거는 9장의 `wrangler.toml`에서 이미 설정했으므로, 여기서는 각 작업의 동작을 설명합니다.

### 13.2 자동화 작업 3가지

| # | 스케줄 | 하는 일 | 왜 필요한가요? |
|---|--------|---------|---------------|
| 1 | 매시간 정각 | 만료된 메일 업로드 파일 삭제 | 불필요한 파일이 계속 쌓이지 않도록 |
| 2 | 매일 새벽 2시 (UTC) | 빈 폴더 정리 + OneDrive 파일 상태 동기화 | 깔끔한 폴더 구조 유지 |
| 3 | 매주 일요일 새벽 3시 (UTC) | 주간 통계 집계 + 오래된 로그 아카이브 | 대시보드 성능 유지 |

### 13.3 작업 1: 만료 파일 삭제 (매시간)

```
매시간 실행:
  1. TransferRecords에서 조건 검색:
     - Source = "mail" (메일 업로드 파일만)
     - TokenExp < 현재 시각 (만료됨)
     - Status = "active" (아직 삭제 안 됨)
  2. 검색된 각 파일에 대해:
     - Graph API로 SharePoint 파일 삭제
     - TransferRecords의 Status를 "deleted"로 변경
  3. 만료 3일 이내인 파일의 발신자에게 알림 메일 발송:
     "귀하가 전송한 '파일명.pdf'의 다운로드 링크가
      3일 후 만료됩니다."
```

### 13.4 작업 2: 정리 및 동기화 (매일)

```
매일 02:00 UTC 실행:
  1. _mail/ 폴더 하위에서 빈 날짜 폴더 삭제
     (파일이 모두 삭제된 후 빈 폴더만 남은 경우)
  2. Source = "personal"인 TransferRecords에서:
     - Graph API로 파일 존재 여부 확인
     - 파일이 삭제되었으면 Status를 "file_deleted"로 변경
     (사용자가 OneDrive에서 파일을 삭제한 경우)
```

### 13.5 작업 3: 통계 집계 (매주)

```
매주 일요일 03:00 UTC 실행:
  1. 이번 주 전송/다운로드 통계를 집계하여 KV에 캐시
     (대시보드에서 빠르게 표시하기 위해)
  2. SharedAssets의 DownloadCount를 DownloadLogs 기준으로 동기화
  3. 90일 이상 지난 DownloadLogs를 아카이브 처리
     (리스트 뷰 5,000건 임계값 관리)
```

### 13.6 완료 체크리스트

```
[ ] Cron 핸들러 (cron.ts) 3가지 작업 구현 완료
[ ] 매시간 작업 테스트 성공 (만료 파일 삭제)
[ ] 매일 작업 테스트 성공 (빈 폴더 정리, 상태 동기화)
[ ] 매주 작업 테스트 성공 (통계 집계)
```

---

## 14. STEP 10 — 테스트 체크리스트

### 14.1 기본 동작 테스트

모든 구축이 끝나면 아래 항목을 하나씩 테스트합니다.

```
기본 기능:
[ ] Add-in에서 공동 폴더 파일 목록이 표시되는가?
[ ] Add-in에서 OneDrive 파일 목록이 표시되는가?
[ ] 공동 폴더에 새 파일을 업로드할 수 있는가?
[ ] 메일 업로드(_mail)에 파일을 업로드할 수 있는가?
[ ] "링크 삽입" 클릭 시 메일 본문에 다운로드 카드가 삽입되는가?
[ ] 메일을 정상적으로 보낼 수 있는가?

다운로드 테스트 (외부 수신자 관점):
[ ] 외부 이메일 계정에서 메일을 수신하는가?
[ ] 메일 내 "다운로드" 버튼을 클릭하면 파일이 다운로드되는가?
[ ] 로그인/인증 화면 없이 바로 다운로드되는가?
[ ] 시크릿(프라이빗) 브라우저 창에서도 다운로드되는가?
[ ] 스마트폰에서도 다운로드되는가?
[ ] 회사 네트워크 외부에서도 다운로드되는가?

만료/제한 테스트:
[ ] 만료된 링크를 클릭하면 "만료" 안내 페이지가 표시되는가?
[ ] 다운로드 횟수 제한(mc > 0) 초과 시 "횟수 초과" 안내가 표시되는가?
[ ] 토큰을 임의로 변조하면 "잘못된 링크" 오류가 표시되는가?

로그 테스트:
[ ] 다운로드 후 DownloadLogs에 기록이 생기는가?
[ ] IP, 국가, UserAgent가 올바르게 기록되는가?
[ ] 전송 기록(TransferRecords)에 CurrentDownloads가 증가하는가?
```

### 14.2 대용량 파일 테스트

```
[ ] 100 MB 파일 업로드 및 다운로드 성공
[ ] 1 GB 파일 업로드 및 다운로드 성공
[ ] 5 GB 파일 업로드 및 다운로드 성공
[ ] 10 GB 파일 업로드 및 다운로드 성공 (공동 폴더)
[ ] 업로드 중 네트워크 끊김 후 재시도 시 이어서 업로드되는가?
```

### 14.3 관리자 포털 테스트

```
[ ] 관리자 포털에 로그인할 수 있는가?
[ ] 대시보드에 통계가 표시되는가?
[ ] 공동 폴더 파일을 추가/삭제할 수 있는가?
[ ] 전송 기록을 검색할 수 있는가?
[ ] 다운로드 로그를 CSV로 내보낼 수 있는가?
[ ] SystemConfig 값을 변경할 수 있는가?
[ ] 설정 변경 시 AuditLog에 기록되는가?
[ ] 관리자가 아닌 사용자는 포털에 접근할 수 없는가?
```

### 14.4 자동화(Cron) 테스트

```
[ ] 만료된 메일 업로드 파일이 자동으로 삭제되는가?
[ ] 삭제된 OneDrive 파일의 전송 기록 상태가 "file_deleted"로 변경되는가?
[ ] 대시보드 통계가 주간 집계 후 갱신되는가?
```

---

## 15. STEP 11 — M365 보안 설정 (선택 사항)

### 15.1 이 단계에서 하는 일

기본 상태에서 SecureFileLink는 누구나 다운로드할 수 있도록 동작합니다. 기밀 문서처럼 보안이 필요한 경우에만 아래의 M365 보안 기능을 추가로 설정합니다. 이 설정들은 SecureFileLink와 무관하게 M365 자체 기능입니다.

### 15.2 옵션 A: 전달 방지 (Do Not Forward)

**누가 설정하나요?** — 메일을 보내는 사람이 직접 설정합니다.

**어떻게 하나요?**

1. Outlook에서 새 메일 작성
2. 상단 메뉴에서 **"옵션"** 탭 클릭
3. **"암호화"** (또는 "권한") 버튼 클릭
4. **"전달 금지"** 선택
5. SecureFileLink로 링크 삽입 후 메일 전송

**결과**: 수신자는 메일을 읽고 다운로드할 수 있지만, 메일을 다른 사람에게 전달하거나 내용을 복사/인쇄할 수 없습니다.

### 15.3 옵션 B: Sensitivity Labels (민감도 레이블)

**누가 설정하나요?** — IT 관리자가 레이블을 만들고, 사용자가 메일 작성 시 선택합니다.

**사전 요구 사항**: Microsoft 365 E3/E5 또는 Business Premium 라이선스

**IT 관리자 설정 절차**:

1. [compliance.microsoft.com](https://compliance.microsoft.com) (또는 [purview.microsoft.com](https://purview.microsoft.com)) 접속
2. **"정보 보호"** → **"레이블"** → **"+ 레이블 만들기"**
3. 레이블 이름: `기밀 - 수신자 전용`
4. 범위: "항목(이메일, 파일)" 선택
5. 보호 설정:
   - ✅ 암호화 적용
   - ✅ 전달 금지
   - ✅ 인쇄 금지
   - ✅ 콘텐츠 표시 (머리글/바닥글에 "기밀" 표시)
6. 레이블을 게시(Publish)하여 사용자가 선택할 수 있도록 합니다.

**사용자 사용 방법**: Outlook 메일 작성 → 상단 "민감도" 버튼 → "기밀 - 수신자 전용" 선택

### 15.4 옵션 C: Exchange 전송 규칙 (자동 암호화)

**누가 설정하나요?** — IT 관리자

**어떻게 하나요?**

1. [admin.exchange.microsoft.com](https://admin.exchange.microsoft.com) 접속
2. **"메일 흐름"** → **"규칙"** → **"+ 규칙 추가"**
3. 조건: 메일 제목 또는 본문에 "대외비" 또는 "Confidential" 포함
4. 동작: "메시지 보안 적용" → "Office 365 메시지 암호화 적용"
5. 저장

**결과**: 직원이 메일 제목에 "대외비"를 넣으면 자동으로 암호화됩니다. 수신자는 OME 포털을 통해 메일을 열람합니다.

### 15.5 보안 옵션 비교 요약

| 옵션 | 설정 주체 | 라이선스 요구 | 보호 수준 | 사용 편의성 |
|------|-----------|-------------|-----------|------------|
| Do Not Forward | 발신자 (수동) | 모든 M365 | 전달/복사/인쇄 차단 | 매우 간편 |
| Sensitivity Labels | 관리자(생성) + 발신자(선택) | E3/E5/Business Premium | 암호화 + 전달/인쇄 차단 + 표시 | 간편 |
| Transport Rules | 관리자 (자동) | 모든 M365 | 자동 암호화 | 사용자 조작 불필요 |
| DLP | 관리자 (자동) | E3/E5/Business Premium | 민감 정보 외부 전송 차단 | 사용자 조작 불필요 |

---

## 16. 운영 가이드

### 16.1 일상 운영

SecureFileLink는 대부분 자동으로 운영됩니다. 관리자가 주기적으로 확인해야 하는 항목은 아래와 같습니다.

**매일 확인 (1분 소요)**

관리자 포털 대시보드에 접속하여 오늘의 전송/다운로드 건수가 정상 범위인지 확인합니다. 실패(failed) 상태 로그가 있으면 원인을 확인합니다.

**매주 확인 (5분 소요)**

다운로드 로그에서 의심스러운 활동(비정상적으로 많은 다운로드, 예상치 못한 국가에서의 접근 등)이 없는지 확인합니다. SharePoint 사이트의 저장 용량을 확인합니다.

**매월 확인 (10분 소요)**

Cloudflare 대시보드에서 Workers 요청 수를 확인합니다 (무료 한계 100,000건/일 대비). SystemConfig 설정값이 적절한지 검토합니다. 관리자 목록을 검토하여 퇴사자 등을 제거합니다.

### 16.2 용량 관리

| 저장소 | 확인 방법 | 대응 |
|--------|-----------|------|
| SharePoint 사이트 (25 TB) | SharePoint 관리 센터 → 활성 사이트 → 저장소 사용량 | 사이트당 25 TB 이므로 대부분 문제 없음. 부족 시 오래된 _mail 파일 삭제 |
| 개인 OneDrive (1~5 TB) | 각 사용자의 OneDrive 설정에서 확인 | 사용자 본인이 관리. 관리자 개입 불필요 |
| Cloudflare KV (1 GB) | Cloudflare 대시보드 → Workers → KV | 캐시 데이터만 저장하므로 대부분 문제 없음 |

### 16.3 클라이언트 시크릿 갱신

5장에서 만든 클라이언트 시크릿은 24개월 후 만료됩니다. 만료 1개월 전에 아래 절차로 갱신하세요.

1. [entra.microsoft.com](https://entra.microsoft.com) → 앱 등록 → SecureFileLink → 인증서 및 암호
2. "+ 새 클라이언트 암호" → 새로운 시크릿 생성
3. Cloudflare Worker의 환경 변수 업데이트:
   ```bash
   wrangler secret put CLIENT_SECRET
   # 새 시크릿 값 입력
   ```
4. 이전 시크릿 삭제

### 16.4 백업

SharePoint Lists와 문서 라이브러리의 데이터는 Microsoft 365 자체 백업 정책(93일간 삭제 항목 보존)에 따라 보호됩니다. 추가 백업이 필요한 경우 관리자 포털의 로그 CSV 내보내기 기능을 활용하세요.

---

## 17. 문제 해결 (Troubleshooting)

자주 발생할 수 있는 문제와 해결 방법을 정리합니다.

### 17.1 다운로드 관련

**증상: 수신자가 링크를 클릭하면 "잘못된 링크" (403) 오류가 표시됨**

- 원인: 토큰이 변조되었거나 TOKEN_SECRET이 불일치
- 해결: Worker의 TOKEN_SECRET 환경 변수가 토큰 생성 시 사용한 값과 동일한지 확인. 재배포 후 새 토큰으로 테스트

**증상: "만료된 링크" (410) 오류가 표시됨**

- 원인: 토큰의 만료 시각(`exp`)이 지남
- 해결: 발신자에게 파일을 다시 보내달라고 요청. 또는 관리자가 SystemConfig에서 `DefaultExpireDays_Mail` 값을 늘림

**증상: 다운로드 버튼을 클릭해도 파일이 받아지지 않음**

- 원인 1: `@microsoft.graph.downloadUrl`이 비어 있음 (Sensitivity Labels 적용 파일)
- 해결 1: SystemConfig에서 `FallbackStreaming = true` 확인. 파일 크기가 5 GB 이하인지 확인
- 원인 2: 서비스 주체가 Conditional Access에 의해 차단됨
- 해결 2: Entra ID → 조건부 액세스 → 정책 → SecureFileLink 서비스 주체를 예외로 추가

**증상: 다운로드가 중간에 끊김**

- 원인: 대용량 파일(수 GB)에서 네트워크 불안정
- 해결: 수신자에게 유선 네트워크 또는 안정적인 Wi‑Fi에서 다시 시도하도록 안내. downloadUrl은 클릭마다 새로 발급되므로 다시 링크를 클릭하면 됨

**증상: "파일이 삭제되었습니다" 페이지가 표시됨**

- 원인: 원본 파일이 SharePoint/OneDrive에서 삭제됨
- 해결: 발신자에게 파일을 다시 업로드하고 새 링크를 보내달라고 요청

### 17.2 업로드 관련

**증상: 파일 업로드가 실패하고 "권한 없음" 오류가 표시됨**

- 원인: Delegated 토큰의 권한이 부족하거나 만료됨
- 해결: Add‑in을 닫았다 다시 열어서 SSO 재인증. 문제 지속 시 5장의 API 권한과 관리자 동의를 다시 확인

**증상: 업로드가 느리거나 중간에 멈춤**

- 원인: 네트워크 속도 또는 청크 크기 문제
- 해결: SystemConfig에서 `ChunkSizeMB`를 5로 줄이거나 `AdaptiveChunk`가 true인지 확인

**증상: "파일 크기 초과" 오류**

- 원인: 파일 크기가 SystemConfig의 제한값을 초과
- 해결: `MaxFileSize_Shared` (기본 10 GB) 또는 `MaxFileSize_Mail` (기본 2 GB) 값 확인. 필요 시 관리자가 값을 늘림 (최대 250 GB까지 가능하나, 실용적으로 10 GB 권장)

### 17.3 Add‑in 관련

**증상: Outlook에 SecureFileLink 버튼이 보이지 않음**

- 원인 1: Add‑in이 배포되지 않았거나 배포 대상에 포함되지 않음
- 해결 1: admin.microsoft.com → 설정 → 통합 앱에서 배포 상태 확인
- 원인 2: Outlook 캐시 문제
- 해결 2: Outlook을 완전히 종료 후 재시작. Outlook on the Web에서도 보이지 않으면 배포 문제

**증상: "로그인 실패" 오류가 Add‑in에서 표시됨**

- 원인: MSAL.js 설정의 CLIENT_ID 또는 TENANT_ID가 잘못됨
- 해결: Add‑in 코드의 msalConfig에서 값을 확인하고, Entra ID 앱 등록의 리디렉션 URI가 정확한지 확인

### 17.4 관리자 포털 관련

**증상: 관리자 포털에 접속하면 Cloudflare Access 로그인 후 빈 화면**

- 원인: 포털이 Worker API와 통신하지 못함
- 해결: 브라우저 개발자 도구(F12) → Console 탭에서 오류 메시지 확인. Worker URL이 올바른지 확인. CORS 설정이 관리자 포털 도메인을 허용하는지 확인

**증상: 설정 변경이 저장되지 않음**

- 원인: Worker의 SharePoint Lists 쓰기 권한 문제
- 해결: Entra ID에서 `Sites.ReadWrite.All` Application 권한이 부여되고 관리자 동의가 완료되었는지 확인

### 17.5 Cloudflare 관련

**증상: 하루에 다운로드 수가 많아지면 오류 발생**

- 원인: 무료 플랜 100,000건/일 한계 초과
- 해결: Cloudflare 대시보드에서 Workers 사용량 확인. 한계 초과 시 $5/월 유료 플랜으로 업그레이드

**증상: Cron 작업이 실행되지 않음**

- 원인: wrangler.toml의 crons 설정이 잘못되었거나 Worker가 배포되지 않음
- 해결: Cloudflare 대시보드 → Workers → securefilelink-worker → Triggers 탭에서 Cron이 등록되어 있는지 확인. `wrangler deploy`로 재배포

---

## 18. 참고 자료 — 기술 사양 요약

### 18.1 SharePoint Online 서비스 한계

| 항목 | 제한값 |
|------|--------|
| 파일 업로드 최대 크기 | 250 GB |
| 사이트 최대 저장소 | 25 TB |
| 조직 전체 저장소 | 1 TB + 10 GB × 라이선스 수 |
| 리스트 항목 최대 수 | 30,000,000개 |
| 리스트 뷰 임계값 | 5,000개 (인덱싱 필요) |

### 18.2 OneDrive for Business 한계

| 항목 | 제한값 |
|------|--------|
| 사용자당 저장소 | 1 TB (E3/E5는 최대 5 TB) |
| 파일 업로드 최대 크기 | 250 GB |

### 18.3 Cloudflare Workers 무료 플랜 한계

| 항목 | 제한값 |
|------|--------|
| 일일 요청 수 | 100,000건 |
| 요청당 CPU 시간 | 10 ms |
| KV 읽기 | 100,000회/일 |
| KV 쓰기 | 1,000회/일 |
| KV 총 저장소 | 1 GB |
| KV 값 최대 크기 | 25 MB |
| Cron 트리거 | 3개 |
| Worker 크기 | 최대 10 MB |

### 18.4 비용

| 항목 | 비용 |
|------|------|
| Microsoft 365 | 기존 라이선스 (추가 비용 없음) |
| Cloudflare Workers 무료 | $0/월 |
| Cloudflare Workers 유료 | $5/월 (100k 요청/일 초과 시) |
| Cloudflare Pages | $0/월 |
| Cloudflare Access (50명 이하) | $0/월 |
| **총 월 비용** | **$0** |

### 18.5 Graph API 권한 요약

| 권한 | 유형 | 용도 |
|------|------|------|
| `Files.Read` | Delegated | 사용자 OneDrive 파일 목록 조회 |
| `Files.ReadWrite.All` | Delegated | SharePoint 파일 업로드/수정 |
| `Files.Read.All` | Application | 다운로드 URL 획득 |
| `Sites.ReadWrite.All` | Application | SharePoint Lists CRUD |
| `Mail.Send` | Delegated | 메일 발송 |
| `User.Read` | Delegated | 사용자 프로필 조회 |
| `User.Read.All` | Application | 사용자 조회 (관리자) |

### 18.6 토큰 JSON 구조

```json
{
  "tid": "전송기록 ID",
  "fid": "파일 ID",
  "fn": "파일명",
  "fs": 파일크기(바이트),
  "fr": "발신자 이메일",
  "src": "shared|personal|mail",
  "exp": 만료시각(Unix),
  "iat": 생성시각(Unix),
  "mc": 최대다운로드횟수,
  "sig": "HMAC-SHA256 서명"
}
```

### 18.7 API 라우트 전체 목록 (16개)

| # | 메서드 | 경로 | 설명 |
|---|--------|------|------|
| 1 | GET | `/d/{token}` | 다운로드 페이지 |
| 2 | POST | `/api/track/{trackId}` | 다운로드 완료 비콘 |
| 3 | GET | `/api/shared/files` | 공동 폴더 목록 |
| 4 | POST | `/api/shared/files` | 공동 폴더 업로드 |
| 5 | DELETE | `/api/shared/files/{fileId}` | 공동 폴더 삭제 |
| 6 | PUT | `/api/shared/files/{fileId}` | 공동 폴더 수정 |
| 7 | POST | `/api/mail/upload` | 메일 업로드 |
| 8 | POST | `/api/token/generate` | 토큰 생성 |
| 9 | GET | `/api/transfers` | 전송 기록 목록 |
| 10 | GET | `/api/transfers/{tid}` | 전송 기록 상세 |
| 11 | GET | `/api/admin/dashboard` | 대시보드 |
| 12 | GET | `/api/admin/config` | 설정 조회 |
| 13 | PUT | `/api/admin/config` | 설정 수정 |
| 14 | GET/POST/DELETE | `/api/admin/admins` | 관리자 관리 |
| 15 | GET | `/api/admin/logs` | 로그 검색 |
| 16 | GET | `/api/admin/logs/export` | 로그 CSV 내보내기 |

### 18.8 SystemConfig 기본값

| ConfigKey | 기본값 | 설명 |
|-----------|--------|------|
| `MaxFileSize_Shared` | `10737418240` | 공동 폴더 최대 10 GB |
| `MaxFileSize_Mail` | `2147483648` | 메일 업로드 최대 2 GB |
| `DefaultExpireDays_Mail` | `30` | 기본 만료 30일 |
| `ExpireOptions` | `[7, 30, 60]` | 만료 선택지 |
| `MaxDownloadsPerLink` | `0` | 무제한 |
| `ChunkSizeMB` | `10` | 업로드 청크 10 MB |
| `AdaptiveChunk` | `true` | 청크 자동 조절 |
| `FallbackStreaming` | `true` | 스트리밍 fallback |
| `StreamingMaxSize` | `5368709120` | 스트리밍 최대 5 GB |
| `SiteName` | `SecureFileLink` | 사이트 이름 |
| `WorkerBaseUrl` | `https://files.company.com` | Worker URL |

### 18.9 개발 로드맵

| Phase | 기간 | 내용 |
|-------|------|------|
| 1. 인프라 및 인증 | 4주 | Entra ID 앱 등록, SharePoint 사이트/리스트 생성, Cloudflare 설정, MSAL SSO |
| 2. 핵심 기능 | 6주 | 업로드, 다운로드, 토큰, Add‑in UI, 메일 발송 |
| 3. 관리자 포털 | 4주 | 포털 UI, Access 설정, Cron 자동화, 보안 가이드 |
| 4. 테스트 및 출시 | 4주 | 테스트, 파일럿, 피드백 반영, 전사 배포 |
| **합계** | **18주** | |

### 18.10 메모한 값 총정리

구축 과정에서 메모한 모든 값의 목록입니다.

```
TENANT_ID          = (5장에서 획득)
CLIENT_ID          = (5장에서 획득)
CLIENT_SECRET      = (5장에서 획득)
SHAREPOINT_SITE_URL = (6장에서 획득)
SHAREPOINT_SITE_ID  = (6장에서 획득)
DRIVE_ID           = (7장에서 획득)
TOKEN_SECRET       = (9장에서 직접 생성)
KV_NAMESPACE_ID    = (9장에서 획득)
WORKER_URL         = (10장에서 획득)
ADMIN_PORTAL_URL   = (12장에서 획득)
```

---

## 19. 버전 이력

| 버전 | 날짜 | 주요 변경 |
|------|------|-----------|
| v1.0 | 2026-08-01 | 초기 설계 (수신자별 토큰, 이메일 검증, OTP 포함) |
| v2.0 | 2026-08-15 | 수신자 인증 제거, 기본 무인증 다운로드, 보안은 M365 위임 |
| v2.1 | 2026-08-20 | 토큰 단순화 (13→9 필드), SystemConfig 축소, 로그 컬럼 축소 |
| v2.2 | 2026-08-25 | M365 보안 인프라 100% 위임 확정, KV 쓰기 96% 절감 |
| v2.3 | 2026-09-01 | 개인 폴더를 OneDrive로 이전, PersonalFolders 리스트 삭제 (7→6개), API 라우트 축소 (20→16개), Cron 3개로 확정, Unified Manifest 채택, downloadUrl 실시간 발급 흐름 확정, 초보자용 구축 가이드 형식으로 재작성 |
