# SecureFileLink 제품 계획서

| 항목 | 내용 |
|------|------|
| **버전** | v2.3 (Final) |
| **작성일** | 2026-09-01 |
| **상태** | 설계 확정 |

---

## 목차

1. [제품 비전](#1-제품-비전)
2. [핵심 가치](#2-핵심-가치)
3. [대상 사용자](#3-대상-사용자)
4. [시스템 아키텍처](#4-시스템-아키텍처)
5. [인증 및 권한](#5-인증-및-권한)
6. [저장소 구조](#6-저장소-구조)
7. [다운로드 토큰](#7-다운로드-토큰)
8. [보안 전략: M365 위임 모델](#8-보안-전략-m365-위임-모델)
9. [API 라우트](#9-api-라우트)
10. [Cloudflare Worker 핵심 로직](#10-cloudflare-worker-핵심-로직)
11. [SharePoint Lists 스키마](#11-sharepoint-lists-스키마)
12. [SystemConfig 설정값](#12-systemconfig-설정값)
13. [Cron 트리거 (자동화)](#13-cron-트리거-자동화)
14. [용량·성능·한계](#14-용량성능한계)
15. [Outlook Add‑in UI](#15-outlook-add-in-ui)
16. [관리자 포털](#16-관리자-포털)
17. [UX 시나리오](#17-ux-시나리오)
18. [개발 로드맵](#18-개발-로드맵)
19. [비용 분석](#19-비용-분석)
20. [제약 사항 및 고려 사항](#20-제약-사항-및-고려-사항)
21. [버전 이력](#21-버전-이력)

---

## 1. 제품 비전

SecureFileLink는 Outlook Add‑in 기반의 기업용 대용량 파일 전송 솔루션이다. 사용자가 이메일에 직접 첨부할 수 없는 대용량 파일을 SharePoint Online 또는 OneDrive for Business에 저장하고, 보안 다운로드 링크만 메일 본문에 삽입하여 전송한다. 별도의 서버 인프라 없이 Cloudflare Workers(서버리스)와 Microsoft 365만으로 운영하며, 보안은 기업이 이미 보유한 M365 보안 인프라(IRM, Sensitivity Labels, DLP, Defender, Conditional Access)에 100% 위임한다.

---

## 2. 핵심 가치

**데이터 주권 보장** — 모든 파일은 기업의 M365 테넌트(SharePoint/OneDrive) 내에 저장되며, 외부 서버로 복사되지 않는다. Microsoft의 저장 시 암호화(at‑rest encryption)가 기본 적용되고, IRM, Sensitivity Labels, DLP 등 기업이 이미 구성한 보안 정책이 그대로 파일에 적용된다.

**완전한 감사 추적** — 파일 업로드, 링크 생성, 다운로드 시도(성공/실패/만료), 관리자 조작 등 모든 이벤트가 SharePoint List에 기록되어 감사 및 컴플라이언스 요구사항을 충족한다.

**3가지 저장소 분리** — 공동 폴더(_shared, SharePoint), 개인 폴더(사용자 OneDrive), 메일 업로드(_mail, SharePoint)로 용도별 저장소를 분리하여 관리 복잡도를 최소화한다.

**제로 인프라** — SharePoint Lists 6개, Document Library 1개, Cloudflare Workers/KV/Pages로 구성되며, 별도의 VM, DB, 스토리지 서버가 불필요하다.

**보안과 편의성의 분리** — SecureFileLink는 편의 기능(대용량 전송, 링크 생성, 다운로드 로그, 자동 삭제, 할당량 관리)만 담당하고, 모든 보안 기능(수신자 인증, 전달 방지, 암호화, 분류)은 Microsoft 365에 위임한다.

**외부 수신자 무인증 다운로드** — 기본 설정에서 외부 수신자는 메일 내 링크를 클릭하면 즉시 파일을 다운로드할 수 있다. 별도의 로그인, 이메일 입력, OTP 절차가 없어 최대한의 사용 편의성을 제공한다.

---

## 3. 대상 사용자

기업 전 직원이 대상이며, 특히 외부 파일 송수신이 잦은 영업, 마케팅, 설계, 엔지니어링 부서가 주요 사용자이다. 중소·중견 기업(50~5,000명 규모)에 최적화되어 있으며, Microsoft 365 Business 또는 Enterprise 라이선스를 보유한 조직을 전제로 한다.

---

## 4. 시스템 아키텍처

### 4.1 전체 구성도

```
┌──────────────────────────────────────────────────────────────┐
│                      Outlook Client                          │
│  ┌────────────────────────────────────────────────────────┐  │
│  │  SecureFileLink Add‑in (React + TypeScript)            │  │
│  │  ┌──────────┐  ┌───────────┐  ┌─────────────────────┐ │  │
│  │  │ MSAL.js  │  │ Graph SDK │  │ Upload/File Picker  │ │  │
│  │  │  (SSO)   │  │(Delegated)│  │                     │ │  │
│  │  └────┬─────┘  └─────┬─────┘  └──────────┬──────────┘ │  │
│  └───────┼───────────────┼───────────────────┼────────────┘  │
└──────────┼───────────────┼───────────────────┼───────────────┘
           │               │                   │
           ▼               ▼                   ▼
┌──────────────────────────────────────────────────────────────┐
│                   Microsoft Graph API                        │
│  ┌───────────┐  ┌──────────────┐  ┌────────────────────────┐│
│  │ Entra ID  │  │  SharePoint  │  │ OneDrive for Business  ││
│  │  (Auth)   │  │   Online     │  │   (개인 폴더)           ││
│  └───────────┘  └──────┬───────┘  └───────────┬────────────┘│
└─────────────────────── ┼──────────────────────┼─────────────┘
                         │                      │
                         ▼                      ▼
┌──────────────────────────────────────────────────────────────┐
│               SharePoint Site: SecureFileLink                │
│  ┌────────────────────────────────────────────────────────┐  │
│  │ Document Library                                       │  │
│  │  ├── _shared/            (공동 폴더)                    │  │
│  │  ├── _mail/YYYY/MM/DD/   (메일 업로드, 자동 삭제)        │  │
│  │  └── _system/            (시스템 파일)                   │  │
│  ├────────────────────────────────────────────────────────┤  │
│  │ SharePoint Lists (×6)                                  │  │
│  │  ├── TransferRecords    (전송 기록)                     │  │
│  │  ├── DownloadLogs       (다운로드 로그)                  │  │
│  │  ├── SharedAssets        (공동 폴더 메타데이터)           │  │
│  │  ├── SystemConfig        (시스템 설정)                   │  │
│  │  ├── AdminUsers          (관리자 목록)                   │  │
│  │  └── AuditLog            (감사 로그)                    │  │
│  └────────────────────────────────────────────────────────┘  │
└──────────────────────────────────────────────────────────────┘
           │
           ▼
┌──────────────────────────────────────────────────────────────┐
│                    Cloudflare (서버리스)                      │
│  ┌─────────────┐  ┌──────────────┐  ┌──────────────────────┐│
│  │   Workers   │  │  Workers KV  │  │  Pages + Access      ││
│  │ (토큰 검증,  │  │ (캐시,        │  │  (관리자 포털)        ││
│  │  다운로드    │  │  rate limit) │  │                      ││
│  │  프록시)     │  │              │  │                      ││
│  └─────────────┘  └──────────────┘  └──────────────────────┘│
└──────────────────────────────────────────────────────────────┘
```

### 4.2 핵심 데이터 흐름

**업로드 및 링크 생성 (발신)**

사용자가 Outlook Add‑in에서 파일을 선택하면, MSAL.js가 Entra ID로부터 Delegated 토큰을 발급받고, Graph API `createUploadSession`을 통해 SharePoint(공동/_mail) 또는 OneDrive(개인)에 청크 업로드를 수행한다. 업로드 완료 후 Worker API(`POST /api/token/generate`)를 호출하여 HMAC‑SHA256 서명된 다운로드 토큰을 생성하고, 토큰이 포함된 다운로드 링크(`https://files.company.com/d/{token}`)를 메일 본문에 삽입한다.

**다운로드 (수신)**

외부 수신자가 메일 내 링크를 클릭하면 Cloudflare Worker가 요청을 받아 토큰의 HMAC 서명을 검증하고, 만료 여부와 다운로드 횟수 제한을 확인한 뒤, App‑only 토큰으로 Graph API를 호출하여 `@microsoft.graph.downloadUrl`(사전 인증된 URL, 약 1시간 유효)을 획득한다. Worker는 이 URL로 수신자 브라우저를 리다이렉트하거나 HTML 자동 다운로드 페이지를 반환한다. 이 사전 인증 URL은 클릭 시점에 실시간으로 생성되므로, 메일 발송 후 며칠이 지나도 토큰이 유효한 한 다운로드가 가능하다. 수신자는 어떠한 인증이나 로그인도 필요하지 않다.

---

## 5. 인증 및 권한

### 5.1 내부 사용자 (Outlook Add‑in)

내부 사용자는 Entra ID SSO를 통해 인증된다. MSAL.js가 Outlook Add‑in 내에서 Silent SSO를 수행하며, Delegated 토큰을 발급받는다. 이 토큰은 파일 업로드(SharePoint/OneDrive), 메일 발송, 공동 폴더 파일 목록 조회 등에 사용된다.

### 5.2 외부 수신자 (다운로드)

외부 수신자는 별도의 인증 절차 없이 HMAC‑SHA256 토큰 검증만으로 다운로드한다. 토큰에는 파일 ID, 발신자, 만료 시각 등이 포함되며, Worker가 서명을 검증하여 변조를 방지한다. 추가 보안이 필요한 경우 M365의 IRM, Sensitivity Labels, OME를 이메일에 적용하여 수신자 제한 및 전달 방지를 구현한다.

### 5.3 관리자 (Admin Portal)

관리자 포털은 Cloudflare Pages에 호스팅되며, Cloudflare Access를 통해 접근이 제어된다. AdminUsers SharePoint List에 등록된 사용자만 접근 가능하다.

### 5.4 Microsoft Graph API 권한

| 권한 | 유형 | 용도 |
|------|------|------|
| `Files.Read` | Delegated | 사용자 OneDrive 파일 목록 조회 (개인 폴더) |
| `Files.ReadWrite.All` | Delegated | SharePoint 파일 업로드/수정 |
| `Files.Read.All` | Application | 다운로드 URL 획득 (`@microsoft.graph.downloadUrl`) |
| `Sites.ReadWrite.All` | Application | SharePoint Lists CRUD, 문서 라이브러리 관리 |
| `Mail.Send` | Delegated | 사용자 대신 메일 발송 |
| `User.Read` | Delegated | 현재 사용자 프로필 조회 |
| `User.Read.All` | Application | 사용자 조회 (관리자 기능) |

---

## 6. 저장소 구조

### 6.1 3대 저장소

| 저장소 | 위치 | 최대 파일 크기 | 용량 | 자동 삭제 | SecureFileLink 역할 | 관리 주체 |
|--------|------|---------------|------|-----------|-------------------|-----------|
| 공동 폴더 (_shared) | SharePoint Document Library | 10 GB (설정 가능, 최대 250 GB) | 사이트 25 TB 내 공유 | 무기한 (관리자 수동 삭제) | 업로드, 링크 생성, 추적, 삭제 | 관리자 |
| 개인 폴더 | 각 사용자 OneDrive `/SecureFileLink/` | OneDrive 정책 따름 | 사용자당 1 TB~5 TB (M365 라이선스별) | 사용자 자율 관리 | 파일 목록 조회(읽기 전용), 링크 생성, 추적 | 사용자 본인 |
| 메일 업로드 (_mail) | SharePoint Document Library | 2 GB (설정 가능, 최대 10 GB) | 사이트 25 TB 내 공유 | 7/30/60일 선택 (Cron 자동 삭제) | 업로드, 링크 생성, 추적, 자동 삭제 | 시스템 자동 |

### 6.2 폴더 구조

```
SharePoint Site: SecureFileLink
├── SecureFileLink/
│   ├── _shared/                          ← 공동 폴더
│   │   ├── 카탈로그_2026.pdf
│   │   ├── 제품소개서_v3.pptx
│   │   └── ...
│   ├── _mail/                            ← 메일 업로드 (자동 삭제)
│   │   └── 2026/
│   │       └── 09/
│   │           └── 01/
│   │               ├── {uuid}_설계도면.dwg
│   │               └── {uuid}_계약서.pdf
│   └── _system/                          ← 시스템 파일
│       └── (향후 템플릿 등)

사용자 OneDrive (각 사용자별)
└── SecureFileLink/                       ← 개인 폴더
    ├── 고객제안서_A사.pdf
    ├── 내부검토자료.xlsx
    └── ...
```

### 6.3 개인 폴더 운영 원칙

개인 OneDrive에 대해 SecureFileLink는 읽기(파일 목록 조회) 및 링크 생성만 수행한다. 업로드, 삭제, 용량 관리 등은 사용자 본인이 OneDrive에서 직접 관리하며, 관리자는 앱을 통해 개인 OneDrive 파일에 접근하지 않는다. 이는 개인 영역에 대한 프라이버시를 보장하고, 관리자의 부적절한 권한 행사 가능성을 원천 차단하기 위함이다.

---

## 7. 다운로드 토큰

### 7.1 토큰 구조 (JSON)

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
  "sig": "HMAC-SHA256-signature"
}
```

| 필드 | 타입 | 설명 |
|------|------|------|
| `tid` | string | TransferRecord ID (SharePoint List 항목 ID) |
| `fid` | string | SharePoint/OneDrive driveItem ID |
| `fn` | string | 파일명 (다운로드 페이지 표시용) |
| `fs` | number | 파일 크기 (바이트) |
| `fr` | string | 발신자 이메일 |
| `src` | string | 저장소 유형: `shared`, `personal`, `mail` |
| `exp` | number | 만료 시각 (Unix timestamp, UTC) |
| `iat` | number | 발급 시각 (Unix timestamp, UTC) |
| `mc` | number | 최대 다운로드 횟수 (0 = 무제한) |
| `sig` | string | HMAC‑SHA256 서명 (`sig` 필드 제외한 JSON에 대해 계산) |

### 7.2 토큰 생성 및 검증

토큰은 `sig` 필드를 제외한 JSON 문자열을 HMAC‑SHA256으로 서명하여 생성한다. 서명 키는 Cloudflare Workers의 환경 변수(`TOKEN_SECRET`)에 저장된다. 생성된 JSON 전체를 Base64URL로 인코딩하여 다운로드 URL에 포함한다.

다운로드 URL 형식: `https://files.company.com/d/{base64url-encoded-token}`

검증 시 Worker는 토큰을 디코딩한 뒤 `sig` 필드를 분리하고, 나머지 JSON에 대해 동일한 HMAC‑SHA256을 계산하여 일치 여부를 확인한다. 서명이 유효하면 `exp` 만료 시각과 `mc` 다운로드 횟수를 추가로 검증한다.

### 7.3 downloadUrl의 이중 구조

메일에 포함되는 URL과 실제 파일 다운로드 URL은 별개이다.

**메일 내 URL (SecureFileLink 토큰 URL)** — `https://files.company.com/d/{token}` 형태이며, 만료 기간은 시스템 설정에 따라 7/30/60일 등으로 구성한다. 이 URL에는 파일의 실제 다운로드 경로가 포함되지 않는다.

**Microsoft downloadUrl** — 수신자가 링크를 클릭한 시점에 Worker가 Graph API를 호출하여 실시간으로 발급받는 사전 인증 URL이다. 약 1시간 동안 유효하며, Authorization 헤더 없이 접근 가능하다. 이 URL은 매 클릭마다 새로 발급되므로, 메일 발송 후 며칠이 지나도 토큰이 유효한 한 다운로드가 가능하다.

| 구분 | SecureFileLink 토큰 URL | Microsoft downloadUrl |
|------|------------------------|----------------------|
| 포함 위치 | 메일 본문 | Worker 내부에서만 사용 |
| 생성 시점 | 메일 발송 시 | 수신자 클릭 시 (실시간) |
| 유효 기간 | 7/30/60일 (설정 가능) | 약 1시간 |
| 인증 필요 | 불필요 (HMAC 검증) | 불필요 (사전 인증) |
| 파일 경로 노출 | 없음 (fid만 포함) | SharePoint 내부 URL |

---

## 8. 보안 전략: M365 위임 모델

SecureFileLink는 자체적으로 수신자 인증(이메일 입력, OTP, 디바이스 바인딩 등)을 구현하지 않는다. 모든 보안 기능은 기업이 이미 보유한 Microsoft 365 보안 인프라에 위임한다.

### 8.1 기본 동작 (편의성 우선)

기본 상태에서 메일 수신자는 링크 클릭 즉시 파일을 다운로드할 수 있다. 어떠한 인증 절차도 없으며, 이는 대부분의 일상적인 파일 전송(영업 자료, 카탈로그, 제안서 등)에 적합하다.

### 8.2 보안 강화 옵션 (M365 기능 활용)

추가 보호가 필요한 경우 발신자가 Outlook에서 메일 발송 시 다음 M365 보안 기능을 적용한다.

**전달 방지 (Do Not Forward)** — Outlook → 옵션 → 암호화 → 전달 금지를 선택하면, 수신자가 메일을 다른 사람에게 전달하거나 내용을 복사/인쇄할 수 없다. 메일 본문에 포함된 다운로드 링크도 함께 보호된다.

**Sensitivity Labels (민감도 레이블)** — Microsoft Purview에서 "기밀 – 수신자 전용" 등의 레이블을 구성하면, 메일 및 첨부 파일에 암호화, 전달 금지, 인쇄 금지 등의 정책이 자동 적용된다. E3/E5 또는 Microsoft 365 Business Premium 라이선스가 필요하다.

**Exchange Transport Rules (전송 규칙)** — 특정 키워드(예: "대외비", "Confidential")가 포함된 메일에 자동으로 암호화를 적용하는 Exchange 전송 규칙을 구성할 수 있다.

**DLP (데이터 손실 방지)** — Microsoft Purview DLP 정책을 통해 주민등록번호, 신용카드 번호 등 민감한 정보가 포함된 파일의 외부 전송을 자동으로 차단하거나 경고할 수 있다.

**Microsoft Defender for Office 365** — Safe Links가 메일 내 URL을 실시간으로 검사하고, Safe Attachments가 악성 파일을 차단한다.

**Conditional Access (조건부 액세스)** — Entra ID 조건부 액세스 정책으로 특정 조건(비관리 디바이스, 해외 IP 등)에서의 메일 접근을 제한할 수 있다.

### 8.3 SecureFileLink vs M365 역할 분담

| 기능 영역 | SecureFileLink 담당 | M365 담당 |
|-----------|-------------------|-----------|
| 대용량 파일 업로드 | ✅ | — |
| 다운로드 링크 생성 | ✅ | — |
| 다운로드 로그 기록 | ✅ | — |
| 만료/자동 삭제 | ✅ | — |
| 토큰 변조 방지 (HMAC) | ✅ | — |
| 수신자 신원 인증 | — | ✅ IRM / OME |
| 메일 전달 방지 | — | ✅ Do Not Forward |
| 파일 암호화 | — | ✅ Sensitivity Labels |
| 민감 정보 유출 방지 | — | ✅ DLP |
| 악성 URL/파일 차단 | — | ✅ Defender |
| 접근 조건 제어 | — | ✅ Conditional Access |

### 8.4 관리자 보안 설정 가이드

관리자 포털에 "보안 설정 가이드" 섹션을 배치하여, 각 M365 보안 기능의 설정 방법과 권장 구성을 안내한다.

| 보안 기능 | 설정 경로 | 권장 구성 |
|-----------|-----------|-----------|
| Do Not Forward | Outlook → 옵션 → 암호화 → 전달 금지 | 민감 메일에 수동 적용 |
| Sensitivity Labels | Microsoft Purview → Information Protection → Labels | "기밀" 레이블: 암호화 + 전달 금지 + 인쇄 금지 |
| DLP | Microsoft Purview → DLP → Policies | 주민번호/카드번호 포함 파일 외부 전송 차단 |
| Transport Rules | Exchange Admin → Mail Flow → Rules | "대외비" 키워드 포함 시 자동 암호화 |
| Defender Safe Links | Microsoft 365 Defender → Policies → Safe Links | 기본 활성화 유지 |

---

## 9. API 라우트

### 9.1 다운로드 (외부 수신자용)

| 메서드 | 경로 | 설명 |
|--------|------|------|
| GET | `/d/{token}` | 다운로드 페이지 (토큰 검증 → downloadUrl 획득 → 자동 다운로드) |
| POST | `/api/track/{trackId}` | 다운로드 완료 비콘 수신 (클라이언트 → Worker) |

### 9.2 공동 폴더 관리

| 메서드 | 경로 | 설명 |
|--------|------|------|
| GET | `/api/shared/files` | 공동 폴더 파일 목록 조회 |
| POST | `/api/shared/files` | 공동 폴더 파일 업로드 |
| DELETE | `/api/shared/files/{fileId}` | 공동 폴더 파일 삭제 |
| PUT | `/api/shared/files/{fileId}` | 공동 폴더 파일 메타데이터 수정 |

### 9.3 메일 업로드

| 메서드 | 경로 | 설명 |
|--------|------|------|
| POST | `/api/mail/upload` | 메일 첨부용 파일 업로드 (SharePoint `_mail/`) |

### 9.4 토큰 생성

| 메서드 | 경로 | 설명 |
|--------|------|------|
| POST | `/api/token/generate` | HMAC‑SHA256 다운로드 토큰 생성 |

### 9.5 전송 기록

| 메서드 | 경로 | 설명 |
|--------|------|------|
| GET | `/api/transfers` | 현재 사용자의 전송 기록 조회 |
| GET | `/api/transfers/{tid}` | 특정 전송 기록 상세 조회 |

### 9.6 관리자 전용

| 메서드 | 경로 | 설명 |
|--------|------|------|
| GET | `/api/admin/dashboard` | 대시보드 통계 (전송 건수, 용량, 활성 링크 등) |
| GET | `/api/admin/config` | SystemConfig 조회 |
| PUT | `/api/admin/config` | SystemConfig 수정 |
| GET | `/api/admin/admins` | 관리자 목록 조회 |
| POST | `/api/admin/admins` | 관리자 추가 |
| DELETE | `/api/admin/admins/{id}` | 관리자 삭제 |
| GET | `/api/admin/logs` | 다운로드 로그 검색 |
| GET | `/api/admin/logs/export` | 로그 CSV 내보내기 |

총 16개 엔드포인트로 구성된다.

---

## 10. Cloudflare Worker 핵심 로직

### 10.1 다운로드 처리 흐름

```
수신자 클릭 → GET /d/{token}
  │
  ├─ 1. Base64URL 디코딩 → JSON 파싱
  ├─ 2. HMAC‑SHA256 서명 검증 → 실패 시 403 반환
  ├─ 3. exp 만료 확인 → 만료 시 410 Gone 반환
  ├─ 4. mc > 0 이면 다운로드 횟수 확인 → 초과 시 429 반환
  ├─ 5. DownloadLogs에 'initiated' 상태 기록
  ├─ 6. App‑only 토큰으로 Graph API 호출:
  │     GET /drives/{driveId}/items/{fid}?select=id,@microsoft.graph.downloadUrl
  │     → downloadUrl 획득
  ├─ 7. downloadUrl이 비어 있으면 Fallback Streaming:
  │     GET /drives/{driveId}/items/{fid}/content
  │     → Worker가 스트림을 중계 (최대 5 GB)
  └─ 8. HTML 페이지 반환:
        - 파일명, 크기, 발신자 표시
        - JavaScript로 downloadUrl 자동 다운로드 시작
        - 다운로드 완료 시 비콘 전송 (POST /api/track/{trackId})
```

### 10.2 Fallback Streaming

`@microsoft.graph.downloadUrl`이 반환되지 않는 경우(Sensitivity Labels 적용 파일 등) Worker가 Graph API의 `/content` 엔드포인트를 호출하여 파일 스트림을 직접 중계한다. 이 경우 Worker 메모리를 경유하므로 최대 5 GB로 제한한다.

### 10.3 App‑only 토큰 캐시

Graph API 호출을 위한 App‑only 액세스 토큰은 Cloudflare Workers KV에 캐시한다. 토큰의 기본 유효 시간은 약 1시간이며, 만료 5분 전에 갱신한다. 이를 통해 매 요청마다 Entra ID에 토큰을 요청하는 오버헤드를 제거한다.

---

## 11. SharePoint Lists 스키마

### 11.1 TransferRecords (전송 기록)

| 컬럼 | 타입 | 설명 |
|------|------|------|
| TransferId | Text (PK) | 전송 기록 고유 ID |
| FileId | Text | driveItem ID |
| FileName | Text | 파일명 |
| FileSize | Number | 파일 크기 (바이트) |
| SenderEmail | Text | 발신자 이메일 |
| SenderName | Text | 발신자 이름 |
| Recipients | Text (Multi‑line) | 수신자 이메일 목록 (JSON 배열) |
| Source | Choice | shared / personal / mail |
| TokenExp | DateTime | 토큰 만료 시각 |
| MaxDownloads | Number | 최대 다운로드 횟수 (0 = 무제한) |
| CurrentDownloads | Number | 현재 다운로드 횟수 |
| Status | Choice | active / expired / deleted / file_deleted |
| CreatedAt | DateTime | 생성 시각 |
| DriveId | Text | 파일이 위치한 Drive ID |

### 11.2 DownloadLogs (다운로드 로그)

| 컬럼 | 타입 | 설명 |
|------|------|------|
| TrackId | Text (PK) | 로그 고유 ID |
| TransferRecordId | Text (FK) | TransferRecords 참조 |
| FileId | Text | driveItem ID |
| FileName | Text | 파일명 |
| FileSize | Number | 파일 크기 |
| SenderEmail | Text | 발신자 이메일 |
| DownloaderIP | Text | 다운로더 IP 주소 |
| Country | Text | GeoIP 국가 코드 |
| UserAgent | Text (Multi‑line) | 브라우저 User‑Agent |
| RequestTime | DateTime | 요청 시각 |
| Status | Choice | initiated / completed / failed / expired / limit_exceeded |
| DownloadMethod | Choice | redirect / streaming |
| CompletedTime | DateTime | 완료 시각 |
| ErrorDetail | Text | 오류 상세 (실패 시) |

### 11.3 SharedAssets (공동 폴더 메타데이터)

| 컬럼 | 타입 | 설명 |
|------|------|------|
| AssetId | Text (PK) | 고유 ID |
| FileId | Text | driveItem ID |
| FileName | Text | 파일명 |
| FileSize | Number | 파일 크기 |
| UploaderEmail | Text | 업로더 이메일 |
| UploaderName | Text | 업로더 이름 |
| Description | Text (Multi‑line) | 파일 설명 |
| Tags | Text | 태그 (콤마 구분) |
| DownloadCount | Number | 누적 다운로드 횟수 |
| CreatedAt | DateTime | 업로드 시각 |
| DriveId | Text | Drive ID |

### 11.4 SystemConfig (시스템 설정)

| 컬럼 | 타입 | 설명 |
|------|------|------|
| ConfigKey | Text (PK) | 설정 키 |
| ConfigValue | Text | 설정 값 |
| Description | Text | 설명 |
| UpdatedAt | DateTime | 최종 수정 시각 |
| UpdatedBy | Text | 수정자 |

### 11.5 AdminUsers (관리자 목록)

| 컬럼 | 타입 | 설명 |
|------|------|------|
| AdminId | Text (PK) | 고유 ID |
| Email | Text | 관리자 이메일 |
| DisplayName | Text | 표시 이름 |
| Role | Choice | super_admin / admin / viewer |
| AddedAt | DateTime | 등록 시각 |
| AddedBy | Text | 등록자 |

### 11.6 AuditLog (감사 로그)

| 컬럼 | 타입 | 설명 |
|------|------|------|
| AuditId | Text (PK) | 고유 ID |
| Action | Choice | config_update / admin_add / admin_remove / file_delete / manual_expire |
| ActorEmail | Text | 수행자 이메일 |
| TargetType | Text | 대상 유형 (config / admin / file / transfer) |
| TargetId | Text | 대상 ID |
| Details | Text (Multi‑line) | 변경 내용 (JSON) |
| Timestamp | DateTime | 수행 시각 |
| IP | Text | 수행자 IP |

---

## 12. SystemConfig 설정값

| ConfigKey | 기본값 | 설명 |
|-----------|--------|------|
| `MaxFileSize_Shared` | `10737418240` (10 GB) | 공동 폴더 최대 파일 크기 (바이트) |
| `MaxFileSize_Mail` | `2147483648` (2 GB) | 메일 업로드 최대 파일 크기 (바이트) |
| `DefaultExpireDays_Mail` | `30` | 메일 업로드 기본 만료일 |
| `ExpireOptions` | `[7, 30, 60]` | 만료 기간 선택지 (일) |
| `MaxDownloadsPerLink` | `0` | 링크당 최대 다운로드 횟수 (0 = 무제한) |
| `ChunkSizeMB` | `10` | 업로드 청크 크기 (MB, 320 KB의 배수) |
| `AdaptiveChunk` | `true` | 네트워크 상태에 따른 청크 크기 자동 조절 |
| `FallbackStreaming` | `true` | downloadUrl 미반환 시 스트리밍 fallback 활성화 |
| `StreamingMaxSize` | `5368709120` (5 GB) | 스트리밍 fallback 최대 파일 크기 (바이트) |
| `SiteName` | `SecureFileLink` | SharePoint 사이트 표시 이름 |
| `WorkerBaseUrl` | `https://files.company.com` | Worker 기본 URL |

---

## 13. Cron 트리거 (자동화)

Cloudflare Workers 무료 플랜은 최대 3개의 Cron 트리거를 지원한다.

### 13.1 매시간 실행 — 만료 파일 삭제

- 스케줄: `0 * * * *` (매시간 정각)
- 동작: TransferRecords에서 `Source = mail`이고 `TokenExp < now`인 항목을 조회하여 해당 SharePoint 파일을 삭제하고, Status를 `deleted`로 변경한다.
- 만료 3일 이내 항목에 대해 발신자에게 알림 메일을 발송한다.

### 13.2 매일 실행 — 정리 및 상태 동기화

- 스케줄: `0 2 * * *` (매일 02:00 UTC)
- 동작: `_mail/` 하위의 빈 날짜 폴더를 삭제한다. `Source = personal`인 TransferRecords에서 파일 존재 여부를 확인하고, 삭제된 파일의 Status를 `file_deleted`로 갱신한다.

### 13.3 매주 실행 — 통계 집계

- 스케줄: `0 3 * * 0` (매주 일요일 03:00 UTC)
- 동작: 주간 전송/다운로드 통계를 집계하여 대시보드 캐시(KV)를 갱신한다. SharedAssets의 DownloadCount를 DownloadLogs 기준으로 동기화한다. 90일 이상 경과한 로그를 아카이브 처리한다.

---

## 14. 용량·성능·한계

### 14.1 SharePoint Online 서비스 한계

| 항목 | 제한값 |
|------|--------|
| 파일 업로드 최대 크기 | 250 GB |
| 사이트(사이트 컬렉션) 최대 저장소 | 25 TB |
| 조직 전체 저장소 | 1 TB + 10 GB × 라이선스 수 |
| 리스트 항목 최대 수 | 30,000,000개 |
| 리스트 뷰 임계값 | 5,000개 (인덱싱 필요) |
| 고유 권한 최대 수 | 50,000개 (권장 5,000개 이하) |
| 사이트 컬렉션 최대 수 | 2,000,000개 |

### 14.2 OneDrive for Business 한계

| 항목 | 제한값 |
|------|--------|
| 사용자당 저장소 | 1 TB (E3/E5는 최대 5 TB) |
| 파일 업로드 최대 크기 | 250 GB |
| 동기화 권장 파일 수 | 300,000개 이하 |

### 14.3 Cloudflare Workers 무료 플랜 한계

| 항목 | 제한값 |
|------|--------|
| 일일 요청 수 | 100,000건 (UTC 00:00 초기화) |
| 요청당 CPU 시간 | 10 ms |
| Workers KV 읽기 | 100,000회/일 |
| Workers KV 쓰기 | 1,000회/일 |
| Workers KV 리스트 | 1,000회/일 |
| Workers KV 삭제 | 1,000회/일 |
| Workers KV 총 저장소 | 1 GB |
| KV 값 최대 크기 | 25 MB |
| Cron 트리거 | 3개 |
| Worker 크기 | 최대 10 MB |

### 14.4 성능 예상치

Worker의 토큰 검증 및 Graph API 호출은 CPU 약 2 ms + Graph I/O 200~500 ms로 처리된다. 파일 전송 속도는 수신자의 네트워크 환경에만 의존하며, Worker를 경유하지 않고 Microsoft CDN에서 직접 다운로드된다(Fallback Streaming 제외).

### 14.5 일일 처리 용량 추정

100명 사용자 기준으로 하루 평균 전송 50건, 다운로드 150건을 가정하면, Worker 요청은 약 500건/일(토큰 생성, 다운로드 페이지, 비콘, API 호출 등)으로 무료 플랜 한계의 0.5%에 해당한다. KV 쓰기는 약 34건/일(토큰 캐시 갱신, 로그 버퍼 등)로 무료 플랜 한계의 3.4%에 해당한다.

---

## 15. Outlook Add‑in UI

### 15.1 Add‑in 매니페스트

Unified App Manifest (v1.3+)를 사용하며, Outlook on the Web, Windows, Mac, iOS, Android를 지원한다.

### 15.2 주요 화면 구성

**메일 작성 패널 (Task Pane)**

Add‑in을 열면 3개 탭이 표시된다.

- **공동 폴더 탭**: SharePoint `_shared/` 폴더의 파일 목록을 표시한다. 파일명, 크기, 업로드 날짜, 다운로드 횟수가 표시되며, "링크 삽입" 버튼을 클릭하면 토큰을 생성하고 메일 본문에 다운로드 링크를 삽입한다.
- **내 OneDrive 탭**: 사용자 OneDrive의 `/SecureFileLink/` 폴더 파일 목록을 읽기 전용으로 표시한다. "선택 → 링크 삽입" 버튼으로 토큰 생성 및 링크 삽입이 가능하다. 업로드/삭제 버튼은 제공하지 않으며, "파일 관리는 OneDrive에서 직접 수행하세요"라는 안내 문구를 표시한다.
- **새 파일 업로드 탭**: 로컬 파일을 선택하여 공동 폴더 또는 메일 업로드(`_mail/`)에 업로드한다. 메일 업로드 시 만료 기간(7/30/60일)을 선택할 수 있다. 업로드 진행률 표시 및 청크 업로드를 지원한다.

**링크 삽입 확인 다이얼로그**

파일 선택 후 "링크 삽입" 클릭 시 확인 다이얼로그가 표시된다. 파일명, 크기, 만료 기간, 최대 다운로드 횟수(선택 사항)를 확인하고 "삽입" 버튼을 누르면 메일 본문에 다운로드 링크가 삽입된다.

### 15.3 메일 본문 삽입 형식

```html
<div style="border:1px solid #e0e0e0; border-radius:8px; padding:16px; margin:8px 0;">
  <p style="margin:0 0 8px; font-weight:bold;">📎 SecureFileLink</p>
  <p style="margin:0 0 4px;">파일명: 회사소개서_2026.pdf</p>
  <p style="margin:0 0 4px;">크기: 11.9 MB</p>
  <p style="margin:0 0 8px;">만료: 2026-10-01</p>
  <a href="https://files.company.com/d/eyJmaWQ..." 
     style="background:#0078d4; color:#fff; padding:10px 20px; 
            border-radius:4px; text-decoration:none; display:inline-block;">
    다운로드
  </a>
</div>
```

---

## 16. 관리자 포털

### 16.1 호스팅 및 접근 제어

Cloudflare Pages에 정적 SPA(React)를 배포하고, Cloudflare Access를 통해 접근을 제어한다. Access 정책은 Entra ID 연동(OIDC)을 사용하며, AdminUsers 리스트에 등록된 이메일만 허용한다.

### 16.2 주요 화면

**대시보드** — 오늘/이번 주/이번 달 전송 건수, 다운로드 건수, 활성 링크 수, 총 전송 용량, 저장소 사용량을 표시한다. 최근 다운로드 로그 10건을 실시간으로 표시한다.

**파일 관리** — 공동 폴더(`_shared/`) 파일 목록을 관리한다. 파일 추가, 삭제, 메타데이터(설명, 태그) 편집이 가능하다. 메일 업로드(`_mail/`) 파일은 조회 및 수동 삭제가 가능하다. 개인 OneDrive 파일은 표시하지 않는다.

**전송 기록** — 모든 사용자의 전송 기록을 조회한다. 발신자, 수신자, 파일명, 상태, 다운로드 횟수 등으로 필터링 및 검색이 가능하다.

**다운로드 로그** — 모든 다운로드 이벤트를 조회한다. IP, 국가, User‑Agent, 상태 등으로 필터링 및 검색이 가능하며, CSV로 내보내기를 지원한다.

**시스템 설정** — SystemConfig의 모든 항목을 GUI로 편집한다. 변경 시 AuditLog에 자동 기록된다.

**관리자 관리** — 관리자 추가/삭제/역할 변경을 수행한다.

**보안 설정 가이드** — M365 보안 기능(Do Not Forward, Sensitivity Labels, DLP, Transport Rules, Defender)의 설정 방법과 권장 구성을 안내하는 문서 페이지이다.

---

## 17. UX 시나리오

### 시나리오 1: 영업팀 — 카탈로그 전송

영업 담당자가 고객에게 제품 카탈로그(50 MB)를 전송한다. Add‑in에서 공동 폴더 탭을 열고, 이미 등록된 카탈로그 파일을 선택하여 "링크 삽입"을 클릭한다. 메일 본문에 다운로드 링크가 삽입되고, 메일을 보낸다. 고객은 메일 내 "다운로드" 버튼을 클릭하면 즉시 파일이 다운로드된다. 로그인이나 인증 절차 없이 원클릭으로 완료된다.

### 시나리오 2: 마케팅팀 — 다수 수신자 전송

마케팅 담당자가 10명의 파트너에게 신제품 자료(200 MB)를 전송한다. 새 파일 업로드 탭에서 파일을 메일 업로드로 업로드하고(만료 30일), 링크를 삽입한 뒤 10명을 CC로 메일을 보낸다. 모든 수신자가 동일한 링크로 다운로드 가능하며, 30일 후 파일이 자동 삭제된다.

### 시나리오 3: 설계팀 — 대용량 도면 전송

설계 엔지니어가 협력사에 CAD 도면(3 GB)을 전송한다. 새 파일 업로드 탭에서 메일 업로드를 선택하고, 청크 업로드로 파일을 업로드한다. 업로드 완료 후 링크를 삽입하여 메일을 보낸다. 협력사 담당자는 링크 클릭 즉시 3 GB 파일을 다운로드한다.

### 시나리오 4: 법무팀 — 기밀 계약서 전송 (M365 보안 적용)

법무 담당자가 계약 상대방에게 기밀 계약서(10 MB)를 전송한다. Add‑in에서 개인 OneDrive 탭을 열고, 미리 저장해 둔 계약서를 선택하여 링크를 삽입한다. Outlook에서 "옵션 → 암호화 → 전달 금지"를 선택하고 메일을 보낸다. 수신자는 메일을 열어 다운로드 링크를 클릭할 수 있지만, 메일을 다른 사람에게 전달하거나 내용을 복사/인쇄할 수 없다.

### 시나리오 5: 개인 OneDrive 파일 전송

담당자가 자신의 OneDrive에 저장해 둔 고객 제안서를 전송한다. Add‑in에서 "내 OneDrive" 탭을 열면 `/SecureFileLink/` 폴더의 파일 목록이 표시된다. 파일을 선택하고 "링크 삽입"을 클릭하면 토큰이 생성되어 메일 본문에 삽입된다. 파일 관리(업로드, 삭제, 이름 변경 등)가 필요하면 OneDrive 웹/앱에서 직접 수행한다.

---

## 18. 개발 로드맵

### Phase 1: 인프라 및 인증 (4주)

- SharePoint Site 생성 및 Document Library 구성
- SharePoint Lists 6개 생성 및 스키마 설정
- Entra ID 앱 등록 및 Graph API 권한 구성
- Cloudflare Workers 프로젝트 셋업 및 KV 네임스페이스 생성
- MSAL.js SSO 구현 및 토큰 발급 테스트
- App‑only 토큰 발급 및 Graph API 호출 테스트

### Phase 2: 핵심 기능 (6주)

- 청크 업로드 구현 (`createUploadSession`, 10 MB 청크)
- 공동 폴더 파일 관리 (목록 조회, 업로드, 삭제, 메타데이터 수정)
- 메일 업로드 구현 (`_mail/YYYY/MM/DD/{uuid}_파일명`)
- OneDrive 개인 폴더 파일 목록 조회 (읽기 전용)
- HMAC‑SHA256 토큰 생성 및 검증 Worker 구현
- 다운로드 페이지 구현 (토큰 검증 → downloadUrl → 자동 다운로드)
- Fallback Streaming 구현 (downloadUrl 미반환 시)
- 다운로드 완료 비콘 및 로그 기록
- Outlook Add‑in Task Pane UI (3개 탭)
- 메일 본문 링크 삽입 기능
- Graph API `sendMail` 연동

### Phase 3: 관리자 포털 및 자동화 (4주)

- Cloudflare Pages 관리자 포털 SPA 개발
- Cloudflare Access 설정 (Entra ID OIDC 연동)
- 대시보드, 파일 관리, 전송 기록, 로그 검색 화면
- SystemConfig GUI 편집 및 AuditLog 기록
- 관리자 관리 (추가/삭제/역할)
- Cron 트리거 3개 구현 (만료 삭제, 정리, 통계)
- 보안 설정 가이드 문서 페이지

### Phase 4: 테스트 및 출시 (4주)

- 단위 테스트 및 통합 테스트
- 대용량 파일 (1 GB, 5 GB, 10 GB) 업로드/다운로드 테스트
- 다양한 브라우저/디바이스에서 다운로드 테스트
- Sensitivity Labels 적용 파일의 Fallback Streaming 테스트
- 부하 테스트 (동시 다운로드 시나리오)
- 파일럿 배포 (특정 부서 대상)
- 피드백 반영 및 버그 수정
- 전사 배포 및 사용자 교육

총 개발 기간: 18주

---

## 19. 비용 분석

| 항목 | 비용 | 비고 |
|------|------|------|
| Microsoft 365 | 기존 라이선스 | 추가 비용 없음 |
| SharePoint Online | 기존 라이선스 포함 | 추가 비용 없음 |
| OneDrive for Business | 기존 라이선스 포함 | 추가 비용 없음 |
| Cloudflare Workers (무료) | $0/월 | 100k 요청/일 이내 |
| Cloudflare Workers (유료) | $5/월 | 100k 초과 시 |
| Cloudflare Pages | $0/월 | 무료 플랜 |
| Cloudflare Access | $0/월 | 50명 이하 무료 |
| **총 월 비용** | **$0** | 무료 플랜 한계 이내 기준 |

대부분의 중소·중견 기업 환경(100~500명)에서는 무료 플랜으로 충분히 운영 가능하다. 일일 다운로드가 100,000건을 초과하는 대규모 환경에서만 Cloudflare 유료 플랜($5/월)이 필요하다.

---

## 20. 제약 사항 및 고려 사항

**SharePoint 리스트 뷰 임계값** — 리스트 뷰에서 5,000개 이상의 항목을 한 번에 표시할 수 없다. TransferRecords, DownloadLogs 등 데이터가 많이 쌓이는 리스트에는 TransferId, CreatedAt, SenderEmail 등의 컬럼에 인덱스를 생성하고, 항상 인덱싱된 컬럼으로 필터링하여 조회해야 한다.

**App‑only 인증 시 downloadUrl 미반환** — Sensitivity Labels이 적용된 파일이나 특정 조건에서 `@microsoft.graph.downloadUrl`이 비어 있을 수 있다. 이 경우 Worker가 `/content` 엔드포인트를 통한 Fallback Streaming으로 처리하되, 최대 5 GB로 제한된다.

**OneDrive 파일 삭제와 링크 무효화** — 사용자가 OneDrive에서 파일을 삭제하면 해당 파일의 다운로드 링크가 무효화된다. Worker는 Graph API 404 응답을 받아 "파일이 삭제되었습니다" 페이지를 표시하고, TransferRecords의 Status를 `file_deleted`로 갱신한다.

**IRM/Do Not Forward 제한** — Do Not Forward 및 Sensitivity Labels 암호화는 Outlook 클라이언트에서만 완전히 지원된다. 타 이메일 클라이언트에서는 OME(Office Message Encryption) 포털을 통해 메일을 열람해야 하며, 일부 기능이 제한될 수 있다.

**Sensitivity Labels 라이선스 요구** — Sensitivity Labels은 Microsoft 365 E3/E5, Business Premium, 또는 Microsoft Purview Information Protection 라이선스가 필요하다. 해당 라이선스가 없는 조직에서는 Do Not Forward 및 Exchange Transport Rules만 사용 가능하다.

**Cloudflare Workers KV 쓰기 제한** — 무료 플랜에서 KV 쓰기는 1,000회/일로 제한된다. 토큰 캐시 갱신과 로그 버퍼 등에 사용되므로, 대규모 환경에서는 쓰기 빈도를 모니터링해야 한다.

**Cloudflare Workers CPU 시간 제한** — 무료 플랜에서 요청당 CPU 시간은 10 ms로 제한된다. HMAC 검증과 JSON 처리는 약 2 ms로 충분하지만, Fallback Streaming 시에는 유료 플랜이 필요할 수 있다.

**createUploadSession 청크 크기** — Graph API의 createUploadSession은 청크 크기가 320 KB의 배수여야 한다. 기본 설정인 10 MB(10,485,760 바이트)는 320 KB × 32,768로 이 조건을 충족한다.

**Conditional Access와 서비스 주체** — App‑only 토큰으로 Graph API를 호출하는 서비스 주체가 Conditional Access 정책에 의해 차단되지 않도록 해야 한다. 필요 시 서비스 주체를 정책 예외로 등록한다.

---

## 21. 버전 이력

| 버전 | 날짜 | 주요 변경 |
|------|------|-----------|
| v1.0 | 2026-08-01 | 초기 설계 (수신자별 토큰, 이메일 검증, OTP 포함) |
| v2.0 | 2026-08-15 | 수신자 인증 제거, 기본 무인증 다운로드, 보안은 M365 위임 |
| v2.1 | 2026-08-20 | 토큰 단순화 (13→9 필드), SystemConfig 축소, 로그 컬럼 축소 |
| v2.2 | 2026-08-25 | M365 보안 인프라 100% 위임 확정, KV 쓰기 96% 절감 |
| v2.3 | 2026-09-01 | 개인 폴더를 OneDrive로 이전, PersonalFolders 리스트 삭제 (7→6개), API 라우트 축소 (20→16개), Cron 3개로 확정, Unified Manifest 채택, downloadUrl 실시간 발급 흐름 확정 |
