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
