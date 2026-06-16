# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

N그룹 성과 대시보드 — WM 뇌새김영어 A그룹 코디네이터의 성과를 분석하고 N_A/N_B/N_C/H 등급을 자동 배정하는 브라우저 전용 SPA. 빌드 없이 `index.html` 하나로 동작하며, 서버/백엔드 없음.

## Running

**배포**: GitHub Pages — `wminfobin-byte/wm_dash` 레포 main 루트 → https://wminfobin-byte.github.io/wm_dash/ (main push 시 자동 재배포). 클라우드 싱크도 같은 레포를 사용.

로컬은 브라우저에서 `index.html` 직접 열기, 또는:
```bash
npx http-server -p 8082
```

## Architecture

단일 `index.html` 파일 (CSS + HTML + JS 인라인). CDN 의존성:
- **xlsx-js-style** — 엑셀 읽기/쓰기 + 셀 스타일링
- **html2canvas** — 대시보드 PNG 캡처
- **JSZip** — xlsx 내부 XML 패치 (눈금선 제거)

## Data Flow

### 4종 파일 업로드
| 타입 | 파일 | 보관 정책 |
|---|---|---|
| `db` | WM_고객관리 (DB수) | 누적 (전체 유지) |
| `cont` | WM_고객관리 (계약수) | 누적 (전체 유지) |
| `cancel` | WM_고객관리 (취소율) | 최신 1건만 |
| `attend` | 근태보고/2-8 그룹표기 (명단) | 최신 1건만 |

### 컬럼 매핑 (COL 상수, 0-indexed)
```
obName: 7        // OB담당자 (코디명)
distributeDate: 8 // 분배일자
dbInfo: 15       // DB정보 (재콜 제외 필터)
mediaCode: 19    // 매체코드 (talkis 제외 필터)
lastCallTime: 21
orderDate: 23
customerStatus: 29 // 고객상태 → 계약/취소/제외 분류
```

### 멤버 추출 조건
명단 파일에서 `담당='뇌새김영어'` AND `그룹='A그룹'`인 이름만 추출.

## Storage

### IndexedDB (`nGroupDashboard` v2)
- **files** store: 업로드된 엑셀 파일 원본 (type, fileName, uploadTime, data)
- **results** store: 기간별 계산 결과 (keyPath = sheetName YYMMDD)

### localStorage
- `ngroup_history` — 등급 이력 `{이름: [현재, 이전, ...]}`
- `ngroup_history_times` — 이력 타임스탬프
- `ngroup_last_saved` — 중복 저장 방지용 마지막 sheetName
- `ngroup_history_start` — 이력 기준점 (2026-04-02)

## Core Business Logic

### KPI 계산
```
DB수 = 분배일 기준 카운트 (주말/공휴일, 삭제요청, 재콜, IN무매체, talkis 제외)
실계약수 = 계약 + 취소 모두 포함
계약률 = 실계약수 / DB수
기준취소율 = 취소수 / 계약수 (코디별)
취소율반영계약수 = 실계약수 × (1 - 기준취소율)
취소율반영계약률 = 취소율반영계약수 / DB수  ← 등급 배정 기준
```

### 등급 배정 (`assignGroupsSmart`)
1. 취소율반영계약률 내림차순 정렬
2. 3등분: 상위 → N_A, 중위 → N_B, 하위 → N_C
3. 동률 처리: 같은 성과값은 같은 그룹에 배정

### 특수 규칙
- **H그룹 진입**: N_C 3연속 → H그룹 강등
- **H그룹 복귀**: 직전 H그룹 + 이번 계약 ≥ 2건 → N_C로 복귀
- **신규 보호**: 입사 후 3개월간 N_C/H 배정 시 N_B로 상향 보호
- **수동 수정**: 드롭다운으로 등급 변경 가능 → IndexedDB results에 저장

## Excel Export (`downloadExcel`)

- 기간별 멀티시트 (시트명 = YYMMDD)
- 그룹별 색상: N_A(파랑 #4361EE), N_B(청록 #2EC4B6), N_C(주황 #F0920E), H(빨강 #D1182B)
- 헤더/합계 스타일: 배경 #2D3748, 폰트 흰색
- JSZip으로 xlsx 내부 XML 패치하여 눈금선 제거
- 타이틀 "WM N그룹 성과" D2:J3 병합

## Key Functions

| 함수 | 역할 |
|---|---|
| `getMembers(data)` | 명단에서 N그룹 멤버 추출 |
| `countDb(data, members)` | 코디별 DB수 카운트 |
| `countContracts(data, members)` | 코디별 계약수 카운트 |
| `calcCancelRates(data, members)` | 코디별 취소율 계산 |
| `assignGroupsSmart(stats)` | 성과 기반 등급 자동 배정 |
| `renderDashboard()` | UI 전체 갱신 오케스트레이션 |
| `buildSheet(periodData, isCurrentPeriod)` | 엑셀 시트 1개 생성 |
| `downloadExcel()` | 전체 기간 엑셀 생성 + 다운로드 |
| `getAllResults()` | IndexedDB에서 전체 기간 결과 조회 |

## 고객상태 분류 로직

```
계약: 배송대기, 배송중, 자동구매완료, 구매완료, AS 포함
취소: 반품입고, 승인요청, 입고대기, 환불요청, 환불완료
  → 단, lastCallTime == orderDate이면 제외 (시스템 오류 간주)
제외: 결제보류, 상담취소, 상담요청, 상담대기, 결제요청 등
```

## 부재인입 탭 (mc* 네임스페이스)

`index.html`에는 N그룹 외에도 여러 탭이 인라인돼 있다 (차등성과 `gd*`, 취소율 `cd*`, 유입분석 `ia*`, 일자별 성과 `dp*`, **부재인입 `mc*`**). 데이터는 별도 IndexedDB(`missedCallDashboard`)에 누적 저장하며, 클라우드 싱크(`sync_data.json.gz`)에도 `payload.mc`로 연동됨 (v7~). mc 파일은 용량이 크므로 `파일 관리`에서 불필요한 누적 파일 정리 권장.

- **목적**: 고객 콜백 중 당사가 미수신한 부재 인입 건수·미수신률을 일/주/월·법인별로 추적. 상세 명세 = `부재인입 대시보드_개발설명서.md`.
- **업로드 7종**: IN call(시트 다수·3행 헤더·I열 그룹으로 법인 자동분류·유효통화="연결"&≥90초), 콜백(전체 모수, A 대표번호/B 발신자/E call시간), DB·계약 각각 위버스마인드/더블유케어 별도 슬롯, 대표번호↔법인 매핑(최신 1건). 모두 누적, 동일 이름·크기 파일 재업로드 시 스킵.
- **법인별 컬럼 차이**: `MC_PHONE_COLS` — 위버스마인드 DB/계약 = 차등성과 형식(전화번호 M열만), 더블유케어 = M·N·O열 3개 OR 매칭. 양쪽 공통: 고객키 G열, 분배일 I열, 오더일 X열. 계약은 X열(오더일) 값 있는 행만 유효.
- **부재 판정 (`mcBuildJudged`, 명세서 §4.1)**: 전화번호 정규화(0 제거+숫자만) → 매핑으로 법인 결정(미분류 제외) → 계약 조회: 콜백 이전 오더일 1건이라도 있으면 제외, 모두 이후면 유효 모수 → 계약 없으면 DB 조회(없으면 모수 제외) → IN call에서 발신자+대표번호 일치 & 통화시작 ±2분 → 매칭 0건 or 전부 유효통화 아님 → **부재 인입**.
- **미수신률**: 전체 = 부재콜/전체모수콜, 유효(주력 ★) = 부재콜/유효모수콜.
- **화면**: 기간/법인/집계단위 필터 + 빠른기간, KPI 4종(부재콜/부재고객/전체%·유효%) + 전기 대비 증감, Chart.js 막대(부재콜)+선(유효%) 이중축(법인 전체 시 법인별 스택), 법인별 비교 테이블, 상세 리스트(검색·페이지네이션 50건·CSV 내보내기).
- **알려진 가정**: 콜백 E열은 `yyyy-mm-dd hh:mm:ss` 같은 날짜+시각이어야 함(시각만이면 해당 행 제외+경고). 명세서의 React/Vite/Dexie 스택은 단일 HTML 패턴에 맞춰 바닐라 JS로 이식함.
- **박제(seal) & 원본 정리 (`MC_DB_VER` 2, `sealed` 스토어)**: mc는 사후 변경이 없으므로 과거 판정 레코드(`mcBuildJudged` out)를 값으로 고정 저장하면 영구 유효. 화면 렌더는 `mcSealed`(과거, callTs<`mcSealCutoffMs`) + 라이브(원본 즉석계산, callTs≥cutoff) 병합(`mcGenerate`). 파일관리 모달의 2단계 버튼: ① `mcSealOldData(cutoffISO)`=기준일 미만 박제(원본 유지·검증용), ② `mcPruneSealed`=박제분 IN call(`startTs`)·콜백(`callTs`) 원본 행 삭제(계약·DB·매핑은 최근 판정에 필요해 유지). 기본 기준일=전월 1일(`mcDefaultCutoffISO`). 클라우드 싱크 `payload.mc`는 `{files,sealed,sealCutoffMs}` 객체(구 배열 포맷 하위호환). **동기화 페이로드가 GitHub blob 한도를 넘는 대용량(mc가 최대 기여)일 때 핵심 감량 수단.** 업로드는 청크 분할(`SYNC_PART_BYTES` 20MB·`sync_manifest.json`)로도 한도 회피.

## 계열별성과 탭 (sp* 네임스페이스)

취소율↔부재인입 사이에 배치. DB 계열(0~5, 0=최상위 품질)별로 분배 DB를 **기대계약률 대비 얼마나 잘 살리는지** 분석. 설계서 = `../성과_데이터_추출_기준.md`. 별도 IndexedDB(`seriesPerfDashboard`), 클라우드 싱크 `payload.sp`(v8~).

- **업로드 4종**: `db`(DB수·고객관리, 누적), `contDaily`(실계약·신청자관리 일마감, 누적), `contCum`(취소반영·신청자관리 월간누적, 누적·월별 최신본만 사용), `attend`(명단, 최신 1건). 업로드 시 슬림 행 객체로 파싱해 IDB·싱크 저장.
- **컬럼**: DB = 계열 C(2)·고객키 G(6)·OB명 H(7)·분배일 I(8)·연령대 L(11)·DB정보 P(15)·매체 T(19)·오더일 X(23)·체류시간 Y(24). 계약 = 고객키 D(3)·담당자 N(13)·신청일 P(15)·분류 Q(16)·반품일 AQ(42).
- **집계 모수(DB수)**: 고객키 중복제거 → IN(매체빈값)·재콜(DB정보)·callback/tv_b 제외 → **담당이 뇌새김영어/제2외국어인 사람만**(담당 판정은 **명단 파일** 기준, DB OB명↔명단 이름 풀/베어 매칭). 톡이즈 제외.
- **계열·연령대·체류시간**은 신청자 파일에 없으므로 **DB수 파일에서 고객키 조인(keyMeta)**으로 확보. **계약 귀속 = 계약파일 담당자(N열)**. 기간 = DB 분배일 / 계약 신청일 (flow 방식). DB 미매칭 계약은 계열판정 제외(경고 표시).
- **실계약(gross)** = 일마감 누적(고객키 dedup), **취소반영(net)** = 월간누적 월별최신본의 반품일 빈 행, **취소수** = 실계약−취소반영. 판정기준 토글(**gross 기본**/net).
- **기대구간 판정(`spJudge`)**: 계열별 계약률 구간(0:15%+ / 1:10~15 / 2:6~10 / 3:4~6 / 4:2~4 / **5:1~2**) 대비 — 하한 미만=낭비, 하한±0.5%p=최소수준, 구간내=기대 충족, 상한 이상=기대 이상(0계열은 하한이상=기대 이상). 5계열 낭비(DB 낭비) 기준은 1%(lo=0.01). 판정색: 기대 이상=녹색·기대 충족(구간내)=파랑·최소=노랑·낭비=빨강(매트릭스 셀 가독성 위해 셀 배경 별도 강조).
- **화면**: 기간/담당/센터/집계(일·주·월)/판정기준 필터, KPI 5종(직전 동일기간 대비), 계열요약 카드+이중축차트(DB bar·계약률 line·기대하한 dash), 개인×계열 매트릭스(판정색·정렬), 계열별 활용 랭킹(기대 이상/낭비, 모수 5건↑), 담당/센터 비교차트, 기간추이 차트+표, 이름 클릭 시 **연령대×체류시간 히트맵** 드릴다운. 스타일 엑셀(계열별 종합 + 개인별×계열 2시트, 눈금선 제거).
- **핵심 함수**: `spCompute()`(roster/keyMeta/dbRecs/grossRecs/netRecs 빌드), `spAggBySeries`/`spPerPerson`(집계), `spJudge`(판정), `spRender*`(섹션별 렌더), `spDrill`(세그먼트), `spDownloadExcel`.
- **공용폴더 자동 연동 (File System Access API)**: 네트워크 공유드라이브 폴더를 1회 지정하면 **탭 열 때마다 자동 스캔**해 신규 엑셀을 가져옴. GitHub Pages(HTTPS)·Chromium 계열에서만 동작. 디렉터리 핸들은 IDB `handles` 스토어(`SP_DB_VER` 2)에 영구 보관. 권한은 페이지 새로고침 후 첫 1회 `requestPermission`(사용자 제스처=「지금 스캔」클릭) 필요, 이후 같은 세션은 무인 스캔. 타입 판별=하위폴더명 우선(`DB수/실계약/취소반영/명단`)→파일명 보조(신청자관리는 날짜 8자리=일마감·6자리=월간누적). 중복은 fileHash(name|size)로 스킵, `~$` 임시파일·비 .xlsx 무시, attend는 파일명 최신 1개만. 폴더 진입 단계에서 **알려진 타입(db/contDaily/contCum/attend) 폴더에만** 들어가고 무관 폴더(일일성과 등)는 진입 안 함(노이즈 방지). 월간누적도 자동 스캔 대상(월별 최신본은 `spCompute`가 선택, 중복은 fileHash로 스킵). 판별 실패 파일은 **미인식 모달(`spShowSkipModal`/`spSkipOverlay`)**로 파일명·폴더 표시(배너 「미인식 N개 보기」 링크 또는 사용자 직접 스캔 시 자동 표시). 가져온 파일은 기존 인제스트(`spIngestFile`)·클라우드 싱크에 그대로 합류. 탭 화면 상단 **배너(`spDirBanner`/`spSetBanner`)**: 권한 필요 시 원클릭 「공용폴더 새로고침」 버튼(클릭 제스처로 `requestPermission` 발동), 자동 스캔으로 신규 유입 시 결과 피드백, 신규 0건 무인 스캔이면 숨김. 드릴 모달은 ESC로 닫힘·6계열 한 줄(max-width 1160px). 핵심 함수: `spConnectFolder`/`spScanFolder`/`spAutoScanOnShow`/`spDetectType`/`spVerifyPerm`/`spUpdateDirStatus`/`spSetBanner`/`spShowSkipModal`.

## 타임보드성과 탭 (tb* 네임스페이스)

부재인입↔유입분석 사이(부재인입 옆)에 배치. **naver_tb 타임보드 DB**를 **집행 시간대(8am/1pm 등)별·영어/제2외국어별로** 계약 성과를 본다. 기준 = `../성과_데이터_추출_기준.md`. 별도 IndexedDB(`timeboardDashboard` v1, `files` 스토어). **공용폴더는 계열별성과(sp)와 같은 핸들 재사용**(`spDirHandle`/`spIdbGetHandle`/`spIdbPutHandle`, 피커 `id:'sp-shared'`). **클라우드 싱크 미포함** — 폴더에서 매번 재생성하므로 동기화 페이로드에 부담을 주지 않음(폴더 접근 가능한 PC에서만 표시).

- **업로드 3종**: `db`(DB수·고객관리, 누적), `contDaily`(계약·신청자관리 일마감, 누적), `attend`(명단, 최신 1건). 취소반영(월간누적)은 안 씀. 파서는 sp와 달리 **매체 T(19)·분류 Q(16) 원문 문자열을 보존**(타임보드·시간대·언어 판정용).
- **컬럼**: DB = 고객키 G(6)·OB명 H(7)·분배일 I(8)·DB정보 P(15,재콜)·매체 T(19). 계약 = 고객키 D(3)·담당자 N(13)·신청일 P(15)·분류 Q(16). 명단 파싱은 `spParseAttend` 재사용.
- **타임보드/시간대/언어 판정**: 타임보드=`naver_tb` 포함(DB는 T열, 계약은 Q열). 시간대=문자열 `_` 분할 **마지막 세그먼트**(`tbSlot`, 8am/1pm…, `tbSlotMin`으로 시각 정렬). 언어=`japer|cnper|esper|frper` 포함→제2외국어, 그 외 영어(`tbLang`).
- **DB수(모수)**: 고객키 dedup → OB 빈값/군자/관리자 제외 → 분배일 없음 제외 → **재콜 완전 제외** → 매체 callback/tv_b 제외 → 타임보드만. `dbKey` 맵에 분배일/시간대/언어 보관(조인용). OB명은 명단(`tbBuildRoster`+`spCanon`)으로 정규화하되 슬롯 집계는 이름 무관.
- **계약수**: 일마감 고객키 dedup → 타임보드(Q naver_tb) → **고객키가 '집계된 타임보드 DB'(dbKey)와 매칭되는 것만**(재콜 자동 제외). **집계 기준=신청일**. **일계약=DB 분배일==신청일**, 이전계약=과거 분배 DB에서 신청일에 들어온 계약(=계약수−일계약). 미매칭 계약은 경고 표시.
- **지표**: 시간대별 `DB수 | 계약수 | 계약률 | 일계약수 | 일계약률 | 이전계약수 | 이전계약률`(모두 /DB수). 화면=기간/빠른기간 필터 + KPI 4종 + 시간대별 계약률 차트(영어vs제2, Chart.js) + **전체/영어/제2외국어** 3섹션(각 시간대별 합계표 + `<details>` 일자별 상세표). 스타일은 `tbEnsureStyle`로 1회 주입.
- **핵심 함수**: `tbCompute`(dbRecs/contRecs/dbKey 빌드), `tbAggBy`(slot 또는 day|slot 키 집계), `tbSlotTable`/`tbDailyTable`/`tbLangSection`/`tbRenderKpi`/`tbRenderSlotChart`/`tbRender`, `tbScanFolder`/`tbAutoScanOnShow`/`tbConnectFolder`(sp 폴더 공유), `tbOnShow`/`tbGenerate`/`tbReady`(db 1개 이상). 탭 등록=`switchPage` 훅 + `AUTH_ACCOUNTS` master tabs + 헤더 `headerRight-timeboard-perf`.
