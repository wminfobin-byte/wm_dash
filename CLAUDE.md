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

부재인입↔유입분석 사이(부재인입 옆)에 배치. **naver_tb 타임보드 DB**를 **집행 시간대(8am/1pm 등)별·영어/제2외국어별로** 계약 성과를 본다. 기준 = `../성과_데이터_추출_기준.md`. 별도 IndexedDB(`timeboardDashboard` v1, `files` 스토어). **공용폴더는 계열별성과(sp)와 같은 핸들 재사용**(`spDirHandle`/`spIdbGetHandle`/`spIdbPutHandle`, 피커 `id:'sp-shared'`). **클라우드 싱크 포함**(`payload.tb` 독립 슬라이스, `tbCollectSync`/`tbRestoreSync`, 헤더 `syncBtn8`) — 폴더 없는 PC/사람도 봄. sp와 같은 파일을 읽으므로 싱크에 일부 중복되나 청크 분할로 감당.

- **업로드 3종**: `db`(DB수·고객관리, 누적), `contDaily`(계약·신청자관리 일마감, 누적), `attend`(명단, 최신 1건). 취소반영(월간누적)은 안 씀. 파서는 sp와 달리 **매체 T(19)·분류 Q(16) 원문 문자열을 보존**(타임보드·시간대·언어 판정용).
- **컬럼**: DB = 계열 C(2)·고객키 G(6)·OB명 H(7)·분배일 I(8)·DB정보 P(15,재콜)·매체 T(19). 계약 = 고객키 D(3)·담당자 N(13)·신청일 P(15)·분류 Q(16). 명단 파싱은 `spParseAttend` 재사용.
- **타임보드/시간대/언어 판정**: 타임보드=`naver_tb` 포함(DB는 T열, 계약은 Q열). 시간대=문자열 `_` 분할 **마지막 세그먼트**(`tbSlot`, 8am/1pm…, `tbSlotMin`으로 시각 정렬). 언어=`japer|cnper|esper|frper` 포함→제2외국어, 그 외 영어(`tbLang`). **영어 세부**(`tbEngSub`): 매체코드에 `tw2` 포함→더위크, 그 외→결합. 언어 필터(`tbLangView`/`tbInRange`) = 전체·영어(더위크+결합)·더위크·결합·제2외국어. 개인별 엑셀은 영어/더위크/결합/제2외국어 4시트.
- **퇴사자 표기**: 명단(roster)에 없는 코디(`spCanon` 미매칭=`rec.left`)는 퇴사자. `tbModel.leftNames` Set으로 개인별 표 행에 회색+「퇴사」뱃지(`tbIsLeft`), 엑셀은 이름에 `(퇴사)` 접미+회색 폰트.
- **DB수(모수)**: 고객키 dedup → OB 빈값/군자/관리자 제외 → 분배일 없음 제외 → **재콜 완전 제외** → 매체 callback/tv_b 제외 → 타임보드만. `dbKey` 맵에 분배일/시간대/언어 보관(조인용). OB명은 명단(`tbBuildRoster`+`spCanon`)으로 정규화하되 슬롯 집계는 이름 무관.
- **계약수(분배일 코호트·DB 앵커)**: 일마감 고객키 dedup → **고객키가 '집계된 타임보드 DB'(dbKey)와 매칭되는 것만**(재콜 자동 제외). **매칭된 DB의 분배일·시간대·언어로 귀속**(계약의 Q열은 분류에 쓰지 않음 — 계약 Q와 DB 매체 T가 달라 언어·시간대가 어긋나던 버그 해결). 계약률=그 분배일 코호트의 전환율. **일계약=DB 분배일==신청일**, 이전계약=분배일≠신청일(이후 계약, =계약수−일계약). 미매칭 계약은 경고 표시.
- **지표**: 시간대별 `DB수 | 계약수 | 계약률 | 일계약수 | 일계약률 | 이전계약수 | 이전계약률`(모두 /DB수). **표본 필터: DB수 ≤ `TB_MIN_DB`(=5)인 행(시간대·집행일)은 리스트·차트에서 제외**(합계도 남은 행만 재계산, 차트는 막대 null). KPI는 전체 총계 유지.
- **화면 레이아웃**: 기간/빠른기간 필터 + KPI 4종(선택 언어 반영) + 시간대별 계약률 차트(영어vs제2, Chart.js) + **전체/영어/제2외국어** 3섹션. 각 섹션 = ① `tbSlotTable` 시간대 요약표(slot×지표) + ② **`tbSlotDateBlocks` 시간대별 분배일 추이** — *시간대를 1차 기준*으로 분배일을 **시간순(오래된→최신)** 나열 + ③ **개인별 성과**(`tbPersonTable`, 코디별·계약률 내림차순, 계약률 옆 `펼치기`로 코디×계열 0~5/미상 분해 `tbPersonDetail`). 코디·계열 귀속=매칭된 DB(`contRecs.name/sr`). 개인별 엑셀(`tbDownloadPersonExcel`)=영어·제2외국어 2시트. 스타일은 `tbEnsureStyle`로 1회 주입.
- **구버전 db 재파싱**: 계열(sr) 없이 파싱된 옛 db 파일은 `tbFileStale`로 감지 → 폴더 재스캔(`tbScanFolder`/자동) 시 같은 fileHash라도 삭제 후 재파싱(공용폴더 새로고침만으로 계열 채움). 폴더 없는 PC는 싱크로 전파.
- **핵심 함수**: `tbCompute`(dbRecs/contRecs/dbKey 빌드), `tbAggBy`(slot 또는 slot|day 키 집계), `tbSlotTable`/`tbSlotDateBlocks`/`tbLangSection`/`tbRenderKpi`/`tbRenderSlotChart`/`tbRender`, `tbScanFolder`/`tbAutoScanOnShow`/`tbConnectFolder`(sp 폴더 공유), `tbOnShow`/`tbGenerate`/`tbReady`(db 1개 이상). 탭 등록=`switchPage` 훅 + `AUTH_ACCOUNTS` master tabs + 헤더 `headerRight-timeboard-perf`.

## 뇌새김 콜 현황 탭 (nc* 네임스페이스) — DB 단계별 콜

톡이즈 성과 옆. **CTI 통화 × 고객관리 DB × 신청자관리 계약**을 매칭해 콜을 **DB 단계별 5분류**로 본다. 별도 IndexedDB(`naesaegimCallDashboard` v1), 공용폴더는 sp/tb/tz와 핸들 공유(`spDirHandle`), 클라우드 싱크 `payload.nc`.

- **업로드 4종**: `db`(고객관리·공용폴더·누적), `contDaily`(신청자관리 일마감·공용폴더·누적), `attend`(명단·누적 합집합), `cti`(걸은전화상세·**수동만**·누적).
- **DB 컬럼**: 계열 C(2)·고객키 G(6)·OB명 H(7)·분배일 I(8)·**유입일 J(9)**·연령대 L(11)·전화 M/N/O(12/13/14)·DB정보 P(15)·매체 T(19)·오더일 X(23). 파싱 시 `diHist`(DB정보 원문 히스토그램)를 파일 레코드에 요약 보관 — 원문을 행마다 저장하지 않고 판정 근거만 검증 가능.
- **재콜 단계 판정 (`ncRecallLevel`)**: DB정보(P)에 `재콜` 출현 횟수 = 0신규/1재콜/2재재콜/3회수콜. 실제 값은 `(부재재콜)` `(통화재콜)` `(통예재콜)` 3종이 각각 1단계이고 **`(재분배)`는 단계에 포함되지 않는다**(담당자 재배정). 예 `(부재재콜)(부재재콜)(통화재콜)(재분배)` → 회수콜. 모달 `ncShowDiCodes`로 원문×건수×판정단계 검증.
- **고객키는 단계마다 새로 발급**된다(일별 고객관리 파일 간 키 중복 0). 따라서 **리드 = 대표전화 + 유입일**(`ncLeadKey`)로 묶어야 생애주기가 보인다. 단계 전환 시 **담당자가 바뀌는 리드가 98%**.
- **콜↔DB 매칭 (핵심)**: 키 = **담당자 + 고객연락처 + (분배일 ≤ 통화일) 중 분배일 최신**. `pcIdx`(`전화|담당자` → 이벤트[])로 조회. 담당자를 안 보면 내가 건 신규 DB 콜이 남이 받아간 재콜 DB 콜로 오분류된다. 담당자 불일치 = `xcodiCall`(제외), 전화 미매칭 = `unmatchedCall`, 분배 전 통화 = `preDistCall`. **케어통화 제외 유지**(실계약 고객 & 콜일 ≥ 최초 신청일; 단 분배 당일 신규 콜은 sales 콜로 우선).
- **콜 5분류 (`ncCatOf`)**: `d0` 당일(rc0 & 분배일==통화일) / `d1` 이전(rc0 & 이후 — 재콜로 넘어가기 전까지) / `r1` 재콜 / `r2` 재재콜 / `r3` 회수콜. 계약도 같은 방식으로 자기 DB 이벤트에서 `cat` 부여.
- **집계 축 2가지 (`ncStageAgg`)**: `cohort`=분배일 기준(그날 분배 DB를 이후 통화까지 추적 — 컨택률·계약률 정합) / `call`=통화일 기준(활동량, 분배DB·컨택률 컬럼 숨김). 버튼 토글(`ncStageAxis`). **`d1`의 분배DB는 `d0`와 같은 신규 코호트**라 합계에서 중복 제외(`ncStageTotal`), 표에선 괄호 회색 표기.
- **지표**: 분배DB · 시도DB(distinct) · 컨택률 · 시도수 · 평균시도 · 연결/연결률 · 유효/유효율 · 계약 · 계약률(분배DB/시도DB) · **시도1천건당 계약**(`ncPer1k`, 단계별 콜 효율).
- **분석 섹션**: ① `ncRenderStage` 단계별 표(+일자별/코디별 × 단계 탭) ② `ncRenderEff` 시도 효율·기회비용(차트 `ncEffChart`) ③ `ncRenderD0` 당일 컨택 유무별 계약률 ④ `ncRenderFunnel` 리드 생애주기 퍼널 ⑤ `ncRenderIdle` 미시도/저시도 DB. 이후 기존 언어별 매트릭스·일자별·추이차트·개인별.
- **데이터 품질 가드**: `ncModel.callCodis`(CTI 기록 1건 이상인 코디)에 없는 코디의 DB는 ③④⑤에서 **제외**하고 사유를 화면에 표기 — 실제 미시도인지 CTI 수집 누락인지 구분 불가하기 때문. 당일컨택 리프트는 「당일 유효통화」 vs 「당일 시도·유효통화 없음」으로만 계산(둘 다 CTI 확인된 모수).
- **구버전 재파싱**: `ncPurgeStaleDb`(유입일 `inf` 없는 db 파일) / `ncPurgeStaleContracts`(분류 `cls` 없는 계약 파일) — **공용폴더 확인 후에만** 삭제 후 재스캔. 폴더 없는 PC는 싱크로 전파.
- **엑셀**: `단계별콜`(코호트 기준 합계·단계별·코디×단계) + `언어별종합` + `개인별` 3시트.

## 톡이즈 성과 탭 (tz* 네임스페이스)

타임보드성과 옆에 배치. **CTI(걸은전화상세) 통화 데이터 × 톡이즈 DB/계약**을 매칭해 **톡이즈 계약 하락 원인**(영업코디가 기존계약자 케어통화=학습케어에 시간을 뺏겨 신규 컨택이 떨어지는지)을 데이터로 검증. 별도 IndexedDB(`talkisPerfDashboard` v1, `files` 스토어). **공용폴더는 sp/tb와 같은 핸들 재사용**(`spDirHandle`, 피커 `id:'sp-shared'`). **클라우드 싱크 포함**(`payload.tz` 독립 슬라이스, v9~, 헤더 `syncBtn9`).

- **업로드 4종**: `cti`(걸은전화상세, **수동 업로드만**·누적), `db`(DB수·고객관리, 공용폴더·누적), `contDaily`(계약·신청자관리 일마감, 공용폴더·누적), `attend`(명단, 최신 1건). 폴더 스캔 WANT=db/contDaily/attend(CTI 제외).
- **톡이즈 판별**: 명단 담당=톡이즈(또는 그룹 톡이즈·센터 송내)로 roster 추출(`tzBuildRoster`/`tzIsTalkis`) + CTI 상담원명 집합으로 보강 → `tzNames`. DB OB명·계약 담당자를 `tzBare`(괄호 제거)로 정규화해 `tzNames` 매칭되는 것만 톡이즈로 집계.
- **CTI 컬럼**(시트마다 헤더행 위치 다름 — '수신자번호'/'상담원명' 포함행 자동 탐지, 1시트는 1행 검색조건 가비지·2행 헤더, 2시트부터 1행 헤더): 상담원명 C(2)·일자 E(4)·통화시간 H(7)·수신자번호 K(10,고객전화)·상태 L(11). 이름은 `(R_송내)` 없이 이름만 → DB/계약과 `tzBare`로 정규화 매칭.
- **콜 지표**: 시도수=전체, 연결수=상태 'Out 연결', 유효통화수=연결 & 통화시간 ≥ 90초(`TZ_VALID_SEC`). 통화시간 `tzDurSec`(HH:MM:SS→초), 전화 `tzNormPhone`(숫자만·10자리 0보정, 양쪽 동일).
- **매칭·분류(`tzCompute`, 콜 1건 단위)**: CTI 전화 K ↔ 고객관리 전화 M/N/O(12/13/14) 조인(`phoneIdx`). 콜일(E) vs 그 고객 DB 분배일(I)/오더일(X)로. **판정 순서 = fresh → care → old**: ① **fresh**(분배일==콜일 & 비재콜) — **당일 분배면 같은 날 오더가 나도 '계약 위한 sales 콜'로 보고 케어보다 우선**(케어 우선분류 문제 해결) / ② **care**(실계약 고객[신청자관리·반품일 없음] & 콜일≥**신청일(신청자관리 P열)** — 고객관리 오더일 X는 일부 비어/깨져 부정확해 신청일 사용, `contractOrder`=미취소계약 최초 신청일 / 즉 분배 다음날 이후 사후통화) / ③ **old**(콜일>분배일 또는 재콜). best DB=콜일 이하 분배일 최댓값(동률이면 비재콜 우선). 미매칭 전화=제외(경고). **재콜(고객관리 P열)이면 fresh 불가→old**. 발신자(상담원)가 명단 톡이즈 roster가 아니면 skip(집계 제외).
- **계약수(신청일 기준, DB정보 무시)**: 신청자관리 신청일==그날 발생 계약. 당일계약(cat2)=분배일==신청일 **AND 신청자관리 R열(17) 코드가 재콜 아님**, 그 외/재콜=이전계약(cat3). 즉 통화는 재콜이면 무조건 이전, 계약은 신청일 기준이되 **신청자관리 R열 재콜 코드면 이전계약**.
- **화면**: 글로벌 필터 = 기간(start/end)+빠른기간 + **코디 멀티체크**(`tzCodiPanel`, 전체선택/해제, 콜·계약 양쪽에 적용). KPI 6종(케어통화/케어통화시간/신규시도/신규유효/신규계약/신규계약률). 일자별 스택 차트(케어·신규당일·신규이전재콜 막대 + 신규계약 선, 이중축). 섹션: ① 케어통화(일자별 시도/연결/연결률/유효/총통화시간/대상고객수/고객당·콜당 평균시간) ② 당일 DB 시도 현황 ③ 이전 DB 시도 현황 ④ 신규 합계(②+③) ⑤ 개인별 성과. ②③④엔 DB수(distinct 시도 DB)·평균시도수(시도/DB)·계약률(계약/DB) 포함. **모든 표 합계행=헤더 바로 아래(상단)**, 첫 열 가운데 정렬. y축=날짜.
- **케어 평균시간 기준**: 총통화시간=연결콜 통화시간 합. **유효대상고객수=유효통화(≥90초) 받은 distinct 고객**(`careCust`는 `c.valid`일 때만 add). 고객당 평균시간=총통화시간/유효대상고객, **콜당 평균시간=총통화시간/유효통화수**(연결수 아님). KPI 케어통화시간 sub도 유효콜당 평균.
- **개인별 성과(`tzPersonTable`)**: 코디별 행, **케어/합계/당일/이전 4그룹** 그룹헤더(신규 3그룹 = DB수·시도·평균시도·연결·연결률·유효·유효율·계약·계약률, 케어그룹은 시도·연결·유효·통화시간·평균케어시간=시간/유효수). 전체 열=33(라벨1+케어5+신규9×3). 코디 왼쪽 ▶로 **일자별 추이(오래된→최신, `tzPersonDailyTable`/`tzAggPersonDates`)** 펼치기 → 특정 코디의 케어통화 최근 증가 여부 검증. 케어(cat1)와 신규(cat2+cat3)는 콜별 배타 분류라 완전 별개. 헬퍼 `tzGroupsOf`/`tzGrpCells`/`tzCareCells`/`tzFullRow`/`tzFullHeader`, 전체집계 `tzAggAll`(distinct DB Set 보존).
- **핵심 함수**: `tzCompute`(roster/keyRec/phoneIdx/contracts/콜분류), `tzAggByDate`/`tzAggByPerson`/`tzSumAgg`(집계), `tzCareTable`/`tzNewTable`(fresh/old/new)/`tzPersonTable`/`tzRenderChart`/`tzRenderKpi`/`tzRender`, `tzScanFolder`/`tzAutoScanOnShow`/`tzConnectFolder`(sp 폴더 공유), `tzOnShow`/`tzGenerate`/`tzReady`(cti+db 필요). 탭 등록=`switchPage` 훅 + `AUTH_ACCOUNTS`(info_bin·admin) + 헤더 `headerRight-talkis-perf`.
