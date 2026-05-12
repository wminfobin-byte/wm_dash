# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

N그룹 성과 대시보드 — WM 뇌새김영어 A그룹 코디네이터의 성과를 분석하고 N_A/N_B/N_C/H 등급을 자동 배정하는 브라우저 전용 SPA. 빌드 없이 `index.html` 하나로 동작하며, 서버/백엔드 없음.

## Running

브라우저에서 `index.html` 직접 열기, 또는:
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

`index.html`에는 N그룹 외에도 여러 탭이 인라인돼 있다 (차등성과 `gd*`, 취소율 `cd*`, 유입분석 `ia*`, 일자별 성과 `dp*`, **부재인입 `mc*`**). 부재인입은 클라우드 싱크 미연동 — 데이터는 별도 IndexedDB(`missedCallDashboard`)에 로컬 누적만.

- **목적**: 고객 콜백 중 당사가 미수신한 부재 인입 건수·미수신률을 일/주/월·법인별로 추적. 상세 명세 = `부재인입 대시보드_개발설명서.md`.
- **업로드 7종**: IN call(시트 다수·3행 헤더·I열 그룹으로 법인 자동분류·유효통화="연결"&≥90초), 콜백(전체 모수, A 대표번호/B 발신자/E call시간), DB·계약 각각 위버스마인드/더블유케어 별도 슬롯, 대표번호↔법인 매핑(최신 1건). 모두 누적, 동일 이름·크기 파일 재업로드 시 스킵.
- **법인별 컬럼 차이**: `MC_PHONE_COLS` — 위버스마인드 DB/계약 = 차등성과 형식(전화번호 M열만), 더블유케어 = M·N·O열 3개 OR 매칭. 양쪽 공통: 고객키 G열, 분배일 I열, 오더일 X열. 계약은 X열(오더일) 값 있는 행만 유효.
- **부재 판정 (`mcBuildJudged`, 명세서 §4.1)**: 전화번호 정규화(0 제거+숫자만) → 매핑으로 법인 결정(미분류 제외) → 계약 조회: 콜백 이전 오더일 1건이라도 있으면 제외, 모두 이후면 유효 모수 → 계약 없으면 DB 조회(없으면 모수 제외) → IN call에서 발신자+대표번호 일치 & 통화시작 ±2분 → 매칭 0건 or 전부 유효통화 아님 → **부재 인입**.
- **미수신률**: 전체 = 부재콜/전체모수콜, 유효(주력 ★) = 부재콜/유효모수콜.
- **화면**: 기간/법인/집계단위 필터 + 빠른기간, KPI 4종(부재콜/부재고객/전체%·유효%) + 전기 대비 증감, Chart.js 막대(부재콜)+선(유효%) 이중축(법인 전체 시 법인별 스택), 법인별 비교 테이블, 상세 리스트(검색·페이지네이션 50건·CSV 내보내기).
- **알려진 가정**: 콜백 E열은 `yyyy-mm-dd hh:mm:ss` 같은 날짜+시각이어야 함(시각만이면 해당 행 제외+경고). 명세서의 React/Vite/Dexie 스택은 단일 HTML 패턴에 맞춰 바닐라 JS로 이식함.
