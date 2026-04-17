# N그룹 성과 대시보드 — 벤치마킹 가이드

> 이 문서는 **wm_dash/index.html**의 **N그룹 성과 탭** 로직을 다른 부서·다른 브랜드에 이식하려는 팀을 위한 해설서입니다.
> "내 부서에도 똑같은 걸 만들고 싶다"는 요청에 대한 재사용 포인트를 정리했습니다.

---

## 1. 무엇을 하는 대시보드인가

### 한 줄 요약
**엑셀 4개를 업로드하면 → 코디별 성과를 집계 → 상/중/하 3등급(+페널티 1등급)으로 자동 배정 → 공유용 엑셀로 다운로드**하는 브라우저 전용 SPA.

### 입·출력 흐름
```
[입력]  WM 고객관리 3종 + 근태보고 1종 (.xlsx / .xls)
   ↓
[처리]  브라우저에서 JS로 파싱 → 조인 → 집계 → 등급 배정
   ↓
[출력]  실시간 대시보드 (KPI 카드 + 그룹별 요약 + 정렬 가능 테이블)
        기간별 멀티시트 엑셀 (눈금선 제거 + 그룹 색상)
        PNG 캡처
```

### 왜 이렇게 만들었는가
- 서버/DB 불필요 — 모든 처리가 브라우저에서 일어난다.
- 고객 민감정보(이름·전화번호)는 브라우저 밖으로 나가지 않는다.
- 단일 `index.html` + CDN — 배포가 파일 복사로 끝난다.
- IndexedDB에 파일·계산 결과를 저장 → 새로고침 후에도 유지.

---

## 2. 아키텍처 개요

| 구성 요소 | 역할 | 참고 |
|---|---|---|
| `index.html` | UI + 로직 + 스타일 전부 인라인 | 빌드 단계 없음 |
| **xlsx-js-style** (CDN) | 엑셀 읽기/쓰기 + 셀 스타일 | `XLSX.read`, `XLSX.write` |
| **html2canvas** (CDN) | 대시보드 PNG 캡처 | |
| **JSZip** (CDN) | 엑셀 내부 XML 패치 (눈금선 제거) | |
| **IndexedDB** (`nGroupDashboard` v2) | `files` (원본 엑셀), `results` (기간별 결과) | |
| **localStorage** | 등급 이력(`ngroup_history`), streak 수동 조정 | |

### 핵심 모듈 (파일 하나지만 기능별로 섹션 분리)
```
① 상수      — COL 매핑, 공휴일, 상태 분류 키워드
② IndexedDB — openIDB, idbAdd, idbGetAll, saveResultToIDB
③ 파일 파싱  — handleFile, parseExcel, 업로드/삭제 UI
④ 집계      — getMembers, countDb, countContracts, calcCancelRates
⑤ 등급 배정  — assignGroupsSmart, roundRate
⑥ 후처리    — H그룹 진입/복귀, 신규 보호, 이력/Streak
⑦ 렌더링    — renderKPI, renderGroupSummary, renderTable, renderHistory
⑧ 엑셀 출력 — downloadExcel + buildSheet (+ JSZip 눈금선 패치)
⑨ 수동 수정  — toggleGroupEdit, applyGroupEdit
```

---

## 3. 입력 데이터 (4종)

### 3-1. 보관 정책 (중요)
| 타입 | 내용 | 정책 |
|---|---|---|
| `db` | WM_고객관리 (DB수용) | **누적** — 기간별로 계속 쌓는다 |
| `cont` | WM_고객관리 (계약수용) | **누적** |
| `cancel` | WM_고객관리 (취소율 산정용, 보통 분기 단위) | **최신 1건** |
| `attend` | 근태보고/2-8 그룹표기 (명단) | **최신 1건** |

### 3-2. 컬럼 매핑 (0-indexed)
`const COL = { customerKey:6, obName:7, distributeDate:8, dbInfo:15, mediaCode:19, lastCallTime:21, orderDate:23, customerStatus:29 }`

| 키 | 엑셀 열 | 의미 |
|---|---|---|
| `customerKey` | G(6) | 고객 고유키 — 조인 키 |
| `obName` | H(7) | OB담당자 (코디명) — `(R_대전)이은정` 형식 |
| `distributeDate` | I(8) | 분배일자 |
| `dbInfo` | P(15) | DB정보 (재콜 판별용) |
| `mediaCode` | T(19) | 매체코드 (talkis/공란 판별) |
| `lastCallTime` | V(21) | 마지막 통화시간 |
| `orderDate` | X(23) | 오더일 (계약일) |
| `customerStatus` | AD(29) | 고객상태 → 계약/취소/제외 분류 |

### 3-3. 명단 파일 (엑셀 컬럼명으로 접근)
- `담당`, `그룹`, `이름`, `입사일`
- 추출 조건: `담당 == '뇌새김영어'` AND `그룹 == 'A그룹'`

---

## 4. 데이터 파이프라인

```
[Upload] ─► IDB(files) 저장
            │
            ▼
[Generate 클릭]
  ├─ getMembers(attend) ──────────────► members[], memberHireDates{}
  ├─ countDb(db, members) ─────────────► {이름: DB수}
  ├─ countContracts(cont, members) ────► {이름: 실계약수}
  ├─ calcCancelRates(cancel, members) ─► {이름: 기준취소율}
  │
  ├─ stats = 코디별 {db, contract, contractRate, refCancelRate, refAdjContract, refAdjRate}
  │
  ├─ assignGroupsSmart(stats) ─────────► {이름: 그룹}, cutlines
  ├─ H그룹 복귀 (직전 H + 2건 이상 → N_C)
  ├─ 신규 입사자 보호 (입사월 + 1개월 말일까지 최소 N_B)
  ├─ Streak 계산 (IDB의 이전 기간 결과 스캔)
  ├─ N_C 연속 4회 → H그룹 강등
  │
  ├─ saveResultToIDB({sheetName, rows, ...})
  └─ renderDashboard()
```

---

## 5. 집계 규칙 (매우 중요 — 이식 시 반드시 재정의 포인트)

### 5-1. DB수 (`countDb`)
`분배일자` 기준으로 카운트. **아래 조건 중 하나라도 걸리면 제외**:
- OB명이 공란이거나 `'군자' | '관리자' | '삭제요청'` 포함
- DB정보에 `'재콜'` 포함
- DB정보에 `'in'` 포함 AND 매체코드 공란
- 매체코드에 `'talkis'` 포함
- 분배일이 주말 또는 공휴일
- OB명이 명단의 N그룹 멤버가 아님

### 5-2. 실계약수 (`countContracts`)
`고객상태`에 따라 `계약` 또는 `취소`로 분류된 건을 합산. **조건**:
- 오더일자가 있어야 함 (없으면 계약 자체가 없는 건)
- 오더일이 주말/공휴일이면 제외
- OB/매체코드 제외 조건은 DB수와 동일
- `고객상태`를 `classifyStatus()`로 3분류:

| 분류 | 키워드 | 참고 |
|---|---|---|
| 계약 | `배송대기, 배송중, 자동구매완료, 구매완료, AS` | |
| 취소 | `반품입고, 승인요청, 입고대기, 환불요청, 환불완료` | `lastCallTime == orderDate`면 **제외**(시스템 오류 간주) |
| 제외 | `결제보류, 상담취소, 상담요청, 상담대기, 결제요청, 결제변경후 요청, 신용정보 보류` | |

### 5-3. 기준취소율 (`calcCancelRates`)
별도 취소율 산정용 고객관리 파일(분기 단위)로 코디별 개인 취소율 계산:
```
기준취소율 = 취소수 / (계약수 + 취소수)       # 분모에 오더일 없는 건은 포함 X
기준취소율 = min(기준취소율, 0.4)              # 0.4(40%) 상한 적용
```
상한을 두는 이유: 극단치(신입이 한두 건 취소로 100% 되는 경우 등)가 그룹 배정을 왜곡하는 것 방지.

---

## 6. KPI 계산 공식

```
DB수            = countDb(...)
실계약수         = countContracts(...)        # 계약 + 취소 모두 포함
계약률           = 실계약수 / DB수

기준취소율       = 개인별 min(0.4, 분기 취소수/계약수)
취소율반영계약수 = 실계약수 × (1 - 기준취소율)
취소율반영계약률 = 취소율반영계약수 / DB수     ← 등급 배정 기준 지표
```

> **핵심**: 등급 배정의 단일 기준은 `취소율반영계약률`. "많이 계약하되 취소 적게 하는" 코디가 상위로 간다.

---

## 7. 등급 자동 배정 (`assignGroupsSmart`)

### 7-1. 기본 3분할
1. 전원을 `취소율반영계약률` 내림차순 정렬
2. 상위 1/3 → **N_A**, 중위 1/3 → **N_B**, 하위 1/3 → **N_C**
3. 기본 커트라인: `cut1 = ⌈n/3⌉`, `cut2 = ⌈2n/3⌉`

### 7-2. 동점자 보정 (같은 성과는 같은 그룹)
```javascript
// 커트라인 바로 위 사람과 값이 같으면 커트라인을 한 칸 아래로 밀어낸다
while(cut1<n && roundRate(entries[cut1].rate)===roundRate(entries[cut1-1].rate)) cut1++;
while(cut2<n && roundRate(entries[cut2].rate)===roundRate(entries[cut2-1].rate)) cut2++;
// 경계 역전 방지
if(cut2<=cut1) cut2=cut1+1;
```
> 소수점 2자리까지 반올림한 값이 같으면 "동률"로 본다 (`roundRate`).

### 7-3. 특수 규칙
| 규칙 | 조건 | 결과 |
|---|---|---|
| **H그룹 진입** | Streak(N_C+H 연속) ≥ 4 | N_C → H |
| **H그룹 복귀** | 직전 기간 H + 이번 실계약 ≥ 2 | 자동 배정 결과 무시하고 N_C |
| **신규 보호** | 입사일 + 1개월 말일 이내 | N_C/H 배정 시 N_B로 상향 |

#### H그룹 진입 — Streak 계산 로직
```
Streak 계산의 기준점: '260402' (2026-04-02) — 이 이전 기간은 무시
현재 기간이 N_C 또는 H면 count=1부터 시작
IDB의 이전 기간 결과를 최신부터 역순으로 스캔
  → 같은 사람의 이전 그룹이 N_C 또는 H면 count++, 아니면 break
```
- 자동 계산되므로 localStorage 이력과 별개로 항상 최신 상태 유지
- 수동 조정: `localStorage.ngroup_streak_adj = {이름: 음수}` (직전 H기간 복귀자가 N_C로 돌아온 뒤 카운트 -1 등)

#### 신규 입사자 보호 — 보호 기간 계산
```javascript
// 입사월의 다음 달 말일 23:59:59까지 보호
const protectEnd = new Date(hireDate.getFullYear(), hireDate.getMonth()+2, 0, 23,59,59);
```
- 3월 입사 → 4월 말까지, 4월 입사 → 5월 말까지

---

## 8. 상태 저장 구조

### 8-1. IndexedDB (`nGroupDashboard` v2)
```
files store
  { id(auto), type, fileName, uploadTime, data }
  type ∈ {'db','cont','cancel','attend'}

results store  (keyPath = sheetName; 예: '260416')
  { sheetName, startStr, endStr, rows:[
      { name, group, isProtected, manualGroup,
        db, contract, contractRate,
        refCancelRate, refAdjContract, refAdjRate, ... }
  ]}
```
- **중요**: `results`는 "그 시점의 계산 결과"를 스냅샷으로 저장 → Streak 계산과 과거 시트 엑셀 재생성의 근거.

### 8-2. localStorage
| 키 | 값 |
|---|---|
| `ngroup_history` | `{이름: [최근등급, 이전등급, ...]}` (예전 버전 호환용) |
| `ngroup_history_times` | 이력 타임스탬프 |
| `ngroup_streak_adj` | `{이름: N}` — streak 수동 보정 |
| `ngroup_history_start` | 이력 기준점 (2026-04-02) |

---

## 9. 수동 그룹 수정

- **UI**: "그룹 수정" 버튼 → 각 행의 등급이 드롭다운으로 변경됨 → "적용"
- **저장**: 해당 기간 결과의 `rows`에 `manualGroup:true` 플래그와 함께 IDB에 영속화
- **재생성 시 반영**: `generate()`가 IDB에서 동일 `sheetName`의 `manualGroup` 플래그가 있는 행은 자동 배정 결과를 덮어쓴다 (`savedPeriod.rows` 루프, `generate()` 내부)
- **Streak 카운팅과의 관계**: 수동으로 N_C로 바꿨어도 카운트에 포함됨 (사용자가 원하면 `adjustStreak`으로 -1 보정)

---

## 10. 렌더링

| 섹션 | 함수 | 출력 |
|---|---|---|
| 집계기간 | `renderPeriodRange` | DB 분배일 min/max |
| KPI 5장 | `renderKPI` | 인원/총DB/총실계약/취소율반영계약수/평균반영계약률 |
| 그룹 요약 | `renderGroupSummary` | 4그룹별 인원 + 평균 + 커트라인 |
| 메인 테이블 | `renderTable` | 코디별 상세 (H그룹 최하단, 신규는 B그룹 마지막) |
| 이력 | `renderHistory` | 최근 10기간 배지 표시 |

### 정렬 규칙 (`renderTable` & `buildSheet` 공통)
```
1) H그룹은 무조건 최하단
2) 일반 멤버와 신규 보호 멤버를 분리
3) 각각 refAdjRate 내림차순 정렬
4) 신규 보호 멤버는 N_B그룹의 마지막 위치에 삽입 (없으면 N_A 마지막 뒤)
5) H그룹은 그 뒤에 refAdjRate 내림차순
```

---

## 11. 엑셀 출력 (`downloadExcel`)

### 구조
- **기간별 멀티시트** — IDB `results`의 모든 기간을 시트로 추가
- **시트명** = DB 분배일의 마지막 날, `YYMMDD` 포맷 (예: `260416`)
- **파일명** = `(공유) N그룹 성과_YY년_MM월.xlsx`

### 시트 레이아웃
```
행 1: (빈 행)
행 2~3: D2:J3 병합 — "WM N그룹 성과" 타이틀 (다크 배경)
       B2~C2: "시작일자 : yyyy-mm-dd(요일)"
       B3~C3: "종료일자 : yyyy-mm-dd(요일)"
행 4: (빈 행)
행 5: 헤더 (등급, 코디명, DB수, 실계약수, 계약률, 기준취소율,
             취소율 반영 계약수, 취소율 반영 계약률, N_C그룹 연속)
행 6: 합계
행 7~: 데이터 (정렬 규칙은 §10과 동일)
```

### 스타일링 포인트
- 그룹별 채우기 색: `N_A #4361EE`, `N_B #2EC4B6`, `N_C #F0920E`, `H #D1182B`
- 코디명 셀: 배경은 light variant, 폰트 색은 진한 variant (`lf`)
- 엑셀 셀 포맷: 정수는 `#,##0`, 소수점 한 자리는 `#,##0.0`, 비율은 `0.0%`

### 눈금선 제거 (공통 패턴 — wmdaily/wmdbpick도 동일)
```javascript
const wbBuf = XLSX.write(wb,{bookType:'xlsx',type:'array',bookSST:true});
const zip = await JSZip.loadAsync(wbBuf);
for(const path of Object.keys(zip.files)){
  if(path.startsWith('xl/worksheets/sheet') && path.endsWith('.xml')){
    const xml = await zip.file(path).async('string');
    zip.file(path, xml.replace(/<sheetView(?=[\s>])/g, '<sheetView showGridLines="0"'));
  }
}
const blob = await zip.generateAsync({type:'blob', compression:'DEFLATE', ...});
```
> xlsx-js-style이 `showGridLines` 속성을 제대로 지원하지 않으므로 JSZip으로 XML을 직접 패치.

---

## 12. 다른 도메인에 이식할 때 — 체크리스트

### 반드시 교체해야 할 것
1. **멤버 추출 조건** (`getMembers`)
   ```javascript
   if(dept==='뇌새김영어' && grp==='A그룹' && name){ ... }
   ```
   → 우리 부서·우리 그룹명으로 교체

2. **컬럼 매핑** (`COL` 상수)
   → CMS 엑셀 레이아웃이 다르면 인덱스 재확인

3. **집계 제외 조건** (`countDb`, `countContracts`)
   ```javascript
   if(obName.includes('군자') || obName.includes('관리자') || ...) continue;
   if(dbInfo.includes('재콜')) continue;
   if(mediaCode.toLowerCase().includes('talkis')) continue;
   ```
   → 부서 정책에 맞게 키워드 교체

4. **고객상태 분류** (`classifyStatus`)
   → 계약/취소/제외 키워드 목록

5. **공휴일** (`HOLIDAYS` Set)
   → 집계 범위에 맞는 연도로 갱신

6. **IndexedDB DB명** (`DB_NAME`)
   → 프로젝트별로 고유하게 (다른 탭과 충돌 방지)

### 조정 포인트 (정책에 따라)
- **기준취소율 상한** (현재 0.4) — 도메인에 따라 조정
- **Streak 기준점** (현재 `'260402'`) — 운영 시작일
- **H그룹 진입 임계** (현재 연속 4) — `if(count>=4) groups[name]='H그룹'`
- **H그룹 복귀 계약수** (현재 2) — `if(stats[name].contract >= 2)`
- **신규 보호 기간** (현재 입사월+1개월 말일) — `new Date(y, m+2, 0, ...)`
- **색상 팔레트** — `gColors` 객체 (엑셀), CSS 변수 `--ga/--gb/--gc/--gh`
- **엑셀 파일명·시트명 규칙** — `fmtDateShort`, `downloadExcel` 파일명 템플릿

### 그대로 가져가도 되는 패턴
- 전체 아키텍처 (단일 HTML + IDB + CDN)
- 엑셀 스타일 + JSZip 눈금선 제거 패턴
- 등급 배정의 동점자 보정 알고리즘
- 수동 수정 + `manualGroup` 플래그를 IDB에 영속화하는 패턴
- Streak을 localStorage가 아닌 IDB results에서 역산하는 방식 (이력 꼬임 방지)

---

## 13. 자주 겪는 이슈 & 팁

- **파일 업로드 후 아무 일도 안 일어남** → 4종이 모두 있어야 `generate()`가 동작. 상태 점(header의 status-dots)으로 확인.
- **그룹이 이상하게 배정됨** → `refAdjRate`가 0인 사람이 많은지 확인 (DB수가 0이면 0이 됨). 이 사람들은 최하위로 몰린다.
- **취소율이 비정상적으로 높음** → 40% 상한에 걸렸는지 확인. 분기 데이터가 충분한지.
- **Streak이 틀림** → IDB의 `results` 전체를 확인 (개발자 도구 → Application → IndexedDB). 수동으로 이상 기간을 삭제하거나 `ngroup_streak_adj`로 보정.
- **모바일에서 테이블 잘림** → 가로 스크롤 + 첫 열 sticky 적용됨 (CSS media query 768px/480px). Step 3 리팩터 참고.

---

## 14. 참고 파일 (같은 폴더)
- `index.html` — 실제 구현 (약 5000행, 단일 파일)
- `CLAUDE.md` — 구조 개요 (LLM용 인덱스)
- `N그룹_성과_자동생성_로직.md` — Python 이식 스크립트 가이드 (백엔드용)
- `../CLAUDE.md` — 공통 패턴 (4개 프로젝트 전체)
