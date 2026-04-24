# N그룹 성과 집계 방식 — Claude AI 참고용

> **목적**: N그룹 성과 대시보드(`wm_dash/index.html`)가 사용하는 **KPI 집계 공식·필터 규칙·등급 배정 로직**을 요약한다. 이 문서를 참고하여 Claude AI가 **동일한 방식으로 다른 성과파일**(예: W그룹·E그룹·톡이즈 등 다른 팀, 또는 다른 기간 구성)을 재현할 수 있다.
>
> 실제 코드 위치: `wm_dash/index.html` (단일 HTML 파일, JS 인라인)

---

## 0. 한눈에 보는 파이프라인

```
[4종 입력 엑셀]           [멤버 필터링]          [건별 집계]
  명단(attend)   ──►   뇌새김영어·A그룹      ──►   코디별 DB/계약/취소
  DB수(db)       ──►   에 속한 코디만
  계약수(cont)   ──►                           [지표 계산]
  취소율(cancel) ──►                          계약률 · 기준취소율 · 취소율반영계약률
                                                    │
                                                    ▼
                                              [등급 자동 배정]
                                              취소율반영계약률 순위 → N_A/N_B/N_C
                                                    │
                                                    ▼
                                              [특수 규칙 적용]
                                              H그룹 진입/복귀 · 신규 보호
                                                    │
                                                    ▼
                                              [최종 등급]
```

---

## 1. 입력 파일 4종

| 타입 | 파일명 예시 | 역할 | 보관 정책 |
|---|---|---|---|
| `attend` (명단) | `WM_근태보고_그룹_20260407.xls` | 대상 코디 목록 + 입사일 | 최신 1건 |
| `db` (DB수) | `WM_고객관리_YYYY-MM-DD.xlsx` | 분배일 기준 DB 집계 원본 | 누적 |
| `cont` (계약수) | `WM_고객관리_YYYY-MM-DD.xlsx` | 오더일 기준 계약 집계 원본 | 누적 |
| `cancel` (취소율) | `WM_고객관리_YYYY-MM-DD.xlsx` | 기준취소율 산출용 (더 넓은 기간) | 최신 1건 |

- **db/cont/cancel은 같은 포맷**(WM_고객관리)이지만 집계 목적과 기간이 다르다.
- 명단 파일은 HTML 형식의 `.xls`(pandas read_html로 파싱) 또는 `.xlsx`.

---

## 2. WM_고객관리 컬럼 매핑 (0-indexed)

```js
const COL = {
  customerKey:    6,   // G: 고유 고객키
  obName:         7,   // H: OB담당자 (예: "(R_대전)이은정")
  distributeDate: 8,   // I: 분배일
  dbInfo:         15,  // P: DB정보 (재콜/in 필터용)
  mediaCode:      19,  // T: 매체코드 (talkis 필터용)
  lastCallTime:   21,  // V: 최종통화시간
  orderDate:      23,  // X: 오더일
  customerStatus: 29,  // AD: 고객상태 (배송대기/환불완료 등)
};
```

---

## 3. 멤버 추출 — 명단 파일

**규칙**: `담당='뇌새김영어' AND 그룹='A그룹'`인 행의 `이름` 만 사용.

```js
function getMembers(data){
  const members = [];
  memberHireDates = {};  // 이름 → 입사일 매핑 (신규 보호용)
  for (const r of data){
    if (r['담당']==='뇌새김영어' && r['그룹']==='A그룹' && r['이름']){
      members.push(r['이름']);
      if (r['입사일']) memberHireDates[r['이름']] = parseDate(r['입사일']);
    }
  }
  return members;
}
```

- `이름` 컬럼 형식: `(R_센터)코디명` (예: `(R_대전)이은정`) — DB/계약 파일의 `OB담당자`와 동일 포맷이라 바로 JOIN 가능.
- **다른 팀 성과파일을 만들 때는 이 두 필터만 변경한다** (예: `W그룹`, `톡이즈` 등).

---

## 4. 공통 필터 규칙

DB수·계약수·취소율 집계 시 **전부 동일하게 적용**되는 제외 조건:

| # | 조건 | 적용 시점 |
|---|---|---|
| ① | `OB담당자`에 `군자`/`관리자`/`삭제요청` 포함 | 모든 집계 |
| ② | `매체코드`에 `talkis` 포함 (대소문자 무시) | 모든 집계 |
| ③ | `DB정보`에 `재콜` 포함 | **DB수만** |
| ④ | `DB정보`에 `in` 포함 + `매체코드` 공란 | DB수·계약수 |
| ⑤ | `분배일` 또는 `오더일`이 주말/공휴일 | DB수(분배일), 계약수(오더일) — **취소율은 미적용** |
| ⑥ | 멤버 목록에 없는 `OB담당자` | 모든 집계 |

- **공휴일 판정**: 코드 내 `HOLIDAYS` Set에 하드코딩(2025~2027년). 다른 국가·규칙 적용 시 이 Set을 교체.

---

## 5. 고객상태 분류 (`classifyStatus`)

`고객상태`(29번 열) 문자열을 **계약 / 취소 / 제외 / 기타**로 변환:

| 구분 | 조건 | 예시 |
|---|---|---|
| **제외** | `결제보류`/`상담취소`/`상담요청`/`상담대기`/`결제요청`/`결제변경후 요청`/`신용정보 보류` | — |
| **계약** | `배송대기`/`배송중`/`자동구매완료`/`구매완료` 중 하나 포함 | 또는 `AS` 포함 |
| **취소** | `반품입고`/`승인요청`/`입고대기`/`환불요청`/`환불완료` 중 하나 포함 | **단**, `최종통화시간==오더일`이면 `제외` (시스템 오류 간주) |
| **기타** | 위 어디에도 해당 안 함 | 카운팅 안 함 |

---

## 6. KPI 집계 공식

### 6-1. DB수 (`countDb`)
**정의**: 해당 기간 내 코디에게 분배된 유효 DB 건수.

```
DB수 = count(
  분배일 유효 AND
  공통필터 ①②③④⑤⑥ 통과
)  — 코디별
```

### 6-2. 실계약수 (`countContracts`)
**정의**: 해당 기간 내 코디의 **계약 + 취소 합산** (이미 취소된 건도 한 번은 계약이었으므로 포함).

```
실계약수 = count(
  오더일 유효 AND
  공통필터 ①②④⑤⑥ 통과 AND   ※ 재콜필터 ③은 미적용
  고객상태 ∈ {계약, 취소}
)  — 코디별
```

### 6-3. 계약률
```
계약률 = 실계약수 / DB수  (DB수 0이면 0)
```

### 6-4. 기준취소율 (`calcCancelRates`) — **별도 기간**

취소율 파일(보통 더 긴 누적 기간)에서 **코디별로** 계산:

```
집계 대상 = count(
  오더일 유효 AND
  OB명 ①필터 AND   ※ 주말/공휴일 필터 ⑤ 미적용
  매체코드 ②필터 AND
  고객상태 ∈ {계약, 취소}
)

분자 = (해당 코디의 '취소'수)
분모 = (해당 코디의 '계약+취소' 합계)
개인취소율 = 분자 / 분모

기준취소율(clamp) = min(개인취소율, 0.4)   ※ 40% 상한
```

- **중요**: 기준취소율은 **개인별 과거 취소 이력 기반 40% 상한 clamp 값**이며, 팀 평균을 쓰지 않는다.
- (내부적으로 `avgCancelRate`도 계산하지만 지표 산출에는 **미사용**.)

### 6-5. 취소율반영 계약수 / 계약률 ← **등급 배정의 핵심 지표**

```
취소율반영계약수  = 실계약수 × (1 − 기준취소율)
취소율반영계약률  = 취소율반영계약수 / DB수
```

내부 필드명: `refAdjContract`, `refAdjRate`.

---

## 7. 등급 자동 배정 (`assignGroupsSmart`)

### 7-1. 기본 규칙
1. 모든 코디를 **`refAdjRate` 내림차순**으로 정렬.
2. 균등 3분할 컷라인: `cut1 = ceil(n/3)`, `cut2 = ceil(2n/3)`.
3. `[0, cut1)` → **N_A그룹**, `[cut1, cut2)` → **N_B그룹**, `[cut2, n)` → **N_C그룹**.

### 7-2. 동점자 보정
같은 `refAdjRate`(소수 2자리 반올림 기준)인 사람은 **절대 다른 그룹으로 나뉘지 않는다**:
- A/B 경계의 `cut1` 위치 값이 `cut1−1`과 같으면 → `cut1++` (위 그룹으로 올림)
- B/C 경계도 동일 로직으로 조정

### 7-3. n ≤ 3 예외
상위 1명 N_A, 2등 N_B, 그 외 N_C로 단순 배정.

---

## 8. 특수 규칙

### 8-1. H그룹 진입 (N_C 연속 강등)

N_C/H가 **연속 4번째 기간**에 달하면 → **H그룹 강등**.

```
streak = (현재 기간이 N_C/H이면 1) + (직전부터 거슬러 올라가며 N_C/H 연속 횟수)
if (현재 그룹 == N_C그룹 AND streak >= 4) → 'H그룹'
```

- streak 카운팅 시작점: `STREAK_START = '260402'` (2026-04-02). 그 이전 기간은 무시.
- IndexedDB `results` 스토어의 시트명(YYMMDD) 오름차순으로 역순 탐색.
- 수동 streak 조정값 `localStorage.ngroup_streak_adj`가 있으면 `count + adj`로 보정 후 재평가.

### 8-2. H그룹 복귀
직전 기간이 **H그룹**인 코디가 **이번 기간에 계약 2건 이상**이면 → **N_C그룹**으로 복귀 (H그룹 자동 배정을 덮어씀).

```
if (history[name][0] == 'H그룹' AND stats[name].contract >= 2) → 'N_C그룹'
```

- `history`는 `localStorage.ngroup_history` (이름 → 등급 배열, 최신 순).
- 복귀 판정은 이력 기준점 `HISTORY_START = 2026-04-02 00:00 KST` 이후의 이력만 사용.

### 8-3. 신규 입사자 보호
입사월 **다음 달 말일까지**는 N_C/H 배정 시 **N_B로 상향**.

```js
const protectEnd = new Date(
  hireDate.getFullYear(),
  hireDate.getMonth() + 2,
  0, 23, 59, 59
);  // = (입사월+1)의 마지막 날

if (현재시각 <= protectEnd && groups[name] ∈ {N_C그룹, H그룹})
  → 'N_B그룹' (protectedMembers set에 추가)
```

- 예: 3월 입사 → 4월 말까지 보호. 4월 입사 → 5월 말까지 보호.
- 보호 대상자는 표·엑셀에서 **N_B 그룹 최하단**에 정렬된다.

### 8-4. 적용 순서 (매우 중요)
```
① 기본 배정 (assignGroupsSmart)
② H그룹 복귀 (직전 H + 계약 2건 ≥) → N_C
③ 신규 보호 (입사월 +1 말일까지 N_C/H → N_B)
④ 수동 수정 복원 (IDB results의 manualGroup 플래그)
⑤ streak 재계산 → N_C 4연속 달성 시 H 강등
```

---

## 9. 출력 형식 — 엑셀 시트

### 9-1. 시트 1개 = 성과 기간 1회
- **시트명**: DB파일 분배일 **최대값** → `YYMMDD` (예: `260420`)
- 최신 분배일 + 최신 업로드 순으로 기간 경계 자동 결정.

### 9-2. 시트 상단
- **D2:J3 병합**: 타이틀 "WM N그룹 성과"
- 집계기간: 시작일자(=분배일 min) / 종료일자(=분배일 max)

### 9-3. 데이터 컬럼 순서
```
그룹 | 코디명 | DB수 | 실계약수 | 계약률 | 기준취소율 |
취소율반영 계약수 | 취소율반영 계약률 | N_C연속 | (여유열)
```
※ 코디명은 `(R_센터)이름` 포맷 유지.

### 9-4. 행 정렬
1. 그룹 순: **N_A → N_B → N_C → H**
2. 그룹 내: `refAdjRate` 내림차순
3. **신규 보호자**는 N_B 그룹 최하단
4. **합계행**: 첫 데이터 행 위쪽 (집계값 sum)

### 9-5. 색상 규칙
| 그룹 | 배경색(HEX) | 폰트 |
|---|---|---|
| N_A | `#4361EE` (파랑) | 흰색 |
| N_B | `#2EC4B6` (청록) | 흰색 |
| N_C | `#F0920E` (주황) | 흰색 |
| H | `#D1182B` (빨강) | 흰색 |
| 헤더/합계 | `#2D3748` | 흰색 |

### 9-6. 눈금선 제거 (Excel 기술 노트)
`xlsx-js-style`이 `showGridLines`를 제대로 출력하지 않아, JSZip으로 xlsx를 열어 `xl/worksheets/sheet*.xml`의 `<sheetView>` 태그에 `showGridLines="0"`을 직접 삽입한다. (wmdaily/wm_dash 공통 패턴)

```javascript
const wbBuf = XLSX.write(wb, { bookType:'xlsx', type:'array', bookSST:true });
const zip = await JSZip.loadAsync(wbBuf);
for (const path of Object.keys(zip.files)){
  if (path.startsWith('xl/worksheets/sheet') && path.endsWith('.xml')){
    const xml = await zip.file(path).async('string');
    zip.file(path, xml.replace(/<sheetView /g, '<sheetView showGridLines="0" '));
  }
}
const blob = await zip.generateAsync({ type:'blob' });
```

---

## 10. 영속성 (브라우저 저장소)

| 키 | 스토리지 | 용도 |
|---|---|---|
| `nGroupDashboard` v2 > `files` | IndexedDB | 업로드된 엑셀 원본 보관 |
| `nGroupDashboard` v2 > `results` | IndexedDB | 기간별 결과 (keyPath: YYMMDD 시트명) |
| `ngroup_history` | localStorage | `{이름: [최신등급, 이전등급, ...]}` |
| `ngroup_history_times` | localStorage | 각 이력 저장 타임스탬프 |
| `ngroup_history_start` | localStorage | 이력 유효 기준점 (`2026-04-02` 이후만 사용) |
| `ngroup_streak_adj` | localStorage | 수동 streak 조정치 `{이름: ±N}` |
| `ngroup_last_saved` | localStorage | 중복 저장 방지용 마지막 시트명 |

---

## 11. 다른 성과파일을 만들 때 변경 포인트 체크리스트

| 항목 | 수정 위치 | 값 |
|---|---|---|
| **대상 팀** | `getMembers()`의 필터 조건 | `담당==='뇌새김영어' && 그룹==='A그룹'` 변경 |
| **OB명 예외** | `countDb`/`countContracts`/`calcCancelRates` | `군자`/`관리자`/`삭제요청` — 팀별 조정 |
| **매체 제외** | 위 3함수 | `talkis` — 다른 팀에선 빼거나 다른 키워드로 |
| **DB정보 제외** | `countDb` | `재콜`, `in+공란` — 팀 규칙 반영 |
| **공휴일** | `HOLIDAYS` Set | 국가·연도 추가 |
| **고객상태 키워드** | `STATUS_EXCLUDE` / `STATUS_CONTRACT` / `STATUS_CANCEL_CHECK` | 제품별 상태명 대응 |
| **기준취소율 상한** | `Math.min(…, 0.4)` | 40% → 팀 규칙에 맞춰 조정 |
| **등급 분할** | `assignGroupsSmart`의 `ceil(n/3)` | 3분할 → 2분할/5분할 등 |
| **등급 이름·색** | `groups` 문자열 + 엑셀 색 매핑 | `N_A` → `W_A` 등 |
| **H그룹 규칙** | streak ≥ 4, 복귀 2건, 보호 +1개월말 | 팀 규칙에 맞춰 상수 조정 |
| **시트명 포맷** | `fmtDateShort(maxDistDate)` | `YYMMDD` → 다른 포맷 |

---

## 12. 핵심 함수 시그니처 요약

```js
// 멤버
function getMembers(attendData): string[]
// 입력: 명단 원본 JSON(row-객체 배열)
// 출력: ["(R_대전)이은정", …]

// 집계
function countDb(dbData, members): { [name]: number }
function countContracts(contData, members): { [name]: number }
function calcCancelRates(cancelData, members): { [name]: rate }

// 등급
function assignGroupsSmart(stats): {
  groups:   { [name]: 'N_A그룹'|'N_B그룹'|'N_C그룹' },
  cutlines: { aToBIdx, bToCIdx, aToBVal, bToCVal }
}

// 분류
function classifyStatus(status, lastCall, orderDt): '계약'|'취소'|'제외'|'기타'

// 유틸
function isHoliday(d): boolean         // 주말+공휴일
function parseDate(v): Date|null       // 엑셀/문자열 → Date
function fmtDateShort(d): 'YYMMDD'     // 시트명용
function pct2(v): string               // 소수 → "12.3%"
```

---

## 13. 재현 절차 — "다른 성과파일" 만드는 최소 단계

1. **입력 파일 4종 확보** (명단·DB수·계약수·취소율)
2. 명단에서 대상 팀 멤버 추출 (§3, 필터만 교체)
3. 공통 필터(§4) + 고객상태 분류(§5) 적용
4. 코디별 DB수·실계약수·개인취소율 계산 (§6-1~4)
5. `refAdjContract`, `refAdjRate` 계산 (§6-5)
6. `refAdjRate` 순위 → 3분할 + 동점자 보정 → 기본 등급 (§7)
7. 특수 규칙 순서대로 적용 (§8-4)
8. 시트 1장 출력: 타이틀/기간/컬럼/행 정렬/색상/눈금선 제거 (§9)

---

## 14. 자주 틀리는 지점 (검증 체크리스트)

- [ ] `실계약수`에 **취소 건도 포함**했는가? (계약+취소)
- [ ] `기준취소율` 분모는 **취소율 파일 기간** 기반인가? (DB·계약 기간과 다름)
- [ ] 기준취소율에 **40% 상한 clamp**를 걸었는가?
- [ ] DB수만 **재콜 필터**가 적용되는가? (계약수·취소율은 미적용)
- [ ] 취소율 집계에는 **주말/공휴일 필터를 적용하지 않았는가**?
- [ ] 동점자가 생겼을 때 **같은 그룹으로 묶이도록** cut 경계가 조정됐는가?
- [ ] H그룹 복귀(N_C)가 신규 보호(N_B)보다 **먼저** 적용됐는가?
- [ ] streak 카운팅 시작점 이전 기간을 **무시**했는가?
- [ ] 엑셀 시트명이 DB `분배일 최대값`의 `YYMMDD` 포맷인가?
- [ ] `showGridLines="0"` 패치가 적용됐는가?
