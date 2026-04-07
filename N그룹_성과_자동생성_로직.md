# WM N그룹 성과 파일 자동 생성 — 개발 로직 문서

> VS Code + Python 환경에서 실행 가능한 스크립트 개발 가이드입니다.

---

## 1. 개요

### 목적
매 성과 산정 기간이 끝날 때마다 **N그룹 성과 Excel 파일**을 자동으로 생성한다.  
기존 파일(`N그룹_성과_26년_3월.xlsx`)의 시트 구조·포맷을 그대로 유지한다.

### 입력 파일 3종
| 파일 | 용도 |
|------|------|
| `N그룹_성과_YY년_MM월.xlsx` | 기존 성과 파일 (템플릿 및 이전 데이터) |
| `WM_기간별_취소율_YYMMDD-YYMMDD_ver*.xlsx` | 개인별 취소율 원본 |
| `WM_근태보고_그룹_YYYYMMDD.xls` (HTML 형식) | N그룹 구성원 판별용 |

---

## 2. 파일 구조 분석

### 2-1. N그룹 성과 파일 시트 목록
```
DB수 / 계약수 / 취소율 / 성과
날짜별 시트 (예: 260330, 260325, 260320 …)
```

#### 날짜별 시트 (예: `260330`) 구조 — 핵심 시트
| 열(col) | 내용 |
|---------|------|
| B (1) | 그룹 (N_A그룹 / N_B그룹 / N_C그룹 / H그룹) |
| C (2) | 코디명 예) `(R_대전)이은정` |
| D (3) | DB수 |
| E (4) | 실계약수 |
| F (5) | 계약률 = 실계약수 / DB수 |
| G (6) | 취소수 (해당 기간 실취소) |
| H (7) | 취소율 (해당 기간 실취소율) |
| I (8) | 취소 반영 계약수 = 실계약수 - 취소수 |
| J (9) | 취소 반영 계약률 = 취소 반영 계약수 / DB수 |
| K (10) | 기준 취소율 (WM 기간별 취소율 파일에서 가져옴) |
| L (11) | 취소율 반영 계약수 = 실계약수 × (1 - 기준취소율) |
| M (12) | 취소율 반영 계약률 = 취소율 반영 계약수 / DB수 |

행 구성:
- 행 0~2: 헤더 영역 (제목, 시작일자, 종료일자)
- 행 3: 컬럼 헤더 (그룹, 코디명, DB수, 실계약수 …)
- 행 4: 합계행
- 행 5~: 개인별 데이터 (N_A그룹 → N_B그룹 → N_C그룹 → H그룹 순, 그룹 내 취소율 반영 계약률 내림차순)

#### `성과` 시트 구조
| 열 | 내용 |
|----|------|
| B | 그룹 (None / 합계) |
| C | 코디명 |
| D | DB수 (누적) |
| E | 실계약수 |
| F | 계약률 |
| G | 기준 취소율 |
| H | 취소율 반영 계약수 |
| I | 취소율 반영 계약률 |
| N (13) | N_A그룹 레이블 |
| P (15) | N_A그룹 멤버 이름 |
| R (17) | N_B그룹 멤버 이름 |
| T (19) | N_C그룹 멤버 이름 |
| U~V | 휴무일 / 휴무사유 |

---

### 2-2. WM 기간별 취소율 파일 구조
시트: `명단` / `RAW` / `취소율`

#### `취소율` 시트 (col 기준, 0-indexed)
| col | 내용 |
|-----|------|
| 1 | 구분 (뇌새김영어 등) |
| 2 | 그룹 (A그룹 / W그룹 / E그룹 …) |
| 3 | 센터 |
| 4 | 팀 |
| 5 | 코디명 (단순 이름, 예: `이은정`) |
| 7 | 최초 계약수 |
| 11 | 취소수 |
| 12 | 취소율 |

- 행 2~3: 시작/종료일자
- 행 5: 컬럼 헤더
- 행 6: 전체 합계
- 행 7~: 개인별 데이터

#### `명단` 시트 (col 기준, 0-indexed)
| col | 내용 |
|-----|------|
| 0 | 고유값 (담당+이름, 예: `뇌새김영어이은정`) |
| 1 | 이름 |
| 2 | 센터 |
| 3 | 팀 |
| 4 | 그룹 |
| 8 | 메신저아이디 |
| 9 | 이름 (R_센터) 형식 예) `(R_대전)이은정` |

---

### 2-3. WM 근태보고 파일 구조 (HTML 형식 .xls)
```python
tables = pd.read_html(파일경로, encoding='utf-8')
df = tables[0]
```

| 컬럼명 | 내용 |
|--------|------|
| 담당 | 담당 구분 (예: `뇌새김영어`) |
| 그룹 | 소속 그룹 (예: `A그룹`) |
| 센터 | 근무 센터 |
| 팀 | 팀명 |
| 이름 | `(R_센터)이름` 형식 |
| 입사일 | 입사일 |
| 04/01(수) ~ 04/30(목) | 날짜별 근태 상태 (출근 / 연차 / NaN) |

---

## 3. N그룹 구성원 판별 로직

```
조건: 근태보고 파일에서 아래 두 조건을 모두 만족하는 행
  - 담당 == "뇌새김영어"
  - 그룹 == "A그룹"
```

```python
def get_n_group_members(df_attendance):
    """
    근태보고 파일에서 N그룹(뇌새김영어 / A그룹) 구성원 목록 반환
    반환값: [(코디명, 센터), ...] 리스트
    """
    mask = (df_attendance['담당'] == '뇌새김영어') & (df_attendance['그룹'] == 'A그룹')
    members = df_attendance[mask]['이름'].tolist()
    return members  # 예: ['(R_대전)이은정', '(R_유성)김영미', ...]
```

---

## 4. 취소율 조회 로직

```python
def get_cancel_rate(name_with_region: str, df_cancel_명단: pd.DataFrame, df_cancel_rate: pd.DataFrame) -> float:
    """
    코디명(예: '(R_대전)이은정')으로 기준 취소율 조회
    
    1단계: 명단 시트에서 코디명 → 단순 이름 변환
    2단계: 취소율 시트에서 단순 이름으로 취소율 조회
    """
    # 명단에서 (R_센터)이름 → 단순이름 매핑
    name_map = df_cancel_명단.set_index(9)[1].to_dict()  # col9: (R_)이름, col1: 단순이름
    simple_name = name_map.get(name_with_region, None)
    
    if simple_name is None:
        return 0.0  # 또는 전체 평균 취소율
    
    # 취소율 시트 (행 7부터 데이터, col5: 코디명, col12: 취소율)
    match = df_cancel_rate[df_cancel_rate[5] == simple_name]
    if match.empty:
        return 0.0
    return float(match.iloc[0][12])
```

---

## 5. 그룹 등급 산정 로직 (핵심)

### 5-1. 기본 등급 구조
```
N_A그룹 (최상위) > N_B그룹 > N_C그룹 (최하위)
특수: H그룹 (제재 등급)
```

### 5-2. 등급 결정 기준
이전 성과 기간의 **취소율 반영 계약률** 순위로 결정:
- 상위 1/3 → N_A그룹
- 중위 1/3 → N_B그룹
- 하위 1/3 → N_C그룹

```python
def assign_groups(members_with_scores: list) -> dict:
    """
    members_with_scores: [(코디명, 취소율반영계약률), ...]
    반환: {코디명: 그룹명}
    """
    sorted_members = sorted(members_with_scores, key=lambda x: x[1], reverse=True)
    n = len(sorted_members)
    third = n // 3
    
    result = {}
    for i, (name, score) in enumerate(sorted_members):
        if i < third:
            result[name] = 'N_A그룹'
        elif i < 2 * third:
            result[name] = 'N_B그룹'
        else:
            result[name] = 'N_C그룹'
    return result
```

### 5-3. H그룹 전환 로직 ⚠️
**N_C그룹에 연속 3회** 배정된 경우 → **H그룹**으로 변경

```python
def apply_h_group_rule(history: dict, current_assignments: dict) -> dict:
    """
    history: {코디명: [이전등급1, 이전등급2, ...]} (최근 순서)
    current_assignments: {코디명: 현재배정등급}
    
    규칙: 현재 포함 직전 3회 연속 N_C그룹이면 H그룹으로 변경
    """
    result = dict(current_assignments)
    
    for name, current_group in current_assignments.items():
        if current_group == 'N_C그룹':
            past = history.get(name, [])
            # 직전 2회 모두 N_C그룹인지 확인 (현재 포함 3회)
            if len(past) >= 2 and past[0] == 'N_C그룹' and past[1] == 'N_C그룹':
                result[name] = 'H그룹'
    
    return result
```

### 5-4. H그룹 → N_C그룹 복귀 로직 ⚠️
**H그룹에서 2건 계약 달성** 시 → 다음 등급 선정 시 **N_C그룹**으로 배정

```python
def apply_h_group_recovery(history: dict, h_group_contracts: dict, current_assignments: dict) -> dict:
    """
    h_group_contracts: {코디명: H그룹 기간 동안의 계약수}
    
    규칙: H그룹인 사람이 2건 이상 계약 시 다음 기간 N_C그룹으로 배정
    """
    result = dict(current_assignments)
    
    for name in current_assignments:
        past = history.get(name, [])
        if past and past[0] == 'H그룹':  # 직전 기간이 H그룹
            contracts = h_group_contracts.get(name, 0)
            if contracts >= 2:
                result[name] = 'N_C그룹'
    
    return result
```

### 5-5. 등급 산정 전체 흐름

```
1. 근태보고 파일에서 N그룹 구성원 목록 추출
2. 이전 성과 시트들에서 각 구성원의 등급 이력 구성
3. H그룹 구성원의 이번 기간 계약수 확인
4. 기본 등급 산정 (취소율 반영 계약률 순위)
5. H그룹 복귀 로직 적용 (H그룹 2건 달성자 → N_C그룹)
6. H그룹 전환 로직 적용 (N_C그룹 3연속 → H그룹)
7. 최종 그룹 배정 결정
```

---

## 6. DB수 / 계약수 산출 로직

### DB수 (N그룹 성과 파일의 `DB수` 시트 기준)
- 해당 성과 기간 내에 해당 코디에게 배분된 DB 수
- `DB수` 시트: 분배일 기준, 제외조건 `포함` 행만 집계
- OB명 컬럼으로 코디 구분

```python
def calc_db_count(df_db: pd.DataFrame, codi_name: str, start_date, end_date) -> int:
    """
    DB수 시트에서 특정 코디의 기간 내 DB수 계산
    - 행 1부터 데이터 시작
    - col1: 제외조건 ('포함'/'제외')
    - col0: 분배일
    - col9: OB명 (예: '(R_대전)이은정')
    """
    df = df_db.iloc[1:].copy()
    df.columns = range(df.shape[1])
    mask = (
        (df[1] == '포함') &
        (df[9] == codi_name) &
        (pd.to_datetime(df[0]) >= start_date) &
        (pd.to_datetime(df[0]) <= end_date)
    )
    return mask.sum()
```

### 계약수 (N그룹 성과 파일의 `계약수` 시트 기준)
- 해당 성과 기간 내 계약이 이루어진 건수
- `계약수` 시트: `계약 포함 여부` == '포함', `계약여부` == '계약' 행 집계

```python
def calc_contract_count(df_contract: pd.DataFrame, codi_name: str, start_date, end_date) -> int:
    """
    계약수 시트에서 특정 코디의 기간 내 계약수 계산
    - col2: 계약 포함 여부 ('포함'/'제외')
    - col3: 계약여부 ('계약'/'제외')
    - col0: 분배일
    - col10: OB명
    """
    df = df_contract.iloc[1:].copy()
    df.columns = range(df.shape[1])
    mask = (
        (df[2] == '포함') &
        (df[3] == '계약') &
        (df[10] == codi_name) &
        (pd.to_datetime(df[0]) >= start_date) &
        (pd.to_datetime(df[0]) <= end_date)
    )
    return mask.sum()
```

### 취소수 (해당 기간 실취소 건수)
- `계약수` 시트: `계약 포함 여부` == '포함', `계약여부` == '계약'이지만 이미 취소된 건
- 또는 별도 취소 이벤트 추적

---

## 7. 전체 스크립트 구조

```
project/
├── generate_n_group_report.py   # 메인 실행 파일
├── config.py                    # 설정값 (파일 경로, 기간 등)
├── data_loader.py               # 파일 읽기 함수
├── cancel_rate.py               # 취소율 조회 로직
├── group_logic.py               # 그룹 등급 산정 로직 (H그룹 포함)
├── report_builder.py            # Excel 출력 로직 (기존 포맷 유지)
└── utils.py                     # 공통 유틸 (날짜 파싱 등)
```

---

## 8. 메인 실행 로직 (generate_n_group_report.py)

```python
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import copy

# ── 설정 ──────────────────────────────────────────────────────────────────
TEMPLATE_FILE   = 'N그룹_성과_26년_3월.xlsx'      # 기존 성과 파일 (템플릿)
CANCEL_FILE     = 'WM_기간별_취소율_XXX.xlsx'      # 취소율 파일
ATTENDANCE_FILE = 'WM_근태보고_그룹_YYYYMMDD.xls'  # 근태보고 파일
OUTPUT_FILE     = 'N그룹_성과_26년_4월.xlsx'        # 출력 파일명

START_DATE = datetime(2026, 4, 1)
END_DATE   = datetime(2026, 4, 6)
SHEET_NAME = '260406'  # 새로 추가할 시트명

# ── STEP 1: 파일 로드 ──────────────────────────────────────────────────────
def load_all():
    xl_perf   = pd.ExcelFile(TEMPLATE_FILE)
    df_db     = pd.read_excel(xl_perf, 'DB수',   header=None)
    df_cont   = pd.read_excel(xl_perf, '계약수', header=None)

    xl_cancel = pd.ExcelFile(CANCEL_FILE)
    df_명단   = pd.read_excel(xl_cancel, '명단',   header=None)
    df_cancel = pd.read_excel(xl_cancel, '취소율', header=None)

    tables     = pd.read_html(ATTENDANCE_FILE, encoding='utf-8')
    df_attend  = tables[0]
    
    return df_db, df_cont, df_명단, df_cancel, df_attend

# ── STEP 2: N그룹 구성원 추출 ────────────────────────────────────────────
def get_n_members(df_attend):
    mask = (df_attend['담당'] == '뇌새김영어') & (df_attend['그룹'] == 'A그룹')
    return df_attend[mask]['이름'].dropna().tolist()

# ── STEP 3: 각 구성원의 DB수/계약수 집계 ─────────────────────────────────
def calc_stats(df_db, df_cont, members, start, end):
    stats = {}
    for name in members:
        db_cnt   = calc_db_count(df_db, name, start, end)
        cont_cnt = calc_contract_count(df_cont, name, start, end)
        cancel_cnt = calc_cancel_count(df_cont, name, start, end)
        stats[name] = {
            'db': db_cnt,
            'contract': cont_cnt,
            'cancel': cancel_cnt,
        }
    return stats

# ── STEP 4: 취소율 조회 ──────────────────────────────────────────────────
def enrich_cancel_rate(stats, df_명단, df_cancel):
    # col9 = (R_)이름, col1 = 단순이름
    name_to_simple = {}
    for _, row in df_명단.iloc[1:].iterrows():
        name_to_simple[row[9]] = row[1]
    
    # col5 = 단순코디명, col12 = 취소율 (행 7부터 데이터)
    cancel_map = {}
    for _, row in df_cancel.iloc[7:].iterrows():
        cancel_map[row[5]] = row[12]
    
    for name, s in stats.items():
        simple = name_to_simple.get(name, '')
        rate = cancel_map.get(simple, 0.0)
        if isinstance(rate, str) and '%' in rate:
            rate = 0.0
        s['cancel_rate_ref']  = float(rate)
        s['adj_contract']     = s['contract'] * (1 - s['cancel_rate_ref'])
        s['contract_rate']    = s['contract'] / s['db'] if s['db'] > 0 else 0
        s['adj_rate']         = s['adj_contract'] / s['db'] if s['db'] > 0 else 0
    return stats

# ── STEP 5: 등급 산정 ────────────────────────────────────────────────────
def assign_groups_with_history(stats, grade_history, h_contracts):
    """
    grade_history: {코디명: ['N_C그룹', 'N_C그룹', ...]}  최신→과거 순
    h_contracts:   {코디명: H그룹 기간 계약수}
    """
    members_scores = [(n, s['adj_rate']) for n, s in stats.items()]
    
    # 기본 등급 산정
    sorted_m = sorted(members_scores, key=lambda x: x[1], reverse=True)
    n = len(sorted_m)
    third = n // 3
    base = {}
    for i, (name, _) in enumerate(sorted_m):
        if i < third:
            base[name] = 'N_A그룹'
        elif i < 2 * third:
            base[name] = 'N_B그룹'
        else:
            base[name] = 'N_C그룹'
    
    # H그룹 복귀 적용 (직전 H그룹 + 2건 이상)
    for name in list(base.keys()):
        hist = grade_history.get(name, [])
        if hist and hist[0] == 'H그룹':
            if h_contracts.get(name, 0) >= 2:
                base[name] = 'N_C그룹'  # H → N_C 복귀
    
    # H그룹 전환 적용 (현재 N_C + 직전 2회 연속 N_C)
    for name in list(base.keys()):
        if base[name] == 'N_C그룹':
            hist = grade_history.get(name, [])
            if len(hist) >= 2 and hist[0] == 'N_C그룹' and hist[1] == 'N_C그룹':
                base[name] = 'H그룹'
    
    return base

# ── STEP 6: Excel 시트 생성 ──────────────────────────────────────────────
def write_sheet(wb, sheet_name, start_date, end_date, stats, groups):
    """
    기존 날짜 시트(예: 260330)를 복사하여 새 시트 생성 후 데이터 기입
    """
    # 가장 최근 날짜 시트를 템플릿으로 복사
    template_sheet = wb.worksheets[-1]  # 또는 특정 시트 지정
    ws = wb.copy_worksheet(template_sheet)
    ws.title = sheet_name
    
    # 헤더 업데이트
    # 시작/종료일자 업데이트 (기존 셀 위치 유지)
    ws['H2'] = f"{start_date.strftime('%Y-%m-%d')}({['월','화','수','목','금','토','일'][start_date.weekday()]})"
    ws['H3'] = f"{end_date.strftime('%Y-%m-%d')}({['월','화','수','목','금','토','일'][end_date.weekday()]})"
    
    # 그룹별 정렬: N_A → N_B → N_C → H, 그룹 내 adj_rate 내림차순
    group_order = {'N_A그룹': 0, 'N_B그룹': 1, 'N_C그룹': 2, 'H그룹': 3}
    sorted_names = sorted(
        stats.keys(),
        key=lambda n: (group_order.get(groups.get(n, 'N_C그룹'), 2), -stats[n]['adj_rate'])
    )
    
    # 데이터 행 기입 (행 5부터, 0-indexed 기준 row=5 → openpyxl row=6)
    start_row = 6
    total_db = total_cont = total_cancel = 0
    
    for i, name in enumerate(sorted_names):
        row = start_row + i
        s = stats[name]
        g = groups.get(name, 'N_C그룹')
        
        ws.cell(row=row, column=2).value = g       # B: 그룹
        ws.cell(row=row, column=3).value = name    # C: 코디명
        ws.cell(row=row, column=4).value = s['db']
        ws.cell(row=row, column=5).value = s['contract']
        ws.cell(row=row, column=6).value = s['contract_rate']
        ws.cell(row=row, column=7).value = s['cancel']
        ws.cell(row=row, column=8).value = s['cancel'] / s['contract'] if s['contract'] > 0 else 0
        ws.cell(row=row, column=9).value = s['contract'] - s['cancel']
        ws.cell(row=row, column=10).value = (s['contract'] - s['cancel']) / s['db'] if s['db'] > 0 else 0
        ws.cell(row=row, column=11).value = s['cancel_rate_ref']
        ws.cell(row=row, column=12).value = s['adj_contract']
        ws.cell(row=row, column=13).value = s['adj_rate']
        
        total_db    += s['db']
        total_cont  += s['contract']
        total_cancel += s['cancel']
    
    # 합계행 (row 5, openpyxl row=5)
    ws.cell(row=5, column=4).value = total_db
    ws.cell(row=5, column=5).value = total_cont
    ws.cell(row=5, column=6).value = total_cont / total_db if total_db > 0 else 0
    ws.cell(row=5, column=7).value = total_cancel
    ws.cell(row=5, column=8).value = total_cancel / total_cont if total_cont > 0 else 0
    
    return wb


# ── STEP 7: 메인 실행 ────────────────────────────────────────────────────
if __name__ == '__main__':
    df_db, df_cont, df_명단, df_cancel, df_attend = load_all()
    
    members = get_n_members(df_attend)
    print(f"N그룹 구성원 수: {len(members)}")
    
    stats = calc_stats(df_db, df_cont, members, START_DATE, END_DATE)
    stats = enrich_cancel_rate(stats, df_명단, df_cancel)
    
    # 이전 등급 이력 및 H그룹 계약수는 별도 관리 (아래 참고)
    grade_history = load_grade_history(TEMPLATE_FILE)  # 구현 필요
    h_contracts   = calc_h_group_contracts(df_cont, members, grade_history)
    
    groups = assign_groups_with_history(stats, grade_history, h_contracts)
    
    # Excel 저장
    wb = load_workbook(TEMPLATE_FILE)
    wb = write_sheet(wb, SHEET_NAME, START_DATE, END_DATE, stats, groups)
    wb.save(OUTPUT_FILE)
    print(f"완료: {OUTPUT_FILE} 저장됨")
```

---

## 9. 등급 이력 관리 (`load_grade_history`)

```python
def load_grade_history(template_file: str) -> dict:
    """
    기존 성과 파일의 날짜별 시트를 역순으로 읽어
    각 코디명의 등급 이력을 구성한다.
    
    반환: {코디명: [가장최근등급, 그전등급, ...]}
    """
    xl = pd.ExcelFile(template_file)
    
    # 날짜 시트만 필터링 (6자리 숫자 시트명, 오류 시트 제외)
    date_sheets = [
        s for s in xl.sheet_names
        if s.isdigit() and len(s) == 6 and '오류' not in s
    ]
    # 최신순 정렬
    date_sheets.sort(reverse=True)
    
    history = {}
    for sheet in date_sheets:
        df = pd.read_excel(xl, sheet_name=sheet, header=None)
        # 행 5부터 데이터, col1=그룹, col2=코디명
        for _, row in df.iloc[5:].iterrows():
            name  = row[2]
            group = row[1]
            if pd.notna(name) and pd.notna(group):
                if name not in history:
                    history[name] = []
                history[name].append(group)
    
    return history
```

---

## 10. 주의사항 및 예외처리

### 취소율 0% 처리
취소율이 `'0.0%'` 문자열로 저장된 경우 → `0.0`으로 변환 필요

```python
def safe_float(val):
    if isinstance(val, str):
        val = val.replace('%', '').strip()
    try:
        return float(val)
    except:
        return 0.0
```

### 신규 구성원 처리
- 이전 시트에 없던 신규 구성원은 등급 이력 없음 → N_C그룹으로 초기 배정

### H그룹 구성원의 DB 배분
- H그룹은 일반 DB 배분에서 제외될 수 있음 → DB수가 0인 경우 처리

### 날짜 시트 이름 규칙
- 형식: `YYMMDD` (예: `260406`)
- 종료일 기준으로 시트명 생성

### 성과 시트 업데이트
- `성과` 시트는 **전체 기간 누적** 집계 시트
- 새 시트 추가 시 성과 시트도 함께 업데이트 필요

---

## 11. 실행 환경 및 의존성

```bash
# 필요 패키지
pip install pandas openpyxl lxml html5lib

# Python 버전
Python 3.9 이상 권장

# 실행
python generate_n_group_report.py
```

---

## 12. 개발 순서 (권장)

1. `data_loader.py` — 세 파일 로딩 함수 구현 및 테스트
2. `cancel_rate.py` — 취소율 조회 함수 구현 및 단위 테스트
3. `group_logic.py` — 기본 그룹 배정 → H그룹 전환 → H그룹 복귀 순서로 구현
4. `report_builder.py` — 기존 시트 복사 후 데이터 기입 로직
5. `generate_n_group_report.py` — 전체 통합 및 최종 테스트
