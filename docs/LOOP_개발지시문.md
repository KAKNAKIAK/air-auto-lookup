# 항공자동조회 LOOP 개발 지시문

작성일: 2026-06-29

## 전달 목적

이 문서는 `항공자동조회` 프로젝트를 다른 AI 개발자에게 넘겨 실제 개발을 진행시키기 위한 지시문이다.

핵심 목표는 한 번에 크게 만들지 않고, 작은 LOOP 단위로 다음 순서를 반복하는 것이다.

```text
읽기 -> 이해 -> 작은 기능 구현 -> 검증 -> 작업일지 기록 -> 다음 LOOP 제안
```

## 작업 위치

```text
C:\Users\naeiltour0001\Desktop\ERP Automatic Fare Update Project\항공자동조회
```

## 반드시 먼저 읽을 파일

아래 파일을 UTF-8로 읽고, 현재 상태를 파악한 뒤 작업을 시작한다.

```powershell
Get-Content -LiteralPath 'G:\내 드라이브\안티그래비티\AGENTS.md' -Encoding UTF8
Get-Content -LiteralPath 'C:\Users\naeiltour0001\Desktop\ERP Automatic Fare Update Project\항공자동조회\작업일지.md' -Encoding UTF8
Get-Content -LiteralPath 'C:\Users\naeiltour0001\Desktop\ERP Automatic Fare Update Project\항공자동조회\개념정리.md' -Encoding UTF8
Get-Content -LiteralPath 'C:\Users\naeiltour0001\Desktop\ERP Automatic Fare Update Project\항공자동조회\기획문서.md' -Encoding UTF8
```

코드/데이터 파일:

```text
air_auto_lookup_mvp.py
flight-master.mjs
run_mvp.bat
```

## 현재 프로젝트 상태

- 기존 ERP 요금 업로드 프로젝트와 분리된 별도 프로젝트다.
- 현재 MVP는 Python Tkinter GUI와 openpyxl 기반 엑셀 생성기로 구성되어 있다.
- GUI는 조회 시작일, 조회 종료일, 상품일수, 노선 카테고리 체크박스를 제공한다.
- 편명 단위 선택은 제거했다. 체크된 노선 카테고리에 속한 전체 편명이 엑셀 생성 대상이다.
- `항공자동조회_MVP_조회계획_*.xlsx`는 더 이상 생성하지 않는다. 조회계획은 LOG 성격이다.
- 사용자 산출물은 `항공자동조회_MVP_결과_YYYYMMDD_HHMMSS.xlsx` 형식의 노선별 결과 엑셀이다.
- 결과 엑셀은 `ICNDAD`, `ICNCXR` 같은 노선별 시트를 만들고, 각 시트에 날짜/항공편별 요금칸/최저가/최저가 항공 수식을 둔다.
- TOPAS 실제 자동입력, 원문 저장, 원문 파싱, 요금계산은 아직 붙지 않았다.

## 절대 지켜야 할 핵심 규칙

1. `flight-master.mjs`를 노선/편명 마스터의 기준으로 둔다.
2. 귀국편 시간은 마스터의 고정값으로 전체 기간을 계산하면 안 된다.
3. 같은 항공사, 같은 노선, 같은 편명이어도 시즌에 따라 귀국시간이 23시대 또는 01시대로 바뀔 수 있다.
4. 박수는 TOPAS 귀국편 결과에서 날짜별 실제 귀국편 출발시간을 파싱해서 확정한다.
5. 5일 상품이면 3박 후보일과 4박 후보일을 모두 조회 대상으로 고려한다.
6. 6일 상품이면 4박 후보일과 5박 후보일을 모두 조회 대상으로 고려한다.
7. 기존 파일 삭제, 대규모 이동, 덮어쓰기는 사용자 승인 없이 하지 않는다.
8. 작업 후 `작업일지.md`를 갱신한다.
9. 검증 없이 완료 처리하지 않는다.

## 현재 구현된 주요 함수

`air_auto_lookup_mvp.py` 기준:

| 함수 | 역할 |
|---|---|
| `load_flight_masters()` | `flight-master.mjs`에서 마스터 로드 |
| `iter_dates()` | 조회 시작일~종료일 날짜 생성 |
| `date_count()` | 조회 날짜 수 계산 |
| `topas_command()` | TOPAS AN 명령 생성 |
| `calculate_nights()` | 귀국시간 기준 박수 계산 |
| `return_candidate_dates()` | 상품일수 기준 귀국 후보일 생성 |
| `build_plan_rows()` | 내부 조회계획/log row 생성. 감지된 귀국시간 입력 가능 |
| `write_route_mvp_excel()` | 노선별 결과 엑셀 생성 |
| `AirAutoLookupApp` | Tkinter GUI |

## 개발 LOOP 규칙

각 LOOP는 아래 형식으로 진행한다.

### 1. Observe

- 관련 파일을 읽는다.
- 현재 동작과 직전 작업일지 내용을 확인한다.
- 기존 코드를 추측으로 덮어쓰지 않는다.

### 2. Plan

- 이번 LOOP에서 만들 기능을 1개로 제한한다.
- 산출물, 수정 파일, 검증 방법을 먼저 적는다.
- 범위가 크면 더 작은 LOOP로 쪼갠다.

### 3. Implement

- 기존 구조를 최대한 유지한다.
- 새 기능은 작고 검증 가능하게 만든다.
- 불필요한 리팩터링은 하지 않는다.

### 4. Verify

최소 검증:

```powershell
python -m py_compile air_auto_lookup_mvp.py
```

기능별 추가 검증:

- 날짜/박수 계산은 작은 Python 스니펫으로 결과를 출력한다.
- 엑셀 생성은 실제 xlsx를 만들고 openpyxl로 시트명/헤더/수식을 확인한다.
- TOPAS 파싱은 샘플 원문 파일을 두고 expected 결과와 비교한다.

### 5. Log

`작업일지.md`에 아래를 남긴다.

- 요청
- 반영
- 검증
- 생성 파일
- 남은 한계
- 다음 LOOP 후보

### 6. Report

사용자에게 짧게 보고한다.

- 수정 파일
- 생성 파일
- 검증 결과
- 남은 리스크
- 다음 액션

## 권장 개발 순서

### LOOP 1. 실행 로그 구조 만들기

목표:

- 엑셀이 아닌 JSONL/JSON 기반 실행 로그 구조를 만든다.

추천 산출물:

```text
output\logs\{runId}\run.json
output\logs\{runId}\events.jsonl
```

필수 필드:

| 필드 | 설명 |
|---|---|
| `runId` | 실행 ID |
| `startedAt` | 시작 시각 |
| `startDate` | 조회 시작일 |
| `endDate` | 조회 종료일 |
| `productDays` | 상품일수 |
| `selectedRoutes` | 선택 노선 |
| `status` | running, completed, failed |

검증:

- 샘플 실행 시 로그 폴더와 `run.json`이 생성되어야 한다.
- 기존 결과 엑셀 생성은 깨지면 안 된다.

### LOOP 2. TOPAS 명령 계획 JSON 생성

목표:

- 출발편 명령과 귀국 후보편 명령을 엑셀이 아닌 JSON으로 만든다.

추천 산출물:

```text
output\logs\{runId}\command-plan.json
```

필수 구조:

```json
{
  "departureCommands": [],
  "returnCandidateCommands": []
}
```

귀국 후보 명령은 5일 상품 기준으로 3박/4박 후보를 모두 포함해야 한다.

검증 예시:

- 2026-07-01 출발, `ICNDAD_ZE593_594`, 5일 상품:
  - 출발편: `AN01JULICNDAD/AZE593`
  - 3박 귀국 후보: `AN04JULDADICN/AZE594`
  - 4박 귀국 후보: `AN05JULDADICN/AZE594`

### LOOP 3. TOPAS 원문 저장 폴더 구조 만들기

목표:

- 실제 TOPAS 실행 전이라도 원문 저장 경로 규칙을 먼저 만든다.

추천 구조:

```text
output\raw\{runId}\{routeKey}\departure\
output\raw\{runId}\{routeKey}\return\
```

파일명 후보:

```text
2026-07-01_departure.txt
2026-07-01_return_3n.txt
2026-07-01_return_4n.txt
```

검증:

- dry-run에서 빈 파일을 만들지 않는다.
- 경로 문자열만 계획 JSON에 기록한다.

### LOOP 4. TOPAS 귀국편 시간 파서 만들기

목표:

- TOPAS 귀국편 원문에서 지정 귀국편의 실제 출발시간을 파싱한다.

입력:

- TOPAS 원문 텍스트
- 항공사 코드
- 귀국편명

출력:

```json
{
  "flight": "ZE594",
  "detectedReturnTime": "00:30",
  "status": "detected"
}
```

주의:

- 원문 포맷이 불확실하면 샘플 원문을 먼저 수집한다.
- 파서가 확신하지 못하면 추측하지 말고 `parse_failed`로 남긴다.
- 여러 후보가 동시에 잡히면 `ambiguous`로 남긴다.

검증:

- 샘플 원문 파일 2개 이상을 둔다.
- 23시대, 01시대 케이스를 각각 테스트한다.

### LOOP 5. 날짜별 박수 확정기 만들기

목표:

- 귀국 후보 결과와 파싱된 시간을 기준으로 출발일별 박수를 확정한다.

규칙:

```text
18:00~23:59 = 상품일수 - 2박
00:00~05:59 = 상품일수 - 1박
그 외 = 시간확인필요
```

출력 필드:

| 필드 | 설명 |
|---|---|
| `baseDepartureDate` | 상품 출발일 |
| `routeKey` | 마스터 키 |
| `detectedReturnDate` | 확정 귀국일 |
| `detectedReturnTime` | 파싱된 귀국시간 |
| `nights` | 확정 박수 |
| `status` | confirmed, no_flight, parse_failed, ambiguous |

검증:

- 23:30은 3박.
- 00:30은 4박.
- 5일 상품의 후보일은 3박/4박 두 개.
- 6일 상품의 후보일은 4박/5박 두 개.

### LOOP 6. 결과 엑셀에 실제 요금/마감 입력 연결

목표:

- 현재 빈칸인 항공편별 요금 칸에 계산 결과를 채운다.

현재 엑셀 규칙:

- 숫자 요금이 있으면 숫자로 입력한다.
- 마감이면 `마감` 텍스트를 입력한다.
- 오른쪽 `최저가`, `최저가 항공`은 기존 수식으로 자동 계산한다.

주의:

- `OUTPUT_MVP.xlsx`의 모양을 유지한다.
- `조회계획` 엑셀을 다시 만들지 않는다.

## 최종 MVP 완료 기준

MVP는 아래가 가능하면 완료로 본다.

1. GUI에서 조회 시작일/종료일과 노선을 선택한다.
2. 선택 노선 기준으로 출발편 TOPAS 명령 계획을 만든다.
3. 상품일수 기준 귀국 후보 명령 계획을 만든다.
4. TOPAS 원문을 저장한다.
5. 귀국편 원문에서 날짜별 실제 출발시간을 파싱한다.
6. 날짜별 박수를 확정한다.
7. 요금 또는 마감 상태를 계산한다.
8. `OUTPUT_MVP.xlsx` 형태의 노선별 결과 엑셀을 생성한다.
9. 실행 로그가 남아 중단/재개 또는 오류 추적이 가능하다.

## 다른 AI에게 줄 첫 작업 지시 예시

아래 문장을 그대로 전달해도 된다.

```text
C:\Users\naeiltour0001\Desktop\ERP Automatic Fare Update Project\항공자동조회 프로젝트를 이어서 개발해 주세요.

먼저 G:\내 드라이브\안티그래비티\AGENTS.md, 작업일지.md, 개념정리.md, 기획문서.md를 UTF-8로 읽고 현재 상태를 파악하세요.

이번 LOOP에서는 실제 TOPAS 실행까지 가지 말고, output\logs\{runId}\run.json 및 events.jsonl 실행 로그 구조를 먼저 구현하세요.

기존 결과 엑셀 생성 기능은 깨뜨리지 마세요. 조회계획 엑셀은 다시 만들지 마세요.

작업 후 python -m py_compile air_auto_lookup_mvp.py를 실행하고, 작은 샘플 실행으로 로그 파일 생성 여부를 검증하세요.

마지막에 작업일지.md에 요청/반영/검증/남은 한계/다음 LOOP 후보를 기록하세요.
```
