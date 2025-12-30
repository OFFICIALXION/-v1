# 교사 시간표 패턴 검사기

엑셀 시간표 파일(.xlsx)을 입력받아 교사별 패턴을 자동 탐지하고 한국어 문장으로 알려주는 GUI 프로그램입니다.

## 설치

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 사용법

```bash
python timetable_checker.py
```

실행하면 GUI가 열리며, 파일 선택 후 검사 실행 버튼으로 결과를 확인합니다.

자체 테스트:

```bash
python timetable_checker.py --self-test
```

## 출력 예시

```
=== 심가영 선생님 ===
- 이 시간표는 심가영 선생님이 수요일에 1~4교시 연속 1학년 1반(101)입니다.
- 이 시간표는 심가영 선생님이 5일 중 3일 이상 (1,4,5,7)교시에 수업이 있는 시간표입니다. (해당 요일: 월,수,금)
```

경고가 없으면:

```
문제 패턴이 발견되지 않았습니다.
```

## 입력 엑셀 형식 요약

- 시트 이름: 기본 `주간시간표` (없으면 첫 번째 시트 사용)
- 1행: 제목(무시)
- 2행: A2=교사, B2부터 요일 블록 시작(요일 표기는 병합셀 가능)
- 3행: 요일마다 1~7 교시 숫자 반복
- 4행부터: A열 교사명, 각 셀은 `학급코드 + 줄바꿈 + 과목` 형태

## PyInstaller로 실행 파일 만들기

macOS에서 `.app` 번들을 만들려면 `--windowed` 옵션을 사용합니다.

```bash
pip install pyinstaller
pyinstaller --windowed --name "시간표점검" timetable_checker.py
```

생성된 실행 파일은 `dist/시간표점검.app`에 위치합니다.

단일 파일이 필요하면:

```bash
pyinstaller --onefile --windowed --name "시간표점검" timetable_checker.py
```

## 주의사항

- 엑셀 내 학급코드가 첫 줄에 숫자로 시작해야 인식됩니다.
- 병합 셀은 `2행`의 요일명이 표시된 첫 칸을 기준으로 7열을 해당 요일로 해석합니다.
