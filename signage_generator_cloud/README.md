# HMG 사이니지 스케줄 생성기 - Streamlit Cloud 배포용

엑셀을 직접 편집하지 않고, 웹 UI에서 날짜와 플랜을 선택해 사이니지용 스케줄 엑셀을 생성하는 앱입니다.

## 필요 파일

업로드할 기준 엑셀에는 아래 시트가 있어야 합니다.

- `스케줄` 또는 앱에서 지정한 출력 시트
- `Plan1`
- `Plan2`
- `Plan3`
- `Plan1 (주말)`
- `Plan2 (주말)`
- `Plan3 (주말)`

각 시트의 헤더에는 다음 의미의 컬럼이 있어야 합니다.

- 날짜 / 일자
- 알람 / 알림
- 시작예정 / 시작 예정 / 시작시간 / 시작 시각 / 시작
- 프로그램명 / 프로그램
- 렉처룸 / 렉쳐룸 / 강의실 / 룸

## Streamlit Community Cloud 배포 방법

1. GitHub에 새 Repository를 만듭니다.
2. 이 폴더 안의 파일을 그대로 업로드합니다.
   - `app.py`
   - `requirements.txt`
   - `.streamlit/config.toml`
   - `README.md`
3. Streamlit Community Cloud에 접속합니다.
4. `New app`을 클릭합니다.
5. GitHub Repository를 연결합니다.
6. Main file path는 `app.py`로 지정합니다.
7. Deploy를 누릅니다.

## 사용 방법

1. 배포된 Streamlit 앱 접속
2. 기준 엑셀 업로드
3. 출력 시트명 확인, 기본값은 `스케줄`
4. 날짜와 Plan1/Plan2/Plan3 입력
5. 토요일 또는 일요일이면 자동으로 주말 시트를 적용합니다.
   - Plan1 -> Plan1 (주말)
   - Plan2 -> Plan2 (주말)
   - Plan3 -> Plan3 (주말)
6. `엑셀 생성` 클릭
7. 완성 엑셀 다운로드

## 보안 메모

Streamlit Community Cloud에 배포하면 앱 코드는 온라인에 올라갑니다. 기준 엑셀 파일은 사용자가 앱 화면에서 업로드할 때 처리됩니다. 회사 내부 자료를 다룰 경우 GitHub 저장소 공개 범위와 Streamlit 접근 권한을 반드시 확인하세요.
