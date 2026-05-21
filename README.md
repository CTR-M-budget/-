# CTR 전결검색

엑셀(`data/*.xlsx`)을 푸시하면 비밀번호 보호된 검색 HTML(`CTR_전결검색.html`)이 자동으로 갱신되는 저장소입니다.

## 동작 방식

```
data/source.xlsx   ──push──▶  GitHub Actions  ──▶  CTR_전결검색.html (AES 암호화)
```

1. `data/` 폴더의 `.xlsx` 파일이 변경되면 `.github/workflows/build.yml` 이 자동 실행됩니다.
2. Actions가 `build.py` 를 돌려 엑셀을 읽고 검색 UI HTML 을 만든 뒤, `PAGE_PASSWORD` Secret 으로 AES-256-CBC 암호화합니다.
3. 결과 `CTR_전결검색.html` 이 같은 저장소에 자동 커밋됩니다.
4. (GitHub Pages를 켜둔 경우) 자동 배포되어 즉시 반영됩니다.

## 최초 1회 셋업

### 1) Repository Secret 등록 (필수)

GitHub 저장소 → **Settings** → **Secrets and variables** → **Actions** → **New repository secret**

| Name | Value |
|------|-------|
| `PAGE_PASSWORD` | 검색 페이지에 입력하는 비밀번호 (기존 사용 중인 비밀번호와 동일하게) |

비밀번호를 바꾸려면 이 Secret 값만 바꾸고 워크플로우를 다시 돌리면 됩니다. 변경 후에는 페이지에 접속하는 사람에게도 새 비밀번호를 공유해야 합니다.

### 2) (선택) GitHub Pages 활성화

저장소 → **Settings** → **Pages** → Source: `Deploy from a branch` → Branch: `main` / root

활성화 후 약 1분 뒤 `https://<organization>.github.io/<repo>/CTR_전결검색.html` 에서 접근 가능합니다.

## 평소 사용법 — 엑셀 갱신하기

방법은 3가지가 있습니다. 본인에게 편한 것을 쓰세요.

### A. GitHub 웹에서 드래그&드롭 (가장 쉬움)

1. 저장소 → `data` 폴더 진입
2. 기존 `source.xlsx` 클릭 → 우측 상단 연필(✏️) 옆 ⋮ → **Delete file** → Commit
3. 같은 폴더에 새 엑셀을 **드래그&드롭** → 파일명을 `source.xlsx` 로 맞춤 → Commit
4. Actions 탭에서 빌드가 도는지 확인 (보통 30~60초)
5. 끝. 페이지에 새 데이터가 반영됩니다.

> 한 번에 덮어쓰고 싶으면: 기존 `source.xlsx` 를 열어서 "Replace file" 옵션을 쓰면 됩니다.

### B. GitHub Desktop / VS Code 로 로컬 작업

1. 로컬 폴더의 `data/source.xlsx` 를 새 파일로 교체
2. Commit & Push
3. 이후 동작은 A와 동일

### C. 수동 트리거

Actions → "Build CTR 전결검색 HTML" → **Run workflow** 버튼.
엑셀은 그대로 두고 빌드만 다시 돌리고 싶을 때 사용.

## 폴더 구조

```
.
├── build.py                       # 엑셀 → 암호화 HTML 빌드 스크립트
├── template/
│   └── wrapper.html               # 비밀번호 입력 + 복호화 로직 (수정 금지)
├── data/
│   └── source.xlsx                # ← 여기 엑셀만 갈아끼우면 됩니다
├── CTR_전결검색.html              # 빌드 결과 (자동 생성, 직접 편집 금지)
├── .github/workflows/build.yml    # GitHub Actions
├── .gitignore
└── README.md
```

## 엑셀 작성 규칙

- 첫 시트만 사용합니다. 시트 이름은 자유.
- **1행 = 헤더**, 2행부터 데이터.
- 빈 행은 자동으로 무시됩니다.
- 셀 값은 모두 텍스트로 변환되어 표시됩니다. (숫자 포맷이 중요하면 엑셀에서 "텍스트로 저장")
- 헤더 개수만큼 컬럼이 자동 생성됩니다. 컬럼 추가/삭제 자유.

## 로컬에서 미리 빌드해 보기

```bash
pip install cryptography openpyxl
PASSWORD="여기에비밀번호" python build.py
# → CTR_전결검색.html 생성, 브라우저로 열어 확인
```

## 보안 메모

- `PAGE_PASSWORD` 는 GitHub Secrets 에만 저장되며, 빌드 로그·커밋·HTML 어디에도 평문으로 노출되지 않습니다.
- 암호화: `PBKDF2-SHA256` (600,000 iterations) → `AES-256-CBC` (Web Crypto 표준 알고리즘). 클라이언트는 브라우저의 `crypto.subtle` 만 사용하며 외부 라이브러리 없음.
- 비밀번호 분실 시 복구는 불가능합니다(설계상). 새 비밀번호로 다시 빌드해야 합니다.

## 트러블슈팅

| 증상 | 원인/해결 |
|------|----------|
| Actions에서 `Repository secret PAGE_PASSWORD 가 설정되어 있지 않습니다` | Settings → Secrets 에 `PAGE_PASSWORD` 등록 |
| Actions는 성공인데 페이지에서 "비밀번호가 올바르지 않습니다" | Secret 값과 페이지 입력값이 일치하는지 확인 |
| `data/ 폴더에 .xlsx 파일이 없습니다` | `data/source.xlsx` 가 푸시되어 있는지 확인 |
| 헤더가 깨져 보임 | 엑셀을 `.xlsx` (Excel 2007 이후 포맷)로 저장. `.xls` 는 지원 안 함 |
