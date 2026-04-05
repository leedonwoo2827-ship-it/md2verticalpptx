# md2pptx

확장 마크다운(Extended Markdown)으로 PowerPoint 프레젠테이션을 자동 생성하는 CLI 도구.

템플릿 PPTX 슬라이드를 복제하고, `@필드` 텍스트를 교체하고, 슬라이드 노트를 추가한 뒤 하나의 PPTX 파일로 병합합니다.

## 주요 기능

- **마크다운 기반 슬라이드 작성** — `@field: value` 문법으로 슬라이드 내용을 정의
- **템플릿 기반 빌드** — 기존 PPTX 템플릿의 디자인을 그대로 유지하며 텍스트만 교체
- **자동 병합** — 개별 슬라이드를 하나의 PPTX로 자동 결합 (2단계 배치 병합)
- **카드/테이블 지원** — 카드형 레이아웃, 데이터 테이블 자동 교체
- **슬라이드 노트** — `@note` 필드로 발표자 노트 자동 삽입
- **미교체 표시** — 교체되지 않은 도형에 `★미교체★` 마커 자동 표시
- **진행률 표시** — Rich 기반 프로그레스 바와 빌드 요약 패널

## 시스템 요구사항

| 항목 | 요구사항 |
|------|----------|
| OS | **Windows 10/11** (COM 자동화 사용) |
| PowerPoint | Microsoft PowerPoint 설치 필수 |
| Python | 3.10 이상 |

> **참고**: 이 도구는 PowerPoint COM 자동화(`comtypes`)를 사용합니다.
> Linux/macOS에서는 Python 환경 설정만 가능하며, 슬라이드 빌드 기능은 Windows에서만 동작합니다.

## 설치

### Windows

```cmd
git clone https://github.com/<your-username>/md2pptx.git
cd md2pptx
install.bat
```

### Linux / macOS

```bash
git clone https://github.com/<your-username>/md2pptx.git
cd md2pptx
chmod +x install.sh run.sh
./install.sh
```

### 수동 설치

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# Linux:   source .venv/bin/activate
pip install -r requirements.txt
pip install -e .
```

## 사용법

### Windows

```cmd
run.bat <body.md> -t <templates_dir> [-o output.pptx]
```

### Linux / macOS

```bash
./run.sh <body.md> -t <templates_dir> [-o output.pptx]
```

### 직접 실행

```bash
python -m md2pptx <body.md> -t <templates_dir> [-o output.pptx]
```

## CLI 옵션

| 옵션 | 기본값 | 설명 |
|------|--------|------|
| `body.md` | (필수) | 확장 마크다운 본문 파일 |
| `-t`, `--templates` | (필수) | `S*.pptx` + `slide_index.json`이 있는 템플릿 디렉토리 |
| `-o`, `--output` | `output/result.pptx` | 최종 PPTX 출력 경로 |
| `--slides-dir` | `output/slides/` | 개별 슬라이드 PPTX 저장 디렉토리 |
| `--batch-size` | 25 | 병합 배치 크기 |
| `--keep-slides` | | 병합 후 개별 슬라이드 파일 유지 |
| `--no-merge` | | 개별 슬라이드만 생성, 병합하지 않음 |
| `--no-notes` | | 슬라이드 노트 작성 안 함 |
| `--continue-on-error` | | 실패한 슬라이드를 건너뛰고 계속 진행 |
| `-v`, `--verbose` | | 상세 출력 |
| `-q`, `--quiet` | | 최소 출력 |

## 실행 예시

### 실제 프로젝트 실행 (세로형 PPTX 생성)

작업 폴더 구조가 아래와 같을 때:

```
260405-1/
├── proposal-body-part3.md      # 입력 마크다운
├── slide_index_part3.json      # 슬라이드 인덱스
├── templates/
│   └── slides/                 # S*.pptx 템플릿 (175개)
│       ├── S2001.pptx
│       ├── S2002.pptx
│       └── ...
└── output/                     # 결과물 저장
```

**Windows (CMD):**

```cmd
cd C:\Users\leedonwoo\Documents\pro2ppt\260405-1

:: 1. slide_index.json을 템플릿 폴더에 복사
copy slide_index_part3.json templates\slides\slide_index.json

:: 2. md2pptx 실행
python -m md2pptx proposal-body-part3.md -t templates\slides -o output\part3-사업관리부문.pptx --continue-on-error -v
```

**Windows (PowerShell):**

```powershell
cd C:\Users\leedonwoo\Documents\pro2ppt\260405-1

# 1. slide_index.json을 템플릿 폴더에 복사
Copy-Item slide_index_part3.json templates\slides\slide_index.json

# 2. md2pptx 실행
python -m md2pptx proposal-body-part3.md -t templates\slides -o output\part3-사업관리부문.pptx --continue-on-error -v
```

> **참고**: `slide_index.json`은 `-t` 템플릿 디렉토리 안에 있어야 합니다.
> 루트에 별도 파일(`slide_index_part3.json`)이 있는 경우, 위처럼 복사 후 실행하세요.

### 기본 사용법

```bash
# 기본 사용
python -m md2pptx proposal-body-part2.md -t ./templates/slides -o output/part2.pptx

# 에러 무시 + 상세 출력
python -m md2pptx proposal-body-part3.md -t ./templates/slides --continue-on-error -v

# 개별 슬라이드만 생성 (병합 안 함)
python -m md2pptx body.md -t ./templates/slides --no-merge --keep-slides
```

## 프로젝트 구조

```
md2pptx/
├── md2pptx/                # 핵심 패키지
│   ├── __init__.py
│   ├── __main__.py         # python -m md2pptx 진입점
│   ├── cli.py              # CLI 인터페이스 (argparse + Rich)
│   ├── parser.py           # 확장 마크다운 파서
│   ├── models.py           # 데이터 클래스 (SlideData, SlideResult, BuildSummary)
│   └── builder.py          # PowerPoint COM 빌더 + 병합
├── install.bat             # Windows 설치 스크립트
├── install.sh              # Linux/macOS 설치 스크립트
├── run.bat                 # Windows 실행 스크립트
├── run.sh                  # Linux/macOS 실행 스크립트
├── requirements.txt
├── pyproject.toml
└── README.md
```

## 확장 마크다운 포맷

### 기본 구조

```markdown
---config
reference_pptx: templates/placeholder.pptx
---

---slide
template: T1
ref_slide: 3005
---
@governing_message: 핵심 메시지 텍스트
@breadcrumb: III. 사업관리 > 1. 투입인력
@content_1: 첫 번째 콘텐츠 영역 텍스트
@note: 발표자 노트 (출처, 주석 등)
```

### 슬라이드 헤더 필드

| 필드 | 설명 |
|------|------|
| `template` | 템플릿 유형 (T1, T2, ...) |
| `ref_slide` | 참조 슬라이드 번호 (S3005.pptx → `3005`) |

### 본문 @필드

| 필드 | 설명 |
|------|------|
| `@governing_message` | 부제목 / 핵심 메시지 |
| `@breadcrumb` | 상단 경로 (예: III. 사업관리 > 1. 인력) |
| `@section_title` | 섹션 제목 |
| `@content_N` | N번째 콘텐츠 영역 (content_1, content_2, ...) |
| `@heading_N` | N번째 소제목 |
| `@label_N` | N번째 라벨 |
| `@카드N_제목` | N번째 카드 테이블 제목 |
| `@카드N_내용` | N번째 카드 테이블 내용 |
| `@note` | 슬라이드 노트 (발표에서 안 보임) |

### 마크다운 테이블

본문에 마크다운 테이블을 포함하면 슬라이드의 데이터 테이블에 자동 매핑됩니다.

```markdown
| 구분 | 1차년도 | 2차년도 | 3차년도 |
|------|---------|---------|---------|
| 예산 | 100 | 200 | 300 |
| 인력 | 5 | 8 | 10 |
```

## 템플릿 디렉토리 구조

```
templates/slides/
├── slide_index.json    # 슬라이드 메타데이터 (shape role 매핑)
├── S2001.pptx          # 템플릿 슬라이드 파일들
├── S2002.pptx
├── S3001.pptx
└── ...
```

### slide_index.json 형식

```json
{
  "slides": [
    {
      "slide_number": 3005,
      "shapes": [
        {"name": "Shape1", "role": "governing_message", "text": "..."},
        {"name": "Shape2", "role": "content_box", "text": "..."}
      ],
      "role_map": {
        "governing_message": [0],
        "content_box": [1],
        "card_table": [2, 3]
      }
    }
  ]
}
```

## 터미널 출력 예시

```
Parsing proposal-body-part3.md... 42 slides found

Building slides ━━━━━━━━━━━━━━━━━━━━ 35/42  [S3025] T1
  ✓ 35 succeeded  ✗ 1 failed  / 42 total

Merging slides ━━━━━━━━━━━━━━━━━━━━━ 41/41

╭──────────── md2pptx ────────────╮
│ ✓ Build complete                │
│ Slides: 41/42 (1 failed)       │
│ Output: output/part3.pptx      │
│ Size:   12.4 MB                │
│ Time:   2m 34s                 │
╰─────────────────────────────────╯
```

## 파이프라인 흐름

```
[proposal-body.md]       확장 마크다운 입력
        │
        ▼
   parser.py             @필드 파싱 → SlideData 리스트
        │
        ▼
   builder.py            템플릿 복사 → COM 텍스트 교체 → 개별 PPTX
        │
        ▼
   builder.py            2단계 배치 병합 → 최종 PPTX
        │
        ▼
  [output/result.pptx]   완성된 프레젠테이션
```

## 의존성

| 패키지 | 용도 |
|--------|------|
| `comtypes` | PowerPoint COM 자동화 (Windows 전용) |
| `python-pptx` | PPTX 파일 검증 |
| `rich` | 터미널 프로그레스 바 + 패널 출력 |

## 라이선스

MIT
