# md2pptx

확장 마크다운(Extended MD)에서 PowerPoint 프레젠테이션을 빌드하는 CLI 도구.

템플릿 PPTX 슬라이드를 복제하고, `@필드` 텍스트를 교체하고, 슬라이드 노트를 추가한 후 하나의 PPTX로 병합합니다.

## 요구사항

- Windows 11 + Microsoft PowerPoint 설치
- Python 3.10+

## 설치

```bash
pip install -r requirements.txt
```

## 사용법

```bash
python -m md2pptx <body.md> -t <templates_dir> [-o output.pptx] [옵션]
```

### 필수 인자

| 인자 | 설명 |
|------|------|
| `body.md` | 확장 마크다운 본문 파일 (`proposal-body-partN.md`) |
| `-t`, `--templates` | `S*.pptx` + `slide_index.json`이 있는 템플릿 디렉토리 |

### 옵션

| 옵션 | 기본값 | 설명 |
|------|--------|------|
| `-o`, `--output` | `output/result.pptx` | 최종 PPTX 출력 경로 |
| `--slides-dir` | `output/slides/` | 개별 슬라이드 PPTX 저장 디렉토리 |
| `--batch-size` | 25 | 병합 배치 크기 |
| `--keep-slides` | | 병합 후 개별 파일 유지 |
| `--no-merge` | | 개별 슬라이드만 생성, 병합 안 함 |
| `--no-notes` | | 슬라이드 노트 작성 안 함 |
| `--continue-on-error` | | 실패 슬라이드 건너뛰고 계속 |
| `-v`, `--verbose` | | 상세 출력 |
| `-q`, `--quiet` | | 최소 출력 |

### 예시

```bash
# 기본 사용
python -m md2pptx proposal-body-part3.md -t ./templates/slides -o output/part3.pptx

# 실패 건너뛰기 + 상세 출력
python -m md2pptx body.md -t ./templates/slides --continue-on-error -v

# 개별 슬라이드만 생성 (병합 안 함)
python -m md2pptx body.md -t ./templates/slides --no-merge --keep-slides
```

## 확장 마크다운 포맷

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
@카드1_제목: 카드 제목
@카드1_내용: 카드 본문 내용
@note: 슬라이드 노트에 들어갈 텍스트 (출처, 주석 등)
```

## 터미널 출력

```
Parsing body.md... 42 slides found

Building slides ━━━━━━━━━━━━━━━━━━━━ 35/42  [S3025] T1
  ✓ 35 ok  ✗ 1 failed  / 42 total

Merging slides ━━━━━━━━━━━━━━━━━━━━━ 41/41

╭──────────── md2pptx ────────────╮
│ ✓ Build complete                │
│ Slides: 41/42 (1 failed)       │
│ Output: output/part3.pptx      │
│ Size:   12.4 MB                │
│ Time:   2m 34s                 │
╰─────────────────────────────────╯
```

## 라이선스

MIT
