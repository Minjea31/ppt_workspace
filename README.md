# PPT Workspace

개인 템플릿을 재사용해 여러 주제의 PowerPoint를 만드는 작업 공간입니다.
특정 과제, 실험 방법, 수치 표시 방식이나 슬라이드 순서에 종속되지 않습니다.

제작 규칙은 [AGENTS.md](AGENTS.md) 한 곳에서 관리합니다. 이 README는 폴더와 도구 사용법을 설명합니다.

## 폴더

```text
templet/                 재사용할 디자인과 레이아웃
  base.pptx              8가지 모양을 모은 편집 가능한 템플릿
  base.pdf               템플릿 미리보기
  layouts/               모양별 PPTX
  previews/              모양별 PNG
  style.json             색상, 글꼴, 크기와 레이아웃 목록
inputs/<작업명>/         사용자 자료와 brief.md
ppt_info/<작업명>/       제작·분석 코드, 렌더링, 검수 기록
output/<작업명>/         최종 PPTX와 PDF
scripts/                 새 작업 생성, PowerPoint 렌더링, 기본 검증
AGENTS.md                이 repo에서 작업하는 에이전트의 공통 지침
```

`templet`은 요청한 폴더 이름을 그대로 사용합니다.

## 새 PPT 만들기

Python 3, Windows용 Microsoft PowerPoint가 필요합니다. Python 스크립트는 외부 패키지를 요구하지 않습니다.

```powershell
python scripts/new_project.py "프로젝트_소개" --title "프로젝트 소개"
```

`YYYYMMDD_프로젝트_소개` 이름으로 세 폴더가 생성됩니다. `inputs/`의 `brief.md`에 목적, 청중, 핵심 내용, 자료 위치를 적고 자료를 넣습니다. `output/`에는 기본 템플릿을 복사한 PPTX가 만들어집니다. 필요한 모양을 복제해 작성하고 쓰지 않는 모양은 삭제합니다.

Codex에게는 예를 들어 이렇게 요청하면 됩니다.

> inputs/20261012_프로젝트_소개의 자료로 발표용 PPT를 만들어줘. 이 repo의 템플릿을 사용하고 output/20261012_프로젝트_소개에 PPTX와 PDF를 저장해줘.

기존 PPT를 수정할 때는 수정한 파일의 위치를 알려주세요. 저장된 최신 파일을 기준으로 작업하며, 요청 없이 원본을 덮어쓰지 않습니다.

## 렌더링과 검수

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/render_presentation.ps1 -PresentationPath "output/20261012_프로젝트_소개/프로젝트_소개.pptx"
python scripts/check_presentation.py "output/20261012_프로젝트_소개/프로젝트_소개.pptx" --audit "ppt_info/20261012_프로젝트_소개/text_geometry.json"
```

렌더링 스크립트는 PPTX를 읽기 전용으로 열어 같은 출력 폴더에 PDF를 만들고, `ppt_info/<작업명>/rendered/`에 PNG를 저장합니다. 이미 PDF가 있으면 기본적으로 덮어쓰지 않습니다. 재검수할 때는 `-Overwrite`를 명시합니다.

검증 스크립트는 파일 구조, 화면 밖 객체, 줄 간격, 기록된 텍스트 잘림을 확인합니다. **자동 검증을 통과해도 각 PNG를 직접 확인해야 합니다.** 이미지 속 글자, 의미와 수치의 정확성, 의도된 겹침, 실제 가독성은 시각적으로 검수합니다.

템플릿의 CBNU 표시도 편집 가능합니다. 발표 주제에 따라 로고나 작성자 표기를 바꿔 달라고 요청할 수 있습니다.

## Git

이 폴더는 로컬 Git repo입니다. 원격 저장소는 연결하지 않았습니다. 템플릿과 공통 코드·지침은 버전 관리하고, 자료·작업 코드·결과물은 기본적으로 `.gitignore`에서 제외합니다. 특정 작업을 버전 관리하려면 제외 규칙을 수정하세요.
