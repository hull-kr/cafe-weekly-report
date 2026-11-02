## 처음 받은 폴더라도 바로 실행한다

- 주간 보고 요청이 오면 설치하라는 말이 없어도 먼저 `python tools\setup.py` 로 빠진 것을 설치한다(weekly-report 스킬 0단계). 설치해도 되는지 묻지 않는다.

## 엑셀·PPT 작업 환경 (document-skills 플러그인, local scope)

- **Python 은 반드시 `.venv\Scripts\python` 을 쓴다.**
  openpyxl · pandas · markitdown · python-pptx · Pillow · defusedxml · lxml 이 `.venv` 에 있다.
- markitdown 은 `.venv\Scripts\python -X utf8 -m markitdown 파일` 로 부른다(한글 깨짐 방지).
- Node 패키지(pptxgenjs · react-icons · react · react-dom · sharp)는 이 폴더 `node_modules` 에 있다. 스크립트는 이 폴더에서 실행한다.
- **스킬의 파이썬 스크립트는 직접 부르지 말고 `tools\skill_run.py` 를 거친다.**
  스킬의 `office/soffice.py` 는 리눅스용이라 Windows 에서 `socket.AF_UNIX` 에러로 죽는다. 실행기가 이를 우회하고 LibreOffice 경로도 잡아 준다.

  ```
  .venv\Scripts\python tools\skill_run.py xlsx recalc.py 파일.xlsx
  .venv\Scripts\python tools\skill_run.py pptx thumbnail.py 덱.pptx 덱-thumbs
  .venv\Scripts\python tools\skill_run.py pptx office/soffice.py --headless --convert-to pdf 덱.pptx
  ```
- PDF → 이미지는 `pdftoppm` 을 그대로 쓴다(PATH 에 있음).

## 스킬 만들기 (skill-creator 플러그인, local scope)

- skill-creator 의 스크립트(`python -m scripts.xxx`)도 `tools\skill_run.py` 로 부른다. 작업 폴더는 이 프로젝트 그대로 둔다.

  ```
  .venv\Scripts\python tools\skill_run.py skill-creator scripts.quick_validate 스킬폴더
  .venv\Scripts\python tools\skill_run.py skill-creator scripts.package_skill 스킬폴더 [출력폴더]
  .venv\Scripts\python tools\skill_run.py skill-creator scripts.run_eval --eval-set 평가.json --skill-path 스킬폴더 --num-workers 1
  ```
- `run_eval` · `run_loop` 는 Windows 에서 파이프에 `select` 를 써서 죽는다. 실행기가 `tools\winpatch\sitecustomize.py` 를 PYTHONPATH 로 붙여 자식 프로세스까지 보정한다.
- **발동 평가는 `--num-workers 1` 로 돌린다.** 작업자가 여럿이면 같은 설명의 임시 스킬이 동시에 여러 개 보여 서로 간섭하고, 발동돼야 할 문장도 0 이 나온다(2026-11-02 확인: 병렬 0/3, 작업자 1개 2/2).

## PPT 검수는 실제 PowerPoint 로 마무리한다

- LibreOffice 그림은 PowerPoint 와 다르게 나온다. 차트 글자에서 한글과 숫자 사이를 띄우고(`943 만`), 한글을 단어 단위로 줄바꿈한다.
  PowerPoint 는 `943만` 으로 붙여 그리지만 한글을 **글자 단위로** 줄바꿈해 단어가 잘린다. 긴 한글 문장은 줄을 직접 나눈다(`\n`).
- `validate.py` 를 통과해도 PowerPoint 가 못 여는 파일이 있다. 예: 차트 XML 에 `<a:ea>` 글꼴을 덧붙인 파일.
- PowerPoint 그림 뽑기: `.venv\Scripts\python -X utf8 .claude\skills\weekly-report\scripts\ppt_render.py 덱.pptx 그림폴더`
  (예전 `tools\ppt_render.ps1` 도 동작하지만 `_검수\` 폴더 고정이라 위 것을 쓴다)

## 주간 보고는 weekly-report 스킬로 한다

- "주간 보고서 만들어 줘" 류 요청은 `.claude\skills\weekly-report` 스킬을 따른다. 합치기부터 PPT 검수까지 그 SKILL.md 가 정식 절차다.
- 결과는 `보고서\<주>\` 에 쌓인다. 주 폴더(`MMDD-MMDD`)에는 결과를 쓰지 않는다.
- `tools\build_merged.py` · `tools\deck_data.py` · `tools\build_deck.js` 와 맨 위 `통합_판매.xlsx` · `주간보고_1019-1025.pptx` 는 1019-1025 한 주를 손으로 만들 때의 처음 버전이다. 새 주에는 쓰지 않는다.
- 스킬 시험 자료와 결과는 `weekly-report-workspace\` 에 있다 (시험 폴더 `fixtures\` 는 가짜 자료).
