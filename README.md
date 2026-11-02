# cafe-weekly-report

카페 지점별 주간 판매 엑셀을 하나로 합치고, 그 숫자로 월요일 회의용 6장짜리 PowerPoint 를 만드는 **Claude Code 스킬**입니다.

Claude Code 에 "주간 보고서 만들어 줘" 한마디만 하면 끝까지 만듭니다. 지점마다 엑셀 모양이 달라도 됩니다. 중복 줄, 환불, 휴무, 단가 오기 같은 이상한 데이터가 있어도 중간에 묻지 않고 정해진 원칙대로 처리합니다. 처리한 내용은 모두 **사후 점검 사항**으로 남깁니다.

> 예제 브랜드 "도담커피"와 저장소에 들어 있는 판매 자료는 모두 시험용 가짜 자료입니다.

## 무엇이 나오나

`보고서\<주>\` 폴더에 아래 파일이 생깁니다.

| 파일 | 내용 |
|---|---|
| `통합_판매_<주>.xlsx` | 지점 자료를 합친 `통합_판매` 시트(매출은 수량×단가 수식). 판단 내역과 원본 대비 검산을 담은 `사후점검` 시트 |
| `주간보고_<주>.pptx` | 6장. 표지, 한눈에 보기, 지점별, 메뉴별, 요일별, 사후 점검 사항 |
| `점검사항.json`, `장표데이터.json` | 중간 자료 |

지난주 자료가 있으면 전주 대비도 자동으로 들어갑니다.

## 바로 써 보기 (Windows)

블로그 [엑셀 주간 보고 3](https://claudecodelab.tistory.com/15)과 같은 가짜 자료 두 주치(`1019-1025`, `1026-1101`)가 들어 있습니다. 받자마자 해 볼 수 있습니다.

미리 깔려 있어야 하는 것은 **Claude Code · Python · Node.js** 세 가지입니다. 없으면 [CCKit](https://claudecodelab.tistory.com/2)으로 한 번에 깝니다.

1. 이 페이지 위쪽 초록색 **Code** 단추를 누르고 **Download ZIP** 을 누릅니다.
2. 받은 zip 을 `D:\` 에 풉니다.
3. 풀린 폴더에서 Claude Code 를 열고 이렇게 말합니다.

   ```
   주간 보고서 만들어 줘
   ```

처음 한 번은 스킬이 필요한 것(파이썬 · 노드 패키지, LibreOffice, document-skills 플러그인)을 스스로 설치하느라 몇 분 더 걸립니다. 따로 설치하라고 시킬 필요가 없습니다. 끝나면 `보고서\1026-1101\` 에 통합 엑셀과 6장짜리 PPT 가 생깁니다. 지난주(`1019-1025`) 대비 숫자도 들어갑니다.

다음 주부터는 새 주 폴더(예: `1102-1108`)에 지점 엑셀을 넣고 같은 말을 하면 됩니다.

---

아래는 직접 설치하거나 자세히 알고 싶을 때 읽습니다.

## 필요한 것

| 항목 | 용도 | 비고 |
|---|---|---|
| Windows 10/11 | 실행 환경 | 실행기와 검수 스크립트가 Windows 기준입니다 |
| [Claude Code](https://claude.com/claude-code) | 스킬 실행 | |
| Python 3.12 이상 | 합치기·장표 데이터 | 3.14 에서 확인했습니다 |
| Node.js 18 이상 | PPT 만들기 (pptxgenjs) | 24 에서 확인했습니다 |
| LibreOffice | 엑셀 수식 다시 계산, PPT 검사 | `C:\Program Files\LibreOffice` 에 설치 |
| Microsoft PowerPoint | 장표 그림 검수 | 없으면 LibreOffice 그림으로 대신 봅니다 |
| Poppler (`pdftoppm`) | PDF 를 그림으로 바꾸기 | PowerPoint 가 없을 때만 필요. PATH 에 둡니다 |

## 직접 설치하려면

스킬이 알아서 하는 설치를 손으로 하려면 `python tools\setup.py` 를 실행하거나 아래처럼 합니다.

```powershell
git clone https://github.com/hull-kr/cafe-weekly-report.git
cd cafe-weekly-report

# 파이썬 패키지 - 반드시 .venv 안에 설치합니다 (스킬이 .venv\Scripts\python 을 부릅니다)
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt

# Node 패키지
npm install
```

Claude Code 를 이 폴더에서 열고 플러그인 두 개를 설치합니다. `document-skills` 는 엑셀 수식 계산과 PPT 검사에 쓰고, `skill-creator` 는 스킬을 평가하거나 고칠 때만 씁니다.

```
/plugin marketplace add anthropics/skills
/plugin install document-skills@anthropic-agent-skills
/plugin install skill-creator@claude-plugins-official
```

설치한 뒤 Claude Code 를 다시 시작합니다.

## 따라 하기 1 - Claude 에게 맡기기

저장소에는 가짜 판매 자료 두 주치가 들어 있습니다.

- `weekly-report-workspace\fixtures\1026-1101\` - 이상 데이터가 섞인 5개 지점 자료와 메모 파일
- `weekly-report-workspace\fixtures\1102-1108\` - 영문 파일 이름(`gangnam_...`), 모양이 특이한 파일, 빠진 지점이 있는 자료

Claude Code 에 이렇게 말합니다.

```
주간 보고서 만들어 줘. 이번 주 폴더는 weekly-report-workspace/fixtures/1026-1101 이야.
```

몇 분 뒤 `보고서\1026-1101\` 에 엑셀과 PPT 가 생기고, Claude 가 총매출, 1위 지점·메뉴, 사후 점검 사항을 짧게 보고합니다. 이 주는 지난주 자료가 없어서 전주 대비는 생략됩니다.

두 번째 주도 해 봅니다.

```
이번 주 자료 넣었어. 회의 자료 뽑아 줘. 폴더는 weekly-report-workspace/fixtures/1102-1108
```

이 주에서는 스킬이 다음 일을 스스로 처리합니다.

- 스크립트가 읽지 못한 성수점 파일을 Claude 가 직접 열어 표준 CSV 로 옮깁니다
- `gangnam_weekly_1102.xlsx` 를 강남점으로 판단합니다
- 해운대 파일이 없다는 사실을 사후 점검에 남깁니다
- 앞의 `1026-1101` 원본을 합쳐 전주 대비를 넣습니다

"이번 주 보고서", "새 주 폴더 넣었어", "지점 판매 합쳐서 PPT까지"처럼 다르게 말해도 같은 스킬이 동작합니다.

## 따라 하기 2 - 스크립트로 직접 돌리기

Claude 없이 단계별로 돌려 볼 수도 있습니다. 프로젝트 폴더에서 실행합니다.

```powershell
$S = ".claude\skills\weekly-report\scripts"
$W = "weekly-report-workspace\fixtures\1026-1101"
$O = "보고서\1026-1101"

# 1. 합치기 - 통합 엑셀과 점검사항.json
.venv\Scripts\python -X utf8 $S\merge_week.py $W $O\통합_판매_1026-1101.xlsx

# 2. 엑셀 수식 계산 확인 - total_errors 가 0 이어야 합니다
.venv\Scripts\python tools\skill_run.py xlsx recalc.py $O\통합_판매_1026-1101.xlsx 60

# 3. 장표 데이터 - 지난주 통합본이 있으면 끝에 --prev <지난주.xlsx> 를 붙입니다
.venv\Scripts\python -X utf8 $S\deck_data.py $O\통합_판매_1026-1101.xlsx $O\점검사항.json $O\장표데이터.json

# 4. PPT 만들기
node $S\build_deck.js $O\장표데이터.json $O\주간보고_1026-1101.pptx

# 5. 검수 - PPT 검사와 PowerPoint 그림 뽑기
.venv\Scripts\python tools\skill_run.py pptx office/validate.py $O\주간보고_1026-1101.pptx
.venv\Scripts\python -X utf8 $S\ppt_render.py $O\주간보고_1026-1101.pptx $O\_검수
```

`merge_week.py` 의 끝 코드가 3 이면 읽지 못한 파일이 있다는 뜻입니다. 그 파일을 `지점,날짜,메뉴,수량,단가,매출` 열의 CSV 로 옮겨 `<원본 파일 이름>.csv` 로 저장합니다. 그다음 `--adapter-dir <CSV 폴더>` 를 붙여 다시 실행합니다. Claude 에게 맡기면 이 단계를 알아서 합니다.

## 내 자료로 쓰기

1. **지점과 메뉴 설정** - `.claude\skills\weekly-report\references\config.json` 을 내 가게에 맞게 고칩니다.

   ```json
   {
    "브랜드": "도담커피",
    "지점": [{"이름": "강남", "별칭": ["강남점"]}],
    "메뉴": [{"이름": "아이스 아메리카노", "단가": 4500, "분류": "커피", "별칭": ["아아"]}]
   }
   ```

   - `별칭`: 파일 이름이나 표 안에서 같은 지점·메뉴로 볼 다른 이름입니다
   - `단가`: 기준 단가입니다. 단가 오기를 찾을 때 씁니다

2. **주 폴더 만들기** - 프로젝트 폴더 바로 아래에 `MMDD-MMDD` 이름의 폴더(예: `1109-1115`)를 만들고 지점 엑셀을 넣습니다.
   - 파일 이름에 지점 이름이나 별칭이 들어 있으면 됩니다
   - 표는 세로형(날짜·메뉴·수량·단가·매출 열)이든 가로형(메뉴×날짜)이든 알아서 읽습니다
   - "수요일 행사" 같은 메모를 `.txt` 로 함께 넣으면 보고서에 남습니다

3. **Claude 에게 말하기** - "주간 보고서 만들어 줘". 폴더를 말하지 않으면 가장 최근 주 폴더를 고릅니다.

이 저장소의 `.gitignore` 는 예제 두 주(`1019-1025`, `1026-1101`)만 올리고, 새로 만든 주 폴더와 `보고서\` 는 올리지 않게 막아 둡니다. 실제 판매 자료를 실수로 공개하지 않기 위해서입니다.

## 이상한 데이터 처리 원칙

원본 숫자를 최대한 살리고, 실수가 분명할 때만 고치고, 고친 것은 모두 남깁니다.

| 상황 | 처리 |
|---|---|
| 날짜·메뉴·수량·단가가 같은 줄이 또 있음 | 복사 실수로 보고 한 줄만 남김 |
| 수량이 음수 | 환불로 보고 그대로 차감 |
| 하루 기록이 통째로 없음 | 휴무로 추정, 일평균은 영업일 기준 |
| 이번 주 기간 밖 날짜 | 제외 |
| 합계·소계 줄 | 제외 |
| 단가 칸만 틀리고 금액은 기준 단가와 맞음 | 기준 단가로 고침 |
| 금액이 기준과 다른 단가로 일관됨 | 행사·할인일 수 있어 파일 값 그대로 |
| 기준에 없는 메뉴 | 신메뉴로 보고 그대로 집계 |
| 지점 파일 없음 | 그 지점을 빼고 보고 |

전체 절차와 판단 기준은 [`.claude/skills/weekly-report/SKILL.md`](.claude/skills/weekly-report/SKILL.md) 에 있습니다.

## 폴더 구조

```
.claude/skills/weekly-report/
  SKILL.md                 스킬 절차 (Claude 가 읽는 문서)
  references/config.json   지점·메뉴·기준 단가·별칭
  scripts/
    merge_week.py          지점 엑셀 합치기 + 이상 데이터 판단
    deck_data.py           장표용 숫자와 문장 만들기 (전주 대비 포함)
    build_deck.js          6장 PPT 만들기 (pptxgenjs)
    ppt_render.py          실제 PowerPoint 로 장표 그림 뽑기
tools/
  skill_run.py             플러그인 스크립트를 Windows 에서 돌리는 실행기
  winpatch/                skill-creator 평가 스크립트의 Windows 보정
  build_merged.py 등       첫 주를 손으로 만들 때 쓴 처음 버전 (지금은 쓰지 않음)
weekly-report-workspace/
  fixtures/                시험용 가짜 판매 자료
  evals.json               스킬 시험 문장과 기대 결과
  grade.py                 시험 결과 자동 채점
  make_fixtures.py         가짜 자료를 만드는 스크립트
  description-opt/발동평가.json   스킬 발동 평가 문장
CLAUDE.md                  이 프로젝트에서 Claude 가 지킬 작업 규칙
```

## 알아 두면 좋은 점

- **PPT 검수는 PowerPoint 그림으로 합니다.** LibreOffice 는 한글을 단어 단위로 줄바꿈하고 `943만` 을 `943 만` 으로 띄워 그려서, PowerPoint 에서 생기는 글자 단위 줄바꿈 문제를 가립니다.
- **플러그인 스크립트는 `tools\skill_run.py` 로 부릅니다.** `document-skills` 의 `office/soffice.py` 는 리눅스용이라 Windows 에서 `socket.AF_UNIX` 오류로 멈춥니다. 실행기가 이를 우회합니다.
- **스킬 발동 평가를 돌릴 때는 진짜 스킬을 잠시 숨깁니다.** `skill-creator` 의 `run_eval`·`run_loop` 는 임시 사본 이름이 불렸을 때만 발동으로 셉니다. 진짜 `weekly-report` 가 보이면 Claude 가 그것을 불러서 발동 0 으로 잘못 나옵니다. 평가 동안만 `SKILL.md` 이름을 바꾸고, `--skill-path` 는 사본 폴더를 가리키게 합니다. 작업자는 `--num-workers 1` 로 둡니다.
