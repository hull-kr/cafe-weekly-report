"""주간 보고에 필요한 것을 확인하고, 빠진 것만 설치한다.

weekly-report 스킬이 맨 처음에 실행한다. 다 갖춰져 있으면 몇 초 만에 끝난다.
처음 한 번은 설치 때문에 몇 분 걸린다.

    python tools\\setup.py

확인하는 것
  1. .venv 와 파이썬 패키지(requirements.txt)
  2. node_modules (package.json)
  3. LibreOffice (엑셀 수식 계산 · PPT 검사)
  4. document-skills 플러그인 (이 폴더에만, local scope)
"""
import glob
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent.parent
VENV_PY = ROOT / ".venv" / "Scripts" / "python.exe"
SOFFICE = Path(r"C:\Program Files\LibreOffice\program\soffice.exe")
PLUGIN = "document-skills@anthropic-agent-skills"
문제 = []


def 실행(명령, **kw):
    print("정보: 실행 -", " ".join(str(x) for x in 명령), flush=True)
    return subprocess.call([str(x) for x in 명령], cwd=ROOT, **kw)


def 파이썬():
    if VENV_PY.exists() and subprocess.call(
            [str(VENV_PY), "-c", "import openpyxl, pandas, pptx, defusedxml, lxml"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL) == 0:
        print("성공: 파이썬 패키지 준비됨")
        return
    print("정보: 파이썬 패키지를 설치합니다(.venv)")
    if not VENV_PY.exists() and 실행([sys.executable, "-m", "venv", ".venv"]) != 0:
        문제.append(".venv 를 만들지 못했습니다")
        return
    실행([VENV_PY, "-m", "pip", "install", "-q", "--upgrade", "pip"])
    if 실행([VENV_PY, "-m", "pip", "install", "-q", "-r", "requirements.txt"]) != 0:
        문제.append("파이썬 패키지 설치 실패")
    else:
        print("성공: 파이썬 패키지 설치")


def 노드():
    if (ROOT / "node_modules" / "pptxgenjs").exists():
        print("성공: 노드 패키지 준비됨")
        return
    npm = shutil.which("npm") or shutil.which("npm.cmd")
    if not npm:
        문제.append("Node.js(npm)가 없습니다. Node.js 를 먼저 설치해야 합니다")
        return
    print("정보: 노드 패키지를 설치합니다(node_modules)")
    if 실행([npm, "install", "--no-fund", "--no-audit"]) != 0:
        문제.append("노드 패키지 설치 실패")
    else:
        print("성공: 노드 패키지 설치")


def 리브레오피스():
    if SOFFICE.exists():
        print("성공: LibreOffice 준비됨")
        return
    print("정보: LibreOffice 를 설치합니다(winget, 몇 분 걸림)")
    실행(["winget", "install", "--id", "TheDocumentFoundation.LibreOffice", "-e",
         "--accept-package-agreements", "--accept-source-agreements", "--silent"])
    if SOFFICE.exists():
        print("성공: LibreOffice 설치")
    else:
        문제.append("LibreOffice 설치 실패 - 관리자 확인 창이 떴다면 「예」를 누르고 다시 실행합니다")


def 플러그인():
    설정 = ROOT / ".claude" / "settings.local.json"
    try:
        켜짐 = json.loads(설정.read_text(encoding="utf-8")).get("enabledPlugins", {}).get(PLUGIN)
    except (OSError, ValueError):
        켜짐 = False
    있음 = glob.glob(os.path.expanduser(
        r"~\.claude\plugins\cache\anthropic-agent-skills\document-skills\*\skills"))
    if 켜짐 and 있음:
        print("성공: document-skills 플러그인 준비됨")
        return
    claude = shutil.which("claude") or shutil.which("claude.cmd") or shutil.which("claude.exe")
    if not claude:
        문제.append("claude 명령을 찾지 못해 플러그인을 설치하지 못했습니다")
        return
    print("정보: document-skills 플러그인을 이 폴더에만(local scope) 설치합니다")
    실행([claude, "plugin", "marketplace", "add", "anthropics/skills"])  # 이미 있으면 그냥 넘어간다
    if 실행([claude, "plugin", "install", PLUGIN, "--scope", "local"]) != 0:
        문제.append("document-skills 플러그인 설치 실패")
    else:
        print("성공: document-skills 플러그인 설치")


if __name__ == "__main__":
    파이썬()
    노드()
    리브레오피스()
    플러그인()
    if 문제:
        print("실패: " + " / ".join(문제))
        sys.exit(1)
    print("성공: 준비 끝")
