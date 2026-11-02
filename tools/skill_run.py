"""document-skills 스크립트를 Windows 에서 돌리는 실행기.

플러그인의 office/soffice.py 는 리눅스 샌드박스용이라 socket.AF_UNIX 를 검사하는데,
Windows 의 Python 에는 그 속성이 없어 AttributeError 로 죽는다.
Windows 에서는 LD_PRELOAD shim 이 필요 없으므로 검사를 건너뛰게 바꿔 끼운 뒤 원래 스크립트를 실행한다.

사용법:
    .venv\\Scripts\\python tools\\skill_run.py xlsx recalc.py 파일.xlsx
    .venv\\Scripts\\python tools\\skill_run.py pptx thumbnail.py 덱.pptx 덱-thumbs
    .venv\\Scripts\\python tools\\skill_run.py pptx office/soffice.py --headless --convert-to pdf 덱.pptx

skill-creator 스크립트는 모듈 이름으로 부른다 (스킬 문서의 `python -m scripts.xxx` 와 같다).
Windows 에서 죽는 select 를 tools/winpatch/sitecustomize.py 로 보정해 자식 프로세스까지 적용한다.
작업 폴더는 이 프로젝트 그대로 둔다 (run_eval 이 여기서 .claude 폴더를 찾는다).
    .venv\\Scripts\\python tools\\skill_run.py skill-creator scripts.quick_validate 스킬폴더
    .venv\\Scripts\\python tools\\skill_run.py skill-creator scripts.package_skill 스킬폴더
    .venv\\Scripts\\python tools\\skill_run.py skill-creator scripts.run_loop --eval-set ... --skill-path ...
"""
import glob
import os
import runpy
import sys
from pathlib import Path

LIBREOFFICE = r"C:\Program Files\LibreOffice\program"


def skills_root() -> Path:
    # 플러그인 버전 폴더는 업데이트마다 바뀌므로 가장 최근 것을 고른다
    pattern = os.path.expanduser(
        r"~\.claude\plugins\cache\anthropic-agent-skills\document-skills\*\skills"
    )
    found = sorted(glob.glob(pattern), key=os.path.getmtime)
    if not found:
        sys.exit("실패: document-skills 플러그인 폴더를 찾지 못했습니다")
    return Path(found[-1])


def run_skill_creator(module: str, rest: list) -> int:
    import subprocess

    pattern = os.path.expanduser(
        r"~\.claude\plugins\cache\claude-plugins-official\skill-creator\*\skills\skill-creator"
    )
    found = sorted(glob.glob(pattern), key=os.path.getmtime)
    if not found:
        sys.exit("실패: skill-creator 플러그인 폴더를 찾지 못했습니다")
    winpatch = Path(__file__).resolve().parent / "winpatch"
    env = os.environ.copy()
    env["PYTHONPATH"] = os.pathsep.join(
        [str(winpatch), found[-1], *filter(None, [env.get("PYTHONPATH")])])
    env["PYTHONUTF8"] = "1"
    return subprocess.call([sys.executable, "-m", module, *rest], env=env)


def main() -> None:
    # 한글 출력을 cp949 로 읽다 죽지 않게 UTF-8 모드로 다시 실행한다
    if not sys.flags.utf8_mode:
        import subprocess

        sys.exit(subprocess.call([sys.executable, "-X", "utf8", *sys.argv]))
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    skill, script, *rest = sys.argv[1:]
    if skill == "skill-creator":
        sys.exit(run_skill_creator(script, rest))
    scripts_dir = skills_root() / skill / "scripts"
    target = scripts_dir / script
    if not target.exists():
        sys.exit(f"실패: 스크립트가 없습니다 - {target}")

    # soffice 를 PATH 에서 찾을 수 있게 한다 (새 터미널 전이라도 동작하도록)
    if LIBREOFFICE not in os.environ.get("PATH", ""):
        os.environ["PATH"] = LIBREOFFICE + os.pathsep + os.environ.get("PATH", "")

    sys.path.insert(0, str(scripts_dir))
    import office.soffice as soffice  # noqa: E402

    soffice._needs_shim = lambda: False  # Windows 에서는 shim 불필요

    # soffice.py 자체를 부르면 새 모듈로 다시 읽혀 위 교체가 무시되므로 직접 호출한다
    if Path(script).as_posix() == "office/soffice.py":
        sys.exit(soffice.run_soffice(rest).returncode)

    # 직접 실행할 때처럼 스크립트가 있는 폴더도 import 경로에 넣는다 (office/validate.py 의 helpers 등)
    sys.path.insert(0, str(target.parent))
    sys.argv = [str(target), *rest]
    runpy.run_path(str(target), run_name="__main__")


if __name__ == "__main__":
    main()
