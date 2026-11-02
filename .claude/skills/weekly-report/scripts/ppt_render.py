"""실제 PowerPoint 로 장표를 열어 장마다 PNG 로 내보낸다 (검수용, 읽기 전용으로 연다).

사용법: .venv\\Scripts\\python -X utf8 ppt_render.py <덱.pptx> <그림폴더>
- 한글 경로가 Bash → PowerShell 인자로 넘어가며 깨지지 않도록 경로를 base64 로 감싸 넘긴다.
- PowerPoint 가 없거나 파일을 못 열면 끝 코드 2. (못 연다 = PowerPoint 가 거부하는 파일이라는 뜻이므로 반드시 고친다)
- 이미 실행 중이거나 다른 프레젠테이션이 열려 있으면 PowerPoint 를 끄지 않는다.
"""
import base64
import subprocess
import sys
from pathlib import Path

PS = r"""
$ErrorActionPreference = 'Stop'
$src = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('%SRC%'))
$out = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('%OUT%'))
$wasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
try { $app = New-Object -ComObject PowerPoint.Application } catch { Write-Output 'RESULT:NOPPT'; exit 0 }
try {
  $p = $app.Presentations.Open($src, -1, 0, 0)
  $n = $p.Slides.Count
  foreach ($i in 1..$n) { $p.Slides.Item($i).Export((Join-Path $out ("slide-$i.png")), 'PNG', 1280, 720) }
  $p.Close()
  Write-Output "RESULT:OK:$n"
} catch {
  Write-Output ("RESULT:FAIL:" + $_.Exception.Message)
} finally {
  if (-not $wasRunning -and $app.Presentations.Count -eq 0) { $app.Quit() }
}
"""


def main():
    src, out = Path(sys.argv[1]).resolve(), Path(sys.argv[2]).resolve()
    out.mkdir(parents=True, exist_ok=True)
    for old in out.glob("slide-*.png"):
        old.unlink()
    b = lambda p: base64.b64encode(str(p).encode("utf-8")).decode()  # noqa: E731
    script = PS.replace("%SRC%", b(src)).replace("%OUT%", b(out))
    enc = base64.b64encode(script.encode("utf-16-le")).decode()
    r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-EncodedCommand", enc],
                       capture_output=True, text=True, encoding="cp949", errors="replace", timeout=300)
    # PowerShell 콘솔 출력은 시스템 코드페이지(cp949)로 나온다
    line = next((l for l in r.stdout.splitlines() if l.startswith("RESULT:")), "RESULT:FAIL:응답 없음")
    if line == "RESULT:NOPPT":
        print("경고: PowerPoint 가 없어 실제 렌더 확인을 못 함 - LibreOffice 렌더로 대신 확인할 것")
        sys.exit(2)
    if line.startswith("RESULT:FAIL"):
        print(f"실패: PowerPoint 가 파일을 열지 못함 - {line[12:]}")
        sys.exit(2)
    n = int(line.split(":")[2])
    print(f"성공: PowerPoint 로 {n}장을 그림으로 내보냄")
    for i in range(1, n + 1):
        print(out / f"slide-{i}.png")


if __name__ == "__main__":
    main()
