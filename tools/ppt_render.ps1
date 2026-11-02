# _검수 폴더의 .pptx 를 실제 PowerPoint 로 열어 장마다 PNG 로 내보낸다 (읽기 전용으로 연다)
# 사용법: _검수 폴더에 덱을 복사한 뒤  powershell -NoProfile -ExecutionPolicy Bypass -File tools\ppt_render.ps1
# 결과: _검수\pp-<파일이름>-<번호>.png, 로그는 _검수\열기시험.log. 이미 실행 중이던 PowerPoint 는 닫지 않는다.
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$dir = Join-Path (Split-Path $PSScriptRoot) '_검수'; $log = Join-Path $dir '열기시험.log'
$wasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$app = New-Object -ComObject PowerPoint.Application
foreach ($f in Get-ChildItem $dir -Filter '*.pptx') {
    try {
        $p = $app.Presentations.Open($f.FullName, -1, 0, 0)
        foreach ($i in 1..$p.Slides.Count) { $p.Slides.Item($i).Export((Join-Path $dir ("pp-" + $f.BaseName + "-$i.png")), 'PNG', 1280, 720) }
        Add-Content $log "성공: $($f.Name) ($($p.Slides.Count)장)" -Encoding UTF8; $p.Close()
    } catch { Add-Content $log "실패: $($f.Name) - $($_.Exception.Message)" -Encoding UTF8 }
}
if (-not $wasRunning) { $app.Quit() }
