param(
    [string]$Url = 'https://erp.xair.cn/K3Cloud/html5/',
    [int]$Port = 9333,
    [string]$ProfileDir = '',
    [int]$WaitSeconds = 60
)

#拉起带 CDP 调试端口的 Edge 实例（网页端后台填写的前置，风格对齐 start_kingdee.ps1）
#  独立 profile（登录态持久化在盘上，首次人工登录一次，之后长期复用）
#  幂等：端口已有响应就直接退出；agent 侧 web_actions.py 直连这个端口，不碰 MAA 控制器
#  端口活着但 K3Cloud 页签被关 → 用 CDP 补开一个（窗口标题变了 Win32 控制器就找不到窗口）
#  四个 disable-* 参数：禁止 Chromium 对后台/遮挡窗口降频渲染与定时器 —— 后台自动化的前提
#  其中 CalculateNativeWinOcclusion 专治实时视图黑屏：Edge 被判"被遮挡"就停止产帧，
#  MFAAvalonia FramePool 捕获随之黑屏（2026-11-24 用户实测：点一下 Edge 再切回 GUI 就黑）
if ([string]::IsNullOrWhiteSpace($ProfileDir)) {
    $ProfileDir = Join-Path $env:LOCALAPPDATA 'MaaXA\web-profile'
}

function Test-Cdp {
    try {
        $r = Invoke-WebRequest -Uri "http://127.0.0.1:$Port/json/version" -UseBasicParsing -TimeoutSec 3
        return $r.StatusCode -eq 200
    } catch {
        return $false
    }
}

function Get-K3CloudTab {
    try {
        $r = Invoke-RestMethod -Uri "http://127.0.0.1:$Port/json" -TimeoutSec 3
        return @($r | Where-Object { $_.type -eq 'page' -and $_.url -match '/k3cloud/' })
    } catch {
        return @()
    }
}

function Ensure-K3CloudTab {
    #已有 K3Cloud 页签 → True；没有就 PUT /json/new 补开一个并等它挂上
    if ((Get-K3CloudTab).Count -gt 0) { return $true }
    try {
        $enc = [uri]::EscapeDataString($Url)
        Invoke-RestMethod -Method Put -Uri "http://127.0.0.1:$Port/json/new?$enc" -TimeoutSec 5 | Out-Null
    } catch {
        Write-Output "open tab failed: $($_.Exception.Message)"
        return $false
    }
    $deadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 800
        if ((Get-K3CloudTab).Count -gt 0) { return $true }
    }
    return $false
}

$mutex = New-Object System.Threading.Mutex($false, 'Local\MaaXA.StartWeb')
$hasLock = $false
$exitCode = 0
try {
    try { $hasLock = $mutex.WaitOne(30000) }
    catch [System.Threading.AbandonedMutexException] { $hasLock = $true }
    if (-not $hasLock) { throw "Timed out waiting for the startup lock." }

    if (Test-Cdp) {
        if (Ensure-K3CloudTab) {
            Write-Output "CDP port $Port alive and K3Cloud tab ensured; skipping startup."
            exit 0
        }
        Write-Error "Edge is up but the K3Cloud tab could not be opened on port $Port."
        exit 3
    }

    $edgeCandidates = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe')
    )
    $exe = $edgeCandidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
    if ($null -eq $exe) {
        Write-Error "msedge.exe not found. Install Edge or pass -Url/-ProfileDir for another Chromium build."
        exit 2
    }
    if ($WaitSeconds -le 0) { Write-Error "WaitSeconds must be greater than zero."; exit 2 }

    if (-not (Test-Path -LiteralPath $ProfileDir)) {
        New-Item -ItemType Directory -Path $ProfileDir -Force | Out-Null
    }

    Start-Process -FilePath $exe -ArgumentList @(
        "--user-data-dir=$ProfileDir",
        "--remote-debugging-port=$Port",
        "--remote-allow-origins=*",
        '--no-first-run',
        '--no-default-browser-check',
        '--disable-background-timer-throttling',
        '--disable-backgrounding-occluded-windows',
        '--disable-renderer-backgrounding',
        '--disable-features=CalculateNativeWinOcclusion',
        $Url
    ) | Out-Null
    Write-Output "Started Edge with CDP on port $Port; waiting for the debug endpoint."

    $deadline = (Get-Date).AddSeconds($WaitSeconds)
    do {
        Start-Sleep -Milliseconds 500
    } while (-not (Test-Cdp) -and (Get-Date) -lt $deadline)
    if (-not (Test-Cdp)) { throw "CDP endpoint on port $Port did not come up within $WaitSeconds seconds." }
    Write-Output "CDP ready: http://127.0.0.1:$Port/json  (profile: $ProfileDir)"
} catch {
    Write-Error $_.Exception.Message
    $exitCode = 1
} finally {
    if ($hasLock) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
exit $exitCode
