param(
    [string]$Exe = '',                                    #不传则自动探测金蝶客户端安装路径
    [Parameter(Mandatory = $true)]
    [string]$WindowRegex,
    [int]$WaitSeconds = 60
)

#金蝶云客户端的默认安装位置（相对各盘的 Program Files 目录）
$kingdeeRelativePath = 'Kingdee\K3Cloud\DeskClient\K3CloudClient\Kingdee.BOS.XPF.App.exe'

function Find-KingdeeExe {
    #遍历所有磁盘 × 两种 Program Files，命中即返回；都没有则返回 $null，由调用方报错
    foreach ($drive in (Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue).Root) {
        foreach ($programFiles in @('Program Files (x86)', 'Program Files')) {
            $candidate = Join-Path $drive (Join-Path $programFiles $kingdeeRelativePath)
            if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
        }
    }
    return $null
}

if ([string]::IsNullOrWhiteSpace($Exe)) {
    $Exe = Find-KingdeeExe
    if ($null -eq $Exe) {
        Write-Error "未能自动找到金蝶云客户端，请用 -Exe 参数手动指定 Kingdee.BOS.XPF.App.exe 的完整路径。"
        exit 2
    }
    Write-Output "Auto-detected Kingdee client: $Exe"
}

if (-not (Test-Path -LiteralPath $Exe -PathType Leaf)) {
    Write-Error "Target executable was not found: $Exe"
    exit 2
}
if ($WaitSeconds -le 0) {
    Write-Error "WaitSeconds must be greater than zero."
    exit 2
}

Add-Type @'
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

public sealed class MaaWindowMatch
{
    public int ProcessId { get; private set; }
    public string Title { get; private set; }

    public MaaWindowMatch(int processId, string title)
    {
        ProcessId = processId;
        Title = title;
    }
}

public static class MaaWindowApi
{
    public delegate bool EnumWindowsProc(IntPtr handle, IntPtr parameter);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr parameter);
    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr handle);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr handle, StringBuilder text, int length);
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);

    public static MaaWindowMatch Find(string pattern, string[] processNames, string expectedPath)
    {
        var regex = new Regex(pattern, RegexOptions.IgnoreCase);
        MaaWindowMatch match = null;
        EnumWindows(delegate(IntPtr handle, IntPtr parameter)
        {
            if (!IsWindowVisible(handle)) return true;
            var text = new StringBuilder(512);
            if (GetWindowText(handle, text, text.Capacity) <= 0 || !regex.IsMatch(text.ToString())) return true;
            uint processId;
            GetWindowThreadProcessId(handle, out processId);
            if (processId == 0) return true;
            try
            {
                using (var process = Process.GetProcessById((int)processId))
                {
                    if (!processNames.Any(name => string.Equals(process.ProcessName, name, StringComparison.OrdinalIgnoreCase))) return true;
                    try
                    {
                        var path = process.MainModule.FileName;
                        if (!string.IsNullOrEmpty(path) && !SamePath(path, expectedPath)) return true;
                    }
                    catch { }
                }
            }
            catch { return true; }
            match = new MaaWindowMatch((int)processId, text.ToString());
            return false;
        }, IntPtr.Zero);
        return match;
    }

    private static bool SamePath(string left, string right)
    {
        return string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), StringComparison.OrdinalIgnoreCase);
    }
}
'@

$targetProcessNames = @(
    [IO.Path]::GetFileNameWithoutExtension($Exe),
    'Kingdee.BOS.XPF'
) | Select-Object -Unique

function Test-TargetProcess {
    param([System.Diagnostics.Process]$Process)
    if ($null -eq $Process -or $targetProcessNames -notcontains $Process.ProcessName) { return $false }
    try {
        $processPath = $Process.MainModule.FileName
        if (-not [string]::IsNullOrWhiteSpace($processPath)) {
            return [string]::Equals(
                [IO.Path]::GetFullPath($processPath),
                [IO.Path]::GetFullPath($Exe),
                [StringComparison]::OrdinalIgnoreCase)
        }
    } catch { }
    return $true
}

function Get-TargetProcess {
    foreach ($name in $targetProcessNames) {
        $process = Get-Process -Name $name -ErrorAction SilentlyContinue |
            Where-Object { Test-TargetProcess $_ } |
            Select-Object -First 1
        if ($null -ne $process) { return $process }
    }
    return $null
}

function Get-TargetWindow {
    try {
        $match = [MaaWindowApi]::Find($WindowRegex, $targetProcessNames, $Exe)
        if ($null -eq $match) { return $null }
        return $match
    } catch {
        throw "Invalid window title regex or window enumeration failed: $($_.Exception.Message)"
    }
}

$mutex = New-Object System.Threading.Mutex($false, 'Local\MAAXA.StartKingdee')
$hasLock = $false
$exitCode = 0
try {
    try { $hasLock = $mutex.WaitOne(30000) }
    catch [System.Threading.AbandonedMutexException] { $hasLock = $true }
    if (-not $hasLock) { throw "Timed out waiting for the startup lock." }

    $window = Get-TargetWindow
    if ($null -ne $window) {
        Write-Output "Target window already exists; skipping startup: $($window.Title)"
    } else {
        $process = Get-TargetProcess
        if ($null -eq $process) {
            Start-Process -FilePath $Exe -WorkingDirectory ([IO.Path]::GetDirectoryName($Exe)) -ErrorAction Stop | Out-Null
            Write-Output "Started target application; waiting for its window."
        } else {
            Write-Output "Target process already exists; waiting for its window."
        }

        $deadline = (Get-Date).AddSeconds($WaitSeconds)
        do {
            Start-Sleep -Milliseconds 500
            $window = Get-TargetWindow
        } while ($null -eq $window -and (Get-Date) -lt $deadline)
        if ($null -eq $window) { throw "No matching target window appeared within $WaitSeconds seconds." }
        Write-Output "Target window is ready: $($window.Title)"
    }
} catch {
    Write-Error $_.Exception.Message
    $exitCode = 1
} finally {
    if ($hasLock) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
exit $exitCode
