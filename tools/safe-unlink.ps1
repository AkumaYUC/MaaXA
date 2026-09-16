# 安全拆除开发期联接：只删链接本身，绝不进入链接指向的真身。
# 背景：publish-test 里的 resource / agent 是指向 MAAXA 真实资源的 Junction，
#      interface.json 是硬链接。用 rm -rf / rmdir /s 会顺着链接把真身删光。
# 用法：powershell -ExecutionPolicy Bypass -File tools\safe-unlink.ps1 -Dir "E:\...\publish-test"
param([Parameter(Mandatory = $true)][string]$Dir)

$names = @('resource', 'agent', 'interface.json')
foreach ($n in $names) {
    $p = Join-Path $Dir $n
    if (-not (Test-Path -LiteralPath $p)) {
        Write-Output ("跳过（不存在）: " + $n)
        continue
    }
    $item = Get-Item -LiteralPath $p -Force
    if ($item.LinkType -in @('Junction', 'SymbolicLink')) {
        # 关键：recursive = false，只移除链接节点本身
        [System.IO.Directory]::Delete($p, $false)
        Write-Output ("已安全移除: " + $n + "  [" + $item.LinkType + "]")
    }
    elseif ($item.LinkType -eq 'HardLink') {
        # 硬链接删任意一个名字即可，内容由其他名字继续保留
        [System.IO.File]::Delete($p)
        Write-Output ("已安全移除: " + $n + "  [HardLink]")
    }
    else {
        Write-Output ("⚠ 不是链接，已跳过以免误删真实文件: " + $n)
    }
}
Write-Output '完成。'
