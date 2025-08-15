param (
    [Parameter(Mandatory=$true, HelpMessage="请输入快捷方式所在目录路径")]
    [string]$ShortcutDirectory
)

# 验证目录是否存在
if (-not (Test-Path -Path $ShortcutDirectory -PathType Container)) {
    Write-Error "目录不存在: $ShortcutDirectory"
    exit 1
}

$paths = @()
Get-ChildItem -Path $ShortcutDirectory -Filter *.lnk | ForEach-Object {
    try {
        $shell = New-Object -ComObject WScript.Shell
        $target = $shell.CreateShortcut($_.FullName).TargetPath
        $paths += $target
    }
    catch {
        Write-Warning "无法解析快捷方式: $($_.FullName)"
    }
}

# 输出结果到控制台并复制到剪贴板
if ($paths.Count -gt 0) {
    $output = Join-Path $ShortcutDirectory "output.txt"
    Write-Output "文件已放置在：$($output) "
    $paths | Out-string | out-file $($output)
} else {
    Write-Output "未找到有效的快捷方式文件"
}
