param (
    [Parameter(Mandatory=$true, HelpMessage="�������ݷ�ʽ����Ŀ¼·��")]
    [string]$ShortcutDirectory
)

# ��֤Ŀ¼�Ƿ����
if (-not (Test-Path -Path $ShortcutDirectory -PathType Container)) {
    Write-Error "Ŀ¼������: $ShortcutDirectory"
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
        Write-Warning "�޷�������ݷ�ʽ: $($_.FullName)"
    }
}

# ������������̨�����Ƶ�������
if ($paths.Count -gt 0) {
    $output = Join-Path $ShortcutDirectory "output.txt"
    Write-Output "�ļ��ѷ����ڣ�$($output) "
    $paths | Out-string | out-file $($output)
} else {
    Write-Output "δ�ҵ���Ч�Ŀ�ݷ�ʽ�ļ�"
}
