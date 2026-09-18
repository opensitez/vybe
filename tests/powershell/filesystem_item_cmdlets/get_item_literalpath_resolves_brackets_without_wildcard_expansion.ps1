# vybe-test: powershell/filesystem_item_cmdlets/get_item_literalpath_resolves_brackets_without_wildcard_expansion
# Get-Item with -LiteralPath accurately resolves file names containing square brackets without wildcard expansion
$dir = Join-Path ([System.IO.Path]::GetTempPath()) "vybe_lit_dir_$PID"
New-Item -Path $dir -ItemType Directory -Force | Out-Null
$bracketFile = Join-Path $dir "sample[1].txt"
Set-Content -LiteralPath $bracketFile -Value "bracketed filename payload"

try {
    $item = Get-Item -LiteralPath $bracketFile

    if ($item.Name -ne "sample[1].txt") {
        Write-Host "FAIL: unexpected item name: '$($item.Name)'"
        exit 1
    }

    if (-not $item.Exists) {
        Write-Host "FAIL: FileInfo.Exists is false for literal path"
        exit 1
    }
} finally {
    if (Test-Path -LiteralPath $dir) {
        Remove-Item -LiteralPath $dir -Recurse -Force
    }
}

Write-Host "PASS"
exit 0
