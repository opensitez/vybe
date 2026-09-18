# vybe-test: powershell/filesystem_item_cmdlets/get_item_wildcard_matching_multiple_files
# Get-Item with wildcard pattern matches multiple files matching the extension
$dir = Join-Path ([System.IO.Path]::GetTempPath()) "vybe_wild_dir_$PID"
New-Item -Path $dir -ItemType Directory -Force | Out-Null
"a" | Set-Content (Join-Path $dir "alpha.log")
"b" | Set-Content (Join-Path $dir "beta.log")
"c" | Set-Content (Join-Path $dir "gamma.txt")

try {
    $matched = @(Get-Item (Join-Path $dir "*.log"))

    if ($matched.Count -ne 2) {
        Write-Host "FAIL: expected 2 .log files, got $($matched.Count)"
        exit 1
    }

    $names = $matched | ForEach-Object { $_.Name } | Sort-Object
    if ($names[0] -ne "alpha.log" -or $names[1] -ne "beta.log") {
        Write-Host "FAIL: matched file names mismatch: $($names -join ', ')"
        exit 1
    }
} finally {
    if (Test-Path $dir) {
        Remove-Item $dir -Recurse -Force
    }
}

Write-Host "PASS"
exit 0
