# vybe-test: powershell/filesystem_item_cmdlets/new_item_overwrites_existing_file_with_force
# New-Item successfully overwrites an existing file and applies new content when -Force is specified
$tmp = [System.IO.Path]::GetTempFileName()
"old content" | Set-Content $tmp

try {
    New-Item -Path $tmp -ItemType File -Value "overwritten content" -Force | Out-Null
    $content = Get-Content -Path $tmp -Raw

    if ($content.Trim() -ne "overwritten content") {
        Write-Host "FAIL: expected 'overwritten content', got '$content'"
        exit 1
    }
} finally {
    if (Test-Path $tmp) {
        Remove-Item $tmp -Force
    }
}

Write-Host "PASS"
exit 0
