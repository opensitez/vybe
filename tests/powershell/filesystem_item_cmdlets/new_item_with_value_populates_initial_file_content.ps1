# vybe-test: powershell/filesystem_item_cmdlets/new_item_with_value_populates_initial_file_content
# New-Item creates a file and immediately populates its content when -Value is provided
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) "vybe_ni_val_$PID.txt"

try {
    New-Item -Path $tmp -ItemType File -Value "initial data payload" -Force | Out-Null
    $content = Get-Content -Path $tmp -Raw

    if ($content.Trim() -ne "initial data payload") {
        Write-Host "FAIL: expected 'initial data payload', got '$content'"
        exit 1
    }
} finally {
    if (Test-Path $tmp) {
        Remove-Item $tmp -Force
    }
}

Write-Host "PASS"
exit 0
