# vybe-test: powershell/filesystem_item_cmdlets/new_item_fails_on_existing_file_without_force
# New-Item fails and throws when creating a file that already exists unless -Force is specified
$tmp = [System.IO.Path]::GetTempFileName()
$threw = $false

try {
    New-Item -Path $tmp -ItemType File -ErrorAction Stop | Out-Null
} catch {
    $threw = $true
} finally {
    if (Test-Path $tmp) {
        Remove-Item $tmp -Force
    }
}

if (-not $threw) {
    Write-Host "FAIL: New-Item on existing file without -Force did not throw"
    exit 1
}

Write-Host "PASS"
exit 0
