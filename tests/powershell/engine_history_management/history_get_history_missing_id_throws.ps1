# vybe-test: powershell/engine_history_management/history_get_history_missing_id_throws
# Calling Get-History with a non-existent Id throws an error with -ErrorAction Stop
$threwError = $false
try {
    Get-History -Id 999999 -ErrorAction Stop
} catch {
    $threwError = $true
}

if (-not $threwError) {
    Write-Host "FAIL: Get-History -Id with non-existent ID did not throw"
    exit 1
}

Write-Host "PASS"
exit 0
