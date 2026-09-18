# vybe-test: powershell/engine_history_management/history_clear_nonexistent_id_throws
# Calling Clear-History with an Id not present in the history buffer throws an error
$threwError = $false
try {
    Clear-History -Id 999999 -ErrorAction Stop
} catch {
    $threwError = $true
}

if (-not $threwError) {
    Write-Host "FAIL: Clear-History -Id with non-existent ID did not throw"
    exit 1
}

Write-Host "PASS"
exit 0
