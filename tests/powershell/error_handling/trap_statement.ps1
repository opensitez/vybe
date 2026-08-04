# vybe-test: powershell/error_handling/trap_statement
trap {
    $global:errorCaught = $true
    continue
}
$global:errorCaught = $false
throw "test error"
if (-not $global:errorCaught) {
    Write-Host "FAIL: expected error to be caught by trap"
    exit 1
}
Write-Host "PASS"
exit 0
