# vybe-test: powershell/token_and_grammar_system/keyword_if_resolved
$bodyRan = $false
if ($true) {
    $bodyRan = $true
}

if (-not $bodyRan) {
    Write-Host 'FAIL: if keyword did not execute true branch'
    exit 1
}

Write-Host 'PASS'
exit 0
