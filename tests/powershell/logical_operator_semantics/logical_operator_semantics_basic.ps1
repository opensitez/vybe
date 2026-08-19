# vybe-test: powershell/logical_operator_semantics/basic
$didRun = $false
$andResult = $true -and ($didRun = $true)

if (-not $andResult -or -not $didRun) {
    Write-Host 'FAIL: -and did not evaluate right side when expected'
    exit 1
}

$didRun = $false
$orResult = $false -or ($didRun = $true)
if (-not $orResult -or -not $didRun) {
    Write-Host 'FAIL: -or did not evaluate right side when expected'
    exit 1
}

$didRun = $false
$shortCircuit = $false -and ($didRun = $true)
if ($shortCircuit -or $didRun) {
    Write-Host 'FAIL: -and short-circuit behavior broken'
    exit 1
}

Write-Host 'PASS'
exit 0
