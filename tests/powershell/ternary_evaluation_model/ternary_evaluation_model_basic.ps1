# vybe-test: powershell/ternary_evaluation_model/basic
$flag = $true
$value = if ($flag) { 'enabled' } else { 'disabled' }

if ($value -ne 'enabled') {
    Write-Host "FAIL: expected 'enabled', got '$value'"
    exit 1
}

$other = if ($false) { 1 } else { 2 }
if ($other -ne 2) {
    Write-Host "FAIL: fallback branch failed, got $other"
    exit 1
}

Write-Host 'PASS'
exit 0
