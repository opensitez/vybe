# vybe-test: powershell/functions/function_ref_parameter
function Increment([ref]$value) {
    $value.Value++
}
$n = 10
Increment ([ref]$n)
if ($n -ne 11) {
    Write-Host "FAIL: expected 11, got $n"
    exit 1
}
Write-Host "PASS"
exit 0
