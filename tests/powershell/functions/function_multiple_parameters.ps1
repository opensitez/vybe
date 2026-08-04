# vybe-test: powershell/functions/function_multiple_parameters
function Add-Numbers {
    param($a, $b)
    return $a + $b
}
$result = Add-Numbers -a 7 -b 8
if ($result -ne 15) {
    Write-Host "FAIL: expected 15, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
