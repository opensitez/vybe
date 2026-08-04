# vybe-test: powershell/functions/function_with_parameter
function Add-Five {
    param($x)
    return $x + 5
}
$result = Add-Five -x 10
if ($result -ne 15) {
    Write-Host "FAIL: expected 15, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
