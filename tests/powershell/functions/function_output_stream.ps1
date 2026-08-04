# vybe-test: powershell/functions/function_output_stream
function Get-Numbers {
    1
    2
    3
}
$result = Get-Numbers
$sum = 0
foreach ($num in $result) {
    $sum += $num
}
if ($sum -ne 6) {
    Write-Host "FAIL: expected 6, got $sum"
    exit 1
}
Write-Host "PASS"
exit 0
