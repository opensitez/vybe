# vybe-test: powershell/splatting/named_and_positional_splat
function Join-Values {
    param($first, $second, $third)
    return "$first:$second:$third"
}
$args = @{ second = 'B'; third = 'C' }
$result = Join-Values 'A' @args
if ($result -ne 'A:B:C') {
    Write-Host "FAIL: expected A:B:C, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
