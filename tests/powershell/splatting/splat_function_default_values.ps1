# vybe-test: powershell/splatting/splat_function_default_values
function Get-Range {
    param($start, $end = 5)
    return $end - $start
}
$args = @{ start = 2 }
$result = Get-Range @args
if ($result -ne 3) {
    Write-Host "FAIL: expected 3, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
