# vybe-test: powershell/splatting/array_position_splatting
function Concat-Strings {
    param($first, $second, $third)
    return "$first,$second,$third"
}
$values = @('a', 'b', 'c')
$result = Concat-Strings @values
if ($result -ne 'a,b,c') {
    Write-Host "FAIL: expected a,b,c, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
