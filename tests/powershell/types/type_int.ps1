# vybe-test: powershell/types/type_int
[int]$x = 42
$type = $x.GetType().Name
if ($type -ne "Int32") {
    Write-Host "FAIL: expected 'Int32', got '$type'"
    exit 1
}
Write-Host "PASS"
exit 0
