# vybe-test: powershell/types/type_double
[double]$x = 3.14
$type = $x.GetType().Name
if ($type -ne "Double") {
    Write-Host "FAIL: expected 'Double', got '$type'"
    exit 1
}
Write-Host "PASS"
exit 0
