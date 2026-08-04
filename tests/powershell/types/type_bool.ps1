# vybe-test: powershell/types/type_bool
[bool]$x = $true
$type = $x.GetType().Name
if ($type -ne "Boolean") {
    Write-Host "FAIL: expected 'Boolean', got '$type'"
    exit 1
}
Write-Host "PASS"
exit 0
