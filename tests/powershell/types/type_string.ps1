# vybe-test: powershell/types/type_string
[string]$x = "hello"
$type = $x.GetType().Name
if ($type -ne "String") {
    Write-Host "FAIL: expected 'String', got '$type'"
    exit 1
}
Write-Host "PASS"
exit 0
