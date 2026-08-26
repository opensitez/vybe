# vybe-test: powershell/json_enum_and_primitive_serialization/parse_json_integer_to_enum
enum TargetEnum { First = 1; Second = 2 }
$json = '{"Choice":2}'
$obj = $json | ConvertFrom-Json
$e = [TargetEnum]$obj.Choice
if ($e -ne [TargetEnum]::Second) {
    Write-Host "FAIL: Cast parsed JSON integer to enum failed"
    exit 1
}
Write-Host "PASS"
exit 0
