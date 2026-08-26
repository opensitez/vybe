# vybe-test: powershell/json_enum_and_primitive_serialization/parse_json_string_to_enum
enum StateEnum { Stop; Go }
$json = '{"State":"Go"}'
$obj = $json | ConvertFrom-Json
$e = [StateEnum]$obj.State
if ($e -ne [StateEnum]::Go) {
    Write-Host "FAIL: Cast parsed JSON string to enum failed"
    exit 1
}
Write-Host "PASS"
exit 0
