# vybe-test: powershell/json_enum_and_primitive_serialization/enum_getnames_against_deserialized_json
enum FruitEnum { Apple; Banana }
$json = '{"Fruit":"Apple"}'
$obj = $json | ConvertFrom-Json
$validNames = [System.Enum]::GetNames([FruitEnum])
if (-not ($validNames -contains $obj.Fruit)) {
    Write-Host "FAIL: Deserialized JSON fruit not in Enum.GetNames"
    exit 1
}
Write-Host "PASS"
exit 0
