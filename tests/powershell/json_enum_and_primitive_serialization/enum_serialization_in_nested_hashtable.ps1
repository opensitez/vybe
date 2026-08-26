# vybe-test: powershell/json_enum_and_primitive_serialization/enum_serialization_in_nested_hashtable
enum ModeType { Debug; Release }
$ht = @{ Config = @{ Mode = [ModeType]::Release } }
$json = $ht | ConvertTo-Json -Depth 3
$recovered = $json | ConvertFrom-Json
if ($recovered.Config.Mode -ne 1 -and $recovered.Config.Mode -ne "Release") {
    Write-Host "FAIL: Enum in nested hashtable serialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
