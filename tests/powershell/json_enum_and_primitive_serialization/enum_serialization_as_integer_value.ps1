# vybe-test: powershell/json_enum_and_primitive_serialization/enum_serialization_as_integer_value
enum LogLevel { Info = 1; Warn = 2; Error = 3 }
$obj = @{ Level = [LogLevel]::Warn }
$json = $obj | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.Level -ne 2 -and $recovered.Level -ne "Warn") {
    Write-Host "FAIL: Enum serialization failed, got '$($recovered.Level)'"
    exit 1
}
Write-Host "PASS"
exit 0
