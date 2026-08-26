# vybe-test: powershell/json_enum_and_primitive_serialization/enum_with_negative_values
enum NegEnum { Min = -1; Zero = 0; Plus = 1 }
$obj = @{ N = [NegEnum]::Min }
$json = $obj | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.N -ne -1 -and $recovered.N -ne "Min") {
    Write-Host "FAIL: Enum with negative values serialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
