# vybe-test: powershell/json_enum_and_primitive_serialization/enum_array_serialization
enum Priority { Low; Medium; High }
$arr = @([Priority]::Low, [Priority]::High)
$json = @{ Pri = $arr } | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.Pri.Count -ne 2) {
    Write-Host "FAIL: Enum array serialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
