# vybe-test: powershell/enums_flags_attribute/enum_flags_underlying_type_int32
[System.FlagsAttribute()]
enum UnderTypeEnum : int {
    X = 1
}
$underlying = [System.Enum]::GetUnderlyingType([UnderTypeEnum])
if ($underlying.Name -ne "Int32") {
    Write-Host "FAIL: GetUnderlyingType expected Int32, got $($underlying.Name)"
    exit 1
}
Write-Host "PASS"
exit 0
