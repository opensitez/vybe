# vybe-test: powershell/enums_flags_attribute/enum_flags_is_defined_check
[System.FlagsAttribute()]
enum DefCheck {
    Alpha = 1
    Beta = 2
}
$isAlpha = [System.Enum]::IsDefined([DefCheck], "Alpha")
$isOmega = [System.Enum]::IsDefined([DefCheck], "Omega")
if (-not $isAlpha -or $isOmega) {
    Write-Host "FAIL: Enum.IsDefined check failed"
    exit 1
}
Write-Host "PASS"
exit 0
