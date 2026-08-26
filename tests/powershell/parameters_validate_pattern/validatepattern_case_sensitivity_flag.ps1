# vybe-test: powershell/parameters_validate_pattern/validatepattern_case_sensitivity_flag
function Set-HexColor {
    param([ValidatePattern('^#[0-9A-F]{6}$')][string]$Hex)
    return $Hex
}
$res = Set-HexColor -Hex "#FFAA00"
if ($res -ne "#FFAA00") {
    Write-Host "FAIL: ValidatePattern uppercase hex failed"
    exit 1
}
Write-Host "PASS"
exit 0
