# vybe-test: powershell/parameters_validate_pattern/validatepattern_reflection_attribute_inspection
function Pattern-Target {
    param([ValidatePattern('^[a-z]+$')][string]$Word)
}
$cmd = Get-Command Pattern-Target
$param = $cmd.Parameters["Word"]
$attr = $param.Attributes | Where-Object { $_.GetType().Name -eq "ValidatePatternAttribute" }
if ($attr -eq $null -or $attr.RegexPattern -ne "^[a-z]+$") {
    Write-Host "FAIL: ValidatePattern reflection inspection failed"
    exit 1
}
Write-Host "PASS"
exit 0
