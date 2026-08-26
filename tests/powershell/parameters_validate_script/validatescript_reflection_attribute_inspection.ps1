# vybe-test: powershell/parameters_validate_script/validatescript_reflection_attribute_inspection
function Script-Target {
    param([ValidateScript({ $true })][string]$Target)
}
$cmd = Get-Command Script-Target
$param = $cmd.Parameters["Target"]
$attr = $param.Attributes | Where-Object { $_.GetType().Name -eq "ValidateScriptAttribute" }
if ($attr -eq $null -or $attr.ScriptBlock -eq $null) {
    Write-Host "FAIL: ValidateScript reflection inspection failed"
    exit 1
}
Write-Host "PASS"
exit 0
