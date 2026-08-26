# vybe-test: powershell/parameters_validate_count/validatecount_reflection_attribute_inspection
function Count-Target {
    param([ValidateCount(2, 8)][string[]]$List)
}
$cmd = Get-Command Count-Target
$param = $cmd.Parameters["List"]
$attr = $param.Attributes | Where-Object { $_.GetType().Name -eq "ValidateCountAttribute" }
if ($attr -eq $null -or $attr.MinLength -ne 2 -or $attr.MaxLength -ne 8) {
    Write-Host "FAIL: ValidateCount reflection inspection failed"
    exit 1
}
Write-Host "PASS"
exit 0
