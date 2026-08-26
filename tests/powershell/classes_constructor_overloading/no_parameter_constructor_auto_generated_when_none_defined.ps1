# vybe-test: powershell/classes_constructor_overloading/no_parameter_constructor_auto_generated_when_none_defined
class AutoConstruct {
    [int]$Num = 10
    [string]$Text = "Default"
}
$ac = [AutoConstruct]::new()
if ($ac.Num -ne 10 -or $ac.Text -ne "Default") {
    Write-Host "FAIL: Auto-generated parameterless constructor failed"
    exit 1
}
Write-Host "PASS"
exit 0
