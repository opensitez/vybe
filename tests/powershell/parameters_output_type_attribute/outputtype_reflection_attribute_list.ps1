# vybe-test: powershell/parameters_output_type_attribute/outputtype_reflection_attribute_list
function Inspect-OutputTypeTarget {
    [OutputType([string])]
    param()
}
$cmd = Get-Command Inspect-OutputTypeTarget
$attrs = @($cmd.ScriptBlock.Attributes | Where-Object { $_.GetType().Name -eq "OutputTypeAttribute" })
if ($attrs.Count -ne 1) {
    Write-Host "FAIL: ScriptBlock Attributes OutputTypeAttribute reflection failed"
    exit 1
}
Write-Host "PASS"
exit 0
