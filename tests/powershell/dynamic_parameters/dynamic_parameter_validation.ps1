# vybe-test: powershell/dynamic_parameters/dynamic_parameter_validation
function Test-DynParam {
    [CmdletBinding()]
    param([string]$Type = "Default")
    DynamicParam {
        $dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
        return $dict
    }
    process {
        return $Type
    }
}
$res = Test-DynParam -Type "DynamicTarget"
if ($res -ne "DynamicTarget") {
    Write-Host "FAIL: DynamicParam function failed"
    exit 1
}
Write-Host "PASS"
exit 0
