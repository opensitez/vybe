# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_switch_type_present
# A dynamic parameter typed as [switch] correctly binds as a boolean flag when provided
function TestSwitchParam {
    [CmdletBinding()]
    param()
    DynamicParam {
        $attrs = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
        $attrs.Add([System.Management.Automation.ParameterAttribute]::new())
        $dp = [System.Management.Automation.RuntimeDefinedParameter]::new("VerboseAudit", [switch], $attrs)
        $dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
        $dict.Add("VerboseAudit", $dp)
        return $dict
    }
    process {
        return $PSBoundParameters["VerboseAudit"].IsPresent
    }
}

$switchResult = TestSwitchParam -VerboseAudit

if ($switchResult -ne $true) {
    Write-Host "FAIL: switch dynamic parameter was not present as true"
    exit 1
}

Write-Host "PASS"
exit 0
