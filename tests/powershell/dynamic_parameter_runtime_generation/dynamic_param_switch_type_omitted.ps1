# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_switch_type_omitted
# An omitted dynamic switch parameter is absent from $PSBoundParameters and evaluates false
function TestOmittedSwitch {
    [CmdletBinding()]
    param()
    DynamicParam {
        $attrs = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
        $attrs.Add([System.Management.Automation.ParameterAttribute]::new())
        $dp = [System.Management.Automation.RuntimeDefinedParameter]::new("IncludeArchived", [switch], $attrs)
        $dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
        $dict.Add("IncludeArchived", $dp)
        return $dict
    }
    process {
        return $PSBoundParameters.ContainsKey("IncludeArchived")
    }
}

$isBound = TestOmittedSwitch

if ($isBound) {
    Write-Host "FAIL: omitted dynamic switch parameter was unexpectedly bound in `$PSBoundParameters"
    exit 1
}

Write-Host "PASS"
exit 0
