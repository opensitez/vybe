# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_positional_binding
# Dynamic parameters with ParameterAttribute.Position specified bind positional arguments correctly
function InvokePositionalDynamic {
    [CmdletBinding()]
    param()
    DynamicParam {
        $attrs = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
        $pAttr = [System.Management.Automation.ParameterAttribute]::new()
        $pAttr.Position = 0
        $attrs.Add($pAttr)
        $dp = [System.Management.Automation.RuntimeDefinedParameter]::new("PositionalName", [string], $attrs)
        $dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
        $dict.Add("PositionalName", $dp)
        return $dict
    }
    process {
        return $PSBoundParameters["PositionalName"]
    }
}

$boundPositional = InvokePositionalDynamic "PositionalArgumentValue"

if ($boundPositional -ne "PositionalArgumentValue") {
    Write-Host "FAIL: positional dynamic parameter binding failed, got: '$boundPositional'"
    exit 1
}

Write-Host "PASS"
exit 0
