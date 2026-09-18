# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_type_coercion_to_integer
# A dynamic parameter typed as [int] automatically coerces string argument inputs into System.Int32
function InvokeScaleOperation {
    [CmdletBinding()]
    param()
    DynamicParam {
        $attrs = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
        $attrs.Add([System.Management.Automation.ParameterAttribute]::new())
        $dp = [System.Management.Automation.RuntimeDefinedParameter]::new("Factor", [int], $attrs)
        $dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
        $dict.Add("Factor", $dp)
        return $dict
    }
    process {
        return $PSBoundParameters["Factor"]
    }
}

$result = InvokeScaleOperation -Factor "42"

if ($result -ne 42) {
    Write-Host "FAIL: expected numeric value 42, got: $result"
    exit 1
}

if ($result.GetType().Name -ne "Int32") {
    Write-Host "FAIL: expected coerced type Int32, got $($result.GetType().Name)"
    exit 1
}

Write-Host "PASS"
exit 0
