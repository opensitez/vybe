# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_bound_parameter_accessible_in_process_block
# Dynamic parameter values are accessible via $PSBoundParameters within each process block execution
function ApplyDynamicTransform {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    DynamicParam {
        $attrs = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
        $attrs.Add([System.Management.Automation.ParameterAttribute]::new())
        $dp = [System.Management.Automation.RuntimeDefinedParameter]::new("Multiplier", [int], $attrs)
        $dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
        $dict.Add("Multiplier", $dp)
        return $dict
    }
    process {
        $InputObject * $PSBoundParameters["Multiplier"]
    }
}

$results = @(1, 2, 3) | ApplyDynamicTransform -Multiplier 10

if ($results.Count -ne 3) {
    Write-Host "FAIL: expected 3 results, got $($results.Count)"
    exit 1
}

$expected = "10, 20, 30"
$actual = $results -join ", "
if ($actual -ne $expected) {
    Write-Host "FAIL: expected '$expected', got '$actual'"
    exit 1
}

Write-Host "PASS"
exit 0
