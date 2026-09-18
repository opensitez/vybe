# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_value_from_pipeline_binding
# Dynamic parameters with ParameterAttribute.ValueFromPipeline=$true successfully bind values streamed from pipeline
function ProcessPipelineDynamic {
    [CmdletBinding()]
    param()
    DynamicParam {
        $attrs = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
        $pAttr = [System.Management.Automation.ParameterAttribute]::new()
        $pAttr.ValueFromPipeline = $true
        $attrs.Add($pAttr)
        $dp = [System.Management.Automation.RuntimeDefinedParameter]::new("InputPayload", [int], $attrs)
        $dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
        $dict.Add("InputPayload", $dp)
        return $dict
    }
    process {
        $PSBoundParameters["InputPayload"] * 2
    }
}

$streamedResults = @(10, 20, 30) | ProcessPipelineDynamic

if ($streamedResults.Count -ne 3) {
    Write-Host "FAIL: expected 3 results, got $($streamedResults.Count)"
    exit 1
}

$expected = @(20, 40, 60)
for ($i = 0; $i -lt 3; $i++) {
    if ($streamedResults[$i] -ne $expected[$i]) {
        Write-Host "FAIL: at index ${i}, expected $($expected[$i]), got $($streamedResults[$i])"
        exit 1
    }
}

Write-Host "PASS"
exit 0
