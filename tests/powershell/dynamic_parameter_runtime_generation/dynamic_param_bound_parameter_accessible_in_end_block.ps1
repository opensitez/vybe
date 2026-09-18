# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_bound_parameter_accessible_in_end_block
# Dynamic parameter values remain accessible via $PSBoundParameters in the function end block
function AggregateDynamicStats {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    DynamicParam {
        $attrs = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
        $attrs.Add([System.Management.Automation.ParameterAttribute]::new())
        $dp = [System.Management.Automation.RuntimeDefinedParameter]::new("LabelPrefix", [string], $attrs)
        $dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
        $dict.Add("LabelPrefix", $dp)
        return $dict
    }
    begin {
        $total = 0
    }
    process {
        $total += $InputObject
    }
    end {
        return "$($PSBoundParameters['LabelPrefix']):$total"
    }
}

$summary = 10..13 | AggregateDynamicStats -LabelPrefix "TotalSum"

# 10 + 11 + 12 + 13 = 46
if ($summary -ne "TotalSum:46") {
    Write-Host "FAIL: end block access to dynamic parameter failed, got: '$summary'"
    exit 1
}

Write-Host "PASS"
exit 0
