# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_conditional_creation_based_on_static_param
# Dynamic parameters are dynamically generated only when a preceding static parameter satisfies conditions
function TestConditionalParam {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProviderType
    )
    DynamicParam {
        if ($ProviderType -eq "Cloud") {
            $attrs = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
            $attrs.Add([System.Management.Automation.ParameterAttribute]::new())
            $dp = [System.Management.Automation.RuntimeDefinedParameter]::new("Region", [string], $attrs)
            $dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
            $dict.Add("Region", $dp)
            return $dict
        }
    }
    process {
        $hasRegion = $PSBoundParameters.ContainsKey("Region")
        $regionVal = if ($hasRegion) { $PSBoundParameters["Region"] } else { "none" }
        return "Provider=${ProviderType}, Region=${regionVal}"
    }
}

$cloudResult = TestConditionalParam -ProviderType Cloud -Region "us-west-2"
$localResult = TestConditionalParam -ProviderType Local

if ($cloudResult -ne "Provider=Cloud, Region=us-west-2") {
    Write-Host "FAIL: cloud invocation failed: '$cloudResult'"
    exit 1
}

if ($localResult -ne "Provider=Local, Region=none") {
    Write-Host "FAIL: local invocation failed: '$localResult'"
    exit 1
}

Write-Host "PASS"
exit 0
