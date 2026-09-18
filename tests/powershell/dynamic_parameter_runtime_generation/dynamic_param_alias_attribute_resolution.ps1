# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_alias_attribute_resolution
# Adding an AliasAttribute to the dynamic parameter attribute collection resolves calls using that alias
function ResolveAliasParam {
    [CmdletBinding()]
    param()
    DynamicParam {
        $attrs = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
        $attrs.Add([System.Management.Automation.ParameterAttribute]::new())
        $attrs.Add([System.Management.Automation.AliasAttribute]::new("TargetHost"))
        $dp = [System.Management.Automation.RuntimeDefinedParameter]::new("Hostname", [string], $attrs)
        $dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
        $dict.Add("Hostname", $dp)
        return $dict
    }
    process {
        return $PSBoundParameters["Hostname"]
    }
}

$resultViaAlias = ResolveAliasParam -TargetHost "db.internal.corp"

if ($resultViaAlias -ne "db.internal.corp") {
    Write-Host "FAIL: dynamic parameter alias resolution failed, got: '$resultViaAlias'"
    exit 1
}

Write-Host "PASS"
exit 0
