# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_validate_set_attribute
# ValidateSetAttribute applied to a dynamic parameter restricts valid input values to the permitted set
function ValidateEnvironmentParam {
    [CmdletBinding()]
    param()
    DynamicParam {
        $attrs = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
        $attrs.Add([System.Management.Automation.ParameterAttribute]::new())
        $attrs.Add([System.Management.Automation.ValidateSetAttribute]::new("Development", "Staging", "Production"))
        $dp = [System.Management.Automation.RuntimeDefinedParameter]::new("Env", [string], $attrs)
        $dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
        $dict.Add("Env", $dp)
        return $dict
    }
    process {
        return $PSBoundParameters["Env"]
    }
}

# Valid entry
$validVal = ValidateEnvironmentParam -Env "Staging"
if ($validVal -ne "Staging") {
    Write-Host "FAIL: valid ValidateSet value failed"
    exit 1
}

# Invalid entry should throw validation error
$threw = $false
try {
    ValidateEnvironmentParam -Env "InvalidEnv" -ErrorAction Stop
} catch {
    $threw = $true
}

if (-not $threw) {
    Write-Host "FAIL: invalid ValidateSet value did not throw validation exception"
    exit 1
}

Write-Host "PASS"
exit 0
