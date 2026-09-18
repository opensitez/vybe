# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_type_coercion_failure_throws
# Passing a value that cannot be converted to the dynamic parameter's declared type throws a binding exception
function RequireIntegerPort {
    [CmdletBinding()]
    param()
    DynamicParam {
        $attrs = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
        $attrs.Add([System.Management.Automation.ParameterAttribute]::new())
        $dp = [System.Management.Automation.RuntimeDefinedParameter]::new("Port", [int], $attrs)
        $dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
        $dict.Add("Port", $dp)
        return $dict
    }
    process {
        return $PSBoundParameters["Port"]
    }
}

$threw = $false
try {
    RequireIntegerPort -Port "NotANumber" -ErrorAction Stop
} catch {
    $threw = $true
}

if (-not $threw) {
    Write-Host "FAIL: non-numeric string passed to [int] dynamic param did not throw"
    exit 1
}

Write-Host "PASS"
exit 0
