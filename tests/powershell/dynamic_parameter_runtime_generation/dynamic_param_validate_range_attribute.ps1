# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_validate_range_attribute
# ValidateRangeAttribute applied to a dynamic parameter enforces numeric lower and upper boundaries
function ValidateNumericRange {
    [CmdletBinding()]
    param()
    DynamicParam {
        $attrs = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
        $attrs.Add([System.Management.Automation.ParameterAttribute]::new())
        $attrs.Add([System.Management.Automation.ValidateRangeAttribute]::new(1, 100))
        $dp = [System.Management.Automation.RuntimeDefinedParameter]::new("Percentage", [int], $attrs)
        $dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
        $dict.Add("Percentage", $dp)
        return $dict
    }
    process {
        return $PSBoundParameters["Percentage"]
    }
}

$ok = ValidateNumericRange -Percentage 75
if ($ok -ne 75) {
    Write-Host "FAIL: valid range input failed"
    exit 1
}

$threw = $false
try {
    ValidateNumericRange -Percentage 150 -ErrorAction Stop
} catch {
    $threw = $true
}

if (-not $threw) {
    Write-Host "FAIL: out of range value did not throw exception"
    exit 1
}

Write-Host "PASS"
exit 0
