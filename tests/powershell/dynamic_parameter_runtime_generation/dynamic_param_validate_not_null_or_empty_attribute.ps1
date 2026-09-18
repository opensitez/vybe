# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_validate_not_null_or_empty_attribute
# ValidateNotNullOrEmptyAttribute prevents passing null or empty strings to the dynamic parameter
function RequireNonEmptyString {
    [CmdletBinding()]
    param()
    DynamicParam {
        $attrs = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
        $attrs.Add([System.Management.Automation.ParameterAttribute]::new())
        $attrs.Add([System.Management.Automation.ValidateNotNullOrEmptyAttribute]::new())
        $dp = [System.Management.Automation.RuntimeDefinedParameter]::new("Identifier", [string], $attrs)
        $dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
        $dict.Add("Identifier", $dp)
        return $dict
    }
    process {
        return $PSBoundParameters["Identifier"]
    }
}

$validCall = RequireNonEmptyString -Identifier "ID-900"
if ($validCall -ne "ID-900") {
    Write-Host "FAIL: valid identifier failed"
    exit 1
}

$threwOnEmpty = $false
try {
    RequireNonEmptyString -Identifier "" -ErrorAction Stop
} catch {
    $threwOnEmpty = $true
}

if (-not $threwOnEmpty) {
    Write-Host "FAIL: empty string did not throw ValidateNotNullOrEmpty exception"
    exit 1
}

Write-Host "PASS"
exit 0
