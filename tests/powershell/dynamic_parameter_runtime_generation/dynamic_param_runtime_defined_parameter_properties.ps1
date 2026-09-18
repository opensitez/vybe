# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_runtime_defined_parameter_properties
# RuntimeDefinedParameter exposes Name, ParameterType, Attributes, Value, and IsSet properties
$attributes = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
$paramAttr = [System.Management.Automation.ParameterAttribute]::new()
$paramAttr.Mandatory = $true
$attributes.Add($paramAttr)

$param = [System.Management.Automation.RuntimeDefinedParameter]::new("MaxRetryCount", [int], $attributes)

if ($param.Name -ne "MaxRetryCount") {
    Write-Host "FAIL: expected Name 'MaxRetryCount', got '$($param.Name)'"
    exit 1
}

if ($param.ParameterType -ne [int]) {
    Write-Host "FAIL: expected ParameterType [int], got $($param.ParameterType)"
    exit 1
}

if ($param.IsSet) {
    Write-Host "FAIL: IsSet should initially be false before value assignment"
    exit 1
}

$param.Value = 5
if (-not $param.IsSet -or $param.Value -ne 5) {
    Write-Host "FAIL: IsSet not updated or value mismatch after assignment"
    exit 1
}

Write-Host "PASS"
exit 0
