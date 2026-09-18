# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_runtime_defined_parameter_is_set_property
# RuntimeDefinedParameter.IsSet evaluates to true once a value is assigned and false otherwise
$attributes = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
$attributes.Add([System.Management.Automation.ParameterAttribute]::new())

$rdp = [System.Management.Automation.RuntimeDefinedParameter]::new("SampleParam", [string], $attributes)

if ($rdp.IsSet) {
    Write-Host "FAIL: IsSet was true before assigning any value"
    exit 1
}

$rdp.Value = "assigned_content"

if (-not $rdp.IsSet) {
    Write-Host "FAIL: IsSet was not true after assigning value"
    exit 1
}

Write-Host "PASS"
exit 0
