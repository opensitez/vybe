# vybe-test: powershell/dynamic_parameter_runtime_generation/dynamic_param_dictionary_instantiation_and_lookup
# A RuntimeDefinedParameterDictionary is constructed and supports adding and looking up RuntimeDefinedParameter objects
$attributes = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
$attributes.Add([System.Management.Automation.ParameterAttribute]::new())

$param = [System.Management.Automation.RuntimeDefinedParameter]::new("ServerHost", [string], $attributes)
$dict = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
$dict.Add("ServerHost", $param)

if ($dict.Count -ne 1) {
    Write-Host "FAIL: expected dictionary count 1, got $($dict.Count)"
    exit 1
}

if (-not $dict.ContainsKey("ServerHost")) {
    Write-Host "FAIL: dictionary does not contain key 'ServerHost'"
    exit 1
}

$retrieved = $dict["ServerHost"]
if ($retrieved.Name -ne "ServerHost" -or $retrieved.ParameterType -ne [string]) {
    Write-Host "FAIL: retrieved parameter metadata mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
