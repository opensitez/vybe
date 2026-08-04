# vybe-test: powershell/dynamic_parameters/dynamic_parameter_type
function Test-Func {
    param()
    dynamicparam {
        $param = New-Object System.Management.Automation.RuntimeDefinedParameter('X',[int],$null)
        $param.Attributes.Add((New-Object System.Management.Automation.ParameterAttribute))
        $dict = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $dict.Add('X',$param)
        return $dict
    }
    process { $X }
}
if ((Test-Func -X 2) -ne 2) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
