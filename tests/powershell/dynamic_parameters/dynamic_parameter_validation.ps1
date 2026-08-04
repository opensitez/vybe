# vybe-test: powershell/dynamic_parameters/dynamic_parameter_validation
function Test-Func {
    param()
    dynamicparam {
        $attr = New-Object System.Management.Automation.ParameterAttribute
        $attr.ValidateRange(1,5)
        $param = New-Object System.Management.Automation.RuntimeDefinedParameter('X',[int],$attr)
        $dict = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $dict.Add('X',$param)
        return $dict
    }
    process { $X }
}
if ((Test-Func -X 4) -ne 4) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
