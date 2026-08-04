# vybe-test: powershell/dynamic_parameters/dynamic_parameter_switch
function Test-Func {
    param()
    dynamicparam {
        $attr = New-Object System.Management.Automation.ParameterAttribute
        $attr.ParameterType = [switch]
        $param = New-Object System.Management.Automation.RuntimeDefinedParameter('X',[switch],$attr)
        $dict = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $dict.Add('X',$param)
        return $dict
    }
    process { $X }
}
if ((Test-Func -X) -ne $true) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
