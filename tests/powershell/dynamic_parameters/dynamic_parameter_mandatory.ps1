# vybe-test: powershell/dynamic_parameters/dynamic_parameter_mandatory
function Test-Func {
    param()
    dynamicparam {
        $attr = New-Object System.Management.Automation.ParameterAttribute
        $attr.Mandatory = $true
        $param = New-Object System.Management.Automation.RuntimeDefinedParameter('X',[int],$attr)
        $dict = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $dict.Add('X',$param)
        return $dict
    }
    process { $X }
}
if ((Test-Func -X 3) -ne 3) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
