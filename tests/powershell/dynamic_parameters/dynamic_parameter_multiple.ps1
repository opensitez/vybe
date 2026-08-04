# vybe-test: powershell/dynamic_parameters/dynamic_parameter_multiple
function Test-Func {
    param()
    dynamicparam {
        $dict = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $param1 = New-Object System.Management.Automation.RuntimeDefinedParameter('X',[int],$null)
        $param2 = New-Object System.Management.Automation.RuntimeDefinedParameter('Y',[int],$null)
        $param1.Attributes.Add((New-Object System.Management.Automation.ParameterAttribute))
        $param2.Attributes.Add((New-Object System.Management.Automation.ParameterAttribute))
        $dict.Add('X',$param1)
        $dict.Add('Y',$param2)
        return $dict
    }
    process { $X + $Y }
}
if ((Test-Func -X 2 -Y 3) -ne 5) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
