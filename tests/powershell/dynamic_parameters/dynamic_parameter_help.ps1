# vybe-test: powershell/dynamic_parameters/dynamic_parameter_help
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
Get-Help Test-Func | Out-Null
Write-Host 'PASS'
exit 0
