# vybe-test: powershell/dynamic_parameters/dynamic_parameter_help_content
function Test-Func {
    param()
    dynamicparam {
        $attr = New-Object System.Management.Automation.ParameterAttribute
        $attr.HelpMessage = 'hi'
        $param = New-Object System.Management.Automation.RuntimeDefinedParameter('X',[int],$attr)
        $dict = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $dict.Add('X',$param)
        return $dict
    }
    process { $X }
}
Get-Help Test-Func | Out-Null
Write-Host 'PASS'
exit 0
