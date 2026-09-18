# vybe-test: powershell/start_sleep_cmdlet/start_sleep_milliseconds_parameter_type_int32
# The Milliseconds parameter of Start-Sleep is typed as System.Int32
$paramDef = (Get-Command Start-Sleep).Parameters["Milliseconds"]

if ($null -eq $paramDef) {
    Write-Host "FAIL: Milliseconds parameter not found on Start-Sleep cmdlet"
    exit 1
}

if ($paramDef.ParameterType.FullName -ne "System.Int32") {
    Write-Host "FAIL: expected ParameterType System.Int32, got: '$($paramDef.ParameterType.FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
