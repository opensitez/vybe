# vybe-test: powershell/desired_state_configuration/compile_configuration_invoke
configuration InvokeConfig {
    Node 'localhost' {
    }
}
InvokeConfig -OutputPath "$PWD/dsc-invoke"
if (-not (Test-Path "$PWD/dsc-invoke/localhost.mof")) {
    Write-Host "FAIL: expected invoke MOF"
    exit 1
}
Remove-Item -Recurse -Force "$PWD/dsc-invoke"
Write-Host 'PASS'
exit 0
