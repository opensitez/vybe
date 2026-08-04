# vybe-test: powershell/desired_state_configuration/compile_empty_configuration
configuration TestConfig {
    Node 'localhost' {
    }
}
TestConfig -OutputPath "$PWD/dsc-empty"
if (-not (Test-Path "$PWD/dsc-empty/localhost.mof")) {
    Write-Host "FAIL: expected MOF file"
    exit 1
}
Remove-Item -Recurse -Force "$PWD/dsc-empty"
Write-Host 'PASS'
exit 0
