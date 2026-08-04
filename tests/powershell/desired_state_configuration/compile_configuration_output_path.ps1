# vybe-test: powershell/desired_state_configuration/compile_configuration_output_path
configuration OutputPathConfig {
    Node 'localhost' {
    }
}
OutputPathConfig -OutputPath "$PWD/dsc-output"
if (-not (Test-Path "$PWD/dsc-output/localhost.mof")) {
    Write-Host "FAIL: expected output path MOF"
    exit 1
}
Remove-Item -Recurse -Force "$PWD/dsc-output"
Write-Host 'PASS'
exit 0
