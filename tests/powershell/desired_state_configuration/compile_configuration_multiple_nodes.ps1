# vybe-test: powershell/desired_state_configuration/compile_configuration_multiple_nodes
configuration MultiConfig {
    Node 'localhost','127.0.0.1' {
    }
}
MultiConfig -OutputPath "$PWD/dsc-multi"
if (-not (Test-Path "$PWD/dsc-multi/localhost.mof")) {
    Write-Host "FAIL: expected one MOF file"
    exit 1
}
Remove-Item -Recurse -Force "$PWD/dsc-multi"
Write-Host 'PASS'
exit 0
