# vybe-test: powershell/desired_state_configuration/compile_configuration_node
configuration NodeConfig {
    Node 'localhost' {
    }
}
NodeConfig -OutputPath "$PWD/dsc-node"
if (-not (Test-Path "$PWD/dsc-node/localhost.mof")) {
    Write-Host "FAIL: expected node MOF"
    exit 1
}
Remove-Item -Recurse -Force "$PWD/dsc-node"
Write-Host 'PASS'
exit 0
