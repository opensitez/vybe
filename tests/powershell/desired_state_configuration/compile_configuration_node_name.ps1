# vybe-test: powershell/desired_state_configuration/compile_configuration_node_name
$node = 'localhost'
configuration NameConfig {
    Node $node {
    }
}
NameConfig -OutputPath "$PWD/dsc-name"
if (-not (Test-Path "$PWD/dsc-name/localhost.mof")) {
    Write-Host "FAIL: expected named node MOF"
    exit 1
}
Remove-Item -Recurse -Force "$PWD/dsc-name"
Write-Host 'PASS'
exit 0
