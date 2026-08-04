# vybe-test: powershell/desired_state_configuration/compile_configuration_duplicate_nodes
configuration DuplicateConfig {
    Node 'localhost','localhost' {
    }
}
DuplicateConfig -OutputPath "$PWD/dsc-dup"
if (-not (Test-Path "$PWD/dsc-dup/localhost.mof")) {
    Write-Host "FAIL: expected duplicate node MOF"
    exit 1
}
Remove-Item -Recurse -Force "$PWD/dsc-dup"
Write-Host 'PASS'
exit 0
