# vybe-test: powershell/desired_state_configuration/compile_configuration_with_pscustomobject
$node = 'localhost'
configuration PSCustomConfig {
    Node $node {
    }
}
PSCustomConfig -OutputPath "$PWD/dsc-psobj"
if (-not (Test-Path "$PWD/dsc-psobj/localhost.mof")) {
    Write-Host "FAIL: expected pscustom object MOF"
    exit 1
}
Remove-Item -Recurse -Force "$PWD/dsc-psobj"
Write-Host 'PASS'
exit 0
