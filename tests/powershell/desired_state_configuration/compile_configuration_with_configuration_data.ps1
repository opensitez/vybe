# vybe-test: powershell/desired_state_configuration/compile_configuration_with_configuration_data
$configurationData = @{ AllNodes = @(@{ NodeName = 'localhost'}) }
configuration DataConfig {
    Node $AllNodes.NodeName {
    }
}
DataConfig -OutputPath "$PWD/dsc-data" -ConfigurationData $configurationData
if (-not (Test-Path "$PWD/dsc-data/localhost.mof")) {
    Write-Host "FAIL: expected data MOF"
    exit 1
}
Remove-Item -Recurse -Force "$PWD/dsc-data"
Write-Host 'PASS'
exit 0
