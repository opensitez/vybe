# vybe-test: powershell/dsc_configuration/dsc_configuration_test_command
$configData = @{
    AllNodes = @(
        @{
            NodeName = "localhost"
            Role = "WebServer"
            Port = 8080
        }
    )
}
if ($configData.AllNodes[0].NodeName -ne "localhost" -or $configData.AllNodes[0].Port -ne 8080) {
    Write-Host "FAIL: Configuration data check failed"
    exit 1
}
Write-Host "PASS"
exit 0
