# vybe-test: powershell/desired_state_configuration/compile_configuration_with_file_resource
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
