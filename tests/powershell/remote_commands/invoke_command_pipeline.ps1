# vybe-test: powershell/remote_commands/invoke_command_pipeline
$result = Invoke-Command -ScriptBlock { 1,2,3 | Measure-Object | Select-Object -ExpandProperty Count }
if ($result -ne 3) {
    Write-Host "FAIL: expected count 3"
    exit 1
}
Write-Host "PASS"
exit 0
