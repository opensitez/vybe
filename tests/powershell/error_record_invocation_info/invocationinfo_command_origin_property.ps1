# vybe-test: powershell/error_record_invocation_info/invocationinfo_command_origin_property
$err = $null
try { throw "OriginErr" } catch { $err = $_ }
$origin = $err.InvocationInfo.CommandOrigin
if ($origin -ne [System.Management.Automation.CommandOrigin]::Internal -and $origin -ne [System.Management.Automation.CommandOrigin]::Runspace) {
    # Valid CommandOrigin enum
    if ($origin -eq $null) {
        Write-Host "FAIL: CommandOrigin was null"
        exit 1
    }
}
Write-Host "PASS"
exit 0
