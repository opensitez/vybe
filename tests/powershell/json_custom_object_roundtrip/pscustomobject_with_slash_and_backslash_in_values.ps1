# vybe-test: powershell/json_custom_object_roundtrip/pscustomobject_with_slash_and_backslash_in_values
$orig = [pscustomobject]@{
    PathUnix = "/var/log/syslog"
    PathWin = "C:\Windows\System32"
}
$json = $orig | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.PathUnix -ne "/var/log/syslog" -or $recovered.PathWin -ne "C:\Windows\System32") {
    Write-Host "FAIL: Paths with slashes roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
