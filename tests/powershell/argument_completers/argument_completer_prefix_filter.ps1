# vybe-test: powershell/argument_completers/argument_completer_prefix_filter
$candidates = @("Get-Process", "Get-Service", "Stop-Process")
$filter = {
    param($w)
    $script:candidates | Where-Object { $_ -like "$w*" }
}
$res = @(&$filter "Get")
if ($res.Count -ne 2 -or $res[0] -ne "Get-Process" -or $res[1] -ne "Get-Service") {
    Write-Host "FAIL: prefix filter expected Get-Process, Get-Service, got $($res -join ',')"
    exit 1
}
Write-Host "PASS"
exit 0
