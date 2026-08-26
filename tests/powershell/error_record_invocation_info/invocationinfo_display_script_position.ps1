# vybe-test: powershell/error_record_invocation_info/invocationinfo_display_script_position
$err = $null
try {
    $val = 10 / 0
} catch {
    $err = $_
}
$disp = $err.InvocationInfo.DisplayScriptPosition
if ($disp -ne $null -and $disp.Length -eq 0) {
    Write-Host "FAIL: DisplayScriptPosition empty check failed"
    exit 1
}
Write-Host "PASS"
exit 0
