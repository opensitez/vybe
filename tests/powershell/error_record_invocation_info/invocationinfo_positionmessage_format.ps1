# vybe-test: powershell/error_record_invocation_info/invocationinfo_positionmessage_format
$err = $null
try {
    throw "PositionCheck"
} catch {
    $err = $_
}
$msg = $err.InvocationInfo.PositionMessage
if (-not ($msg.Contains("line") -or $msg.Contains("char") -or $msg.Contains("+"))) {
    Write-Host "FAIL: PositionMessage format check failed, got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
