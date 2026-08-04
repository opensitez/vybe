# vybe-test: powershell/writer_streams/debug_writer
$DebugPreference = 'Continue'
Write-Debug 'dbg'
if ($DebugPreference -ne 'Continue') {
    Write-Host "FAIL: expected debug continue"
    exit 1
}
Write-Host 'PASS'
exit 0
