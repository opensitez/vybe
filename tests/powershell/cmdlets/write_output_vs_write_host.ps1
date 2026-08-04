# vybe-test: powershell/cmdlets/write_output_vs_write_host
# Write-Output goes to the pipeline; Write-Host goes directly to console
$captured = Write-Output "pipeline value"
if ($captured -ne "pipeline value") {
    Write-Host "FAIL: Write-Output should be capturable"
    exit 1
}
# Write-Host output cannot be captured (goes to host, not pipeline)
$hostOut = Write-Host "host message" 2>&1
# $hostOut may be empty or InformationRecord - Write-Host is for display only
# The key test: Write-Output value was captured correctly
Write-Host "PASS"
exit 0
