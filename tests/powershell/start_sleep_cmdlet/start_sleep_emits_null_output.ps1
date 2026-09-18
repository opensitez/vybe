# vybe-test: powershell/start_sleep_cmdlet/start_sleep_emits_null_output
# Start-Sleep emits nothing ($null) to the pipeline upon completion
$output = Start-Sleep -Milliseconds 1

if ($null -ne $output) {
    Write-Host "FAIL: expected `$null output from Start-Sleep, got: $output"
    exit 1
}

Write-Host "PASS"
exit 0
