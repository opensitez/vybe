# vybe-test: powershell/get_unique_cmdlet/get_unique_pipeline_streaming
$words = @("echo", "echo", "foxtrot", "golf", "golf")
$unique = @($words | Get-Unique)

if ($unique.Count -ne 3) {
    Write-Host "FAIL: expected 3 unique items from stream, got $($unique.Count)"
    exit 1
}

if ($unique[0] -ne "echo" -or $unique[1] -ne "foxtrot" -or $unique[2] -ne "golf") {
    Write-Host "FAIL: pipeline stream mismatch: @($($unique -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
