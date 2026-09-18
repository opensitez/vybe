# vybe-test: powershell/test_path_cmdlet/test_path_pipeline_streaming
$paths = @("tests", "completely_fake_dir_9988")

# Piping paths into Test-Path emits a boolean for each streamed path
$results = @($paths | Test-Path)

if ($results.Count -ne 2) {
    Write-Host "FAIL: expected 2 boolean results, got $($results.Count)"
    exit 1
}

if ($results[0] -ne $true -or $results[1] -ne $false) {
    Write-Host "FAIL: expected @(`$true, `$false), got: @($($results -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
